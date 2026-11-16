#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Exports every user attribute a dynamic membership rule can target, for the members of one
    or more Entra groups, plus a summary of which attributes are worth targeting.

.DESCRIPTION
    Reads the members of the group(s) you name and writes two CSV files:

      ..._Users.csv     One row per user. Column names are exactly what you type after "user."
                        in a dynamic membership rule (department, jobTitle, extensionAttribute3,
                        extension_<appId>_<name>, ...). Columns starting with "ref." can NOT be
                        used in rules. They are included so you can match rows to your AD export
                        (ref.onPremisesSamAccountName = sAMAccountName) and see synced data that
                        rules can't reach (employeeType, costCenter, division).

      ..._Summary.csv   One row per column: whether rules can use it, the Graph property it was
                        read from, the AD attribute that feeds it under the default Entra Connect
                        sync rules, how many users have a value, how many distinct values exist,
                        and the most common values. Start here when choosing what to target.

    The attribute list follows Microsoft's supported properties for user rules:
    https://learn.microsoft.com/entra/identity/users/groups-dynamic-membership
    Directory extension properties (extension_<appId>_<name>) are discovered automatically,
    including any that your sync creates from AD attributes.

    Some rule property names differ from the Graph property they're read from:
        Rule property               Graph property
        dirSyncEnabled              onPremisesSyncEnabled
        facsimileTelephoneNumber    faxNumber
        mobile                      mobilePhone
        objectId                    id
        physicalDeliveryOfficeName  officeLocation
        telephoneNumber             businessPhones
        sipProxyAddress             imAddresses (closest Graph equivalent)
        extensionAttribute1-15      onPremisesExtensionAttributes

    Read-only: the script changes nothing.

.PARAMETER Group
    One or more group display names or object IDs. If omitted, you're prompted.

.PARAMETER IncludeNested
    Include users from nested groups (transitive membership). By default only direct members.

.PARAMETER IncludeMemberOf
    Add a memberOf column with each user's direct group memberships. Rules can target these
    with user.memberOf (preview). Costs one extra Graph call per user.

.PARAMETER IncludeManager
    Add a ref.manager column. It isn't a rule property, but the "Direct Reports" rule is
    built on it. Costs one extra Graph call per user.

.PARAMETER OutputFolder
    Where to write the CSVs. Defaults to the current directory.

.EXAMPLE
    .\Get-EntraUserAttributes.ps1 -Group "SG-Finance"

.EXAMPLE
    .\Get-EntraUserAttributes.ps1 -Group "SG-Finance", "SG-IT-Admins" -IncludeNested -IncludeMemberOf

.NOTES
    Required module:  Microsoft.Graph.Authentication
    Graph scope:      Directory.Read.All (delegated). Listing directory extension definitions
                      requires it, and it also covers reading users and group members.
                      Listing extension definitions also needs an Entra role such as
                      Application Administrator or Cloud Application Administrator. Without
                      one, the script continues without extension_ columns and says so.

    The members query uses Graph advanced queries (ConsistencyLevel: eventual), so changes
    made in the last few minutes might not show yet.
#>

[CmdletBinding()]
param(
    [string[]]$Group,
    [switch]$IncludeNested,
    [switch]$IncludeMemberOf,
    [switch]$IncludeManager,
    [string]$OutputFolder = (Get-Location).Path
)

$ErrorActionPreference = 'Stop'

# ======================== HELPERS ========================

function Invoke-GraphGetAll {
    # GETs a collection and follows @odata.nextLink. OData casts such as /microsoft.graph.user
    # need ConsistencyLevel: eventual (with $count=true in the URI), so the header is always sent.
    param([Parameter(Mandatory)][string]$Uri)
    $next = $Uri
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -Headers @{ ConsistencyLevel = 'eventual' }
        foreach ($item in $resp.value) { $item }
        $next = $resp.'@odata.nextLink'
    }
}

function Get-ValueAtPath {
    # Reads a dotted path (e.g. onPremisesExtensionAttributes.extensionAttribute1) from the
    # hashtables that Invoke-MgGraphRequest returns.
    param($Object, [string]$Path)
    $current = $Object
    foreach ($part in $Path.Split('.')) {
        if ($null -eq $current) { return $null }
        $current = $current[$part]
    }
    $current
}

function ConvertTo-CellValue {
    # Flattens a Graph value into one CSV cell. Booleans are lowercase to match rule syntax.
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [bool]) { return $Value.ToString().ToLower() }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
    if ($Value -is [datetimeoffset]) { return $Value.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ') }
    if ($Value -is [System.Collections.IDictionary]) { return ($Value | ConvertTo-Json -Compress -Depth 5) }
    if ($Value -is [System.Collections.IEnumerable]) {
        return (@($Value | ForEach-Object { ConvertTo-CellValue $_ }) | Where-Object { $_ -ne '' }) -join '; '
    }
    return [string]$Value
}

function Format-Short {
    param([string]$Text, [int]$Max = 60)
    if ($Text.Length -le $Max) { return $Text }
    return $Text.Substring(0, $Max - 3) + '...'
}

# ======================== GRAPH CONNECTION ========================
$requiredScope = 'Directory.Read.All'
$context = Get-MgContext
$hasScope = $context -and (@($context.Scopes) -contains $requiredScope -or @($context.Scopes) -contains 'Directory.ReadWrite.All')
if (-not $hasScope) {
    Write-Host "Connecting to Microsoft Graph ($requiredScope)..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes $requiredScope -NoWelcome
    $context = Get-MgContext
    if (-not $context) {
        Write-Host "Failed to connect to Microsoft Graph. Exiting." -ForegroundColor Red
        return
    }
}
Write-Host "Connected as: $($context.Account)" -ForegroundColor Green

# ======================== GROUP INPUT ========================
if (-not $Group) {
    Write-Host ""
    Write-Host "Enter group names or object IDs (one per line, blank line to finish):" -ForegroundColor Cyan
    $Group = @()
    while ($true) {
        $entry = Read-Host "  Group"
        if ([string]::IsNullOrWhiteSpace($entry)) { break }
        $Group += $entry.Trim()
    }
}

if (-not $Group) {
    Write-Host "No groups specified. Exiting." -ForegroundColor Yellow
    return
}

# ======================== RESOLVE GROUPS ========================
Write-Host ""
Write-Host "Resolving groups..." -ForegroundColor Cyan

$guidPattern = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
$resolvedGroups = [System.Collections.Generic.List[object]]::new()

foreach ($groupInput in $Group) {
    $found = @()
    if ($groupInput -match $guidPattern) {
        try {
            $found = @(Invoke-MgGraphRequest -Method GET -Uri "v1.0/groups/${groupInput}?`$select=id,displayName")
        }
        catch { $found = @() }
    }
    if ($found.Count -eq 0) {
        $filter = [uri]::EscapeDataString("displayName eq '$($groupInput -replace "'", "''")'")
        $found = @(Invoke-GraphGetAll -Uri "v1.0/groups?`$filter=$filter&`$select=id,displayName")
    }

    if ($found.Count -eq 1) {
        if ($resolvedGroups.id -contains $found[0].id) {
            Write-Host "  Skipped duplicate: $($found[0].displayName)" -ForegroundColor DarkGray
        }
        else {
            $resolvedGroups.Add($found[0])
            Write-Host "  Resolved: $($found[0].displayName) ($($found[0].id))" -ForegroundColor Green
        }
    }
    elseif ($found.Count -gt 1) {
        Write-Host "  AMBIGUOUS: '$groupInput' matches $($found.Count) groups. Use the object ID instead:" -ForegroundColor Red
        foreach ($f in $found) { Write-Host "    $($f.displayName)  $($f.id)" -ForegroundColor DarkGray }
    }
    else {
        Write-Host "  NOT FOUND: $groupInput" -ForegroundColor Red
    }
}

if ($resolvedGroups.Count -eq 0) {
    Write-Host "No valid groups found. Exiting." -ForegroundColor Red
    return
}

# ======================== DIRECTORY EXTENSIONS ========================
Write-Host ""
Write-Host "Looking up directory extension properties..." -ForegroundColor Cyan

$userExtensions = @()
$extensionWarning = $null
try {
    $resp = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/directoryObjects/getAvailableExtensionProperties' `
        -Body '{}' -ContentType 'application/json'
    $userExtensions = @($resp.value | Where-Object { @($_.targetObjects) -contains 'User' } | Sort-Object { $_.name })
    Write-Host "  Found $($userExtensions.Count) user extension property(ies)" -ForegroundColor Green
}
catch {
    $extensionWarning = "Couldn't list directory extension properties, so extension_<appId>_<name> columns are missing. " +
        "Your account probably needs an Entra role such as Application Administrator. Error: $($_.Exception.Message)"
    Write-Host "  $extensionWarning" -ForegroundColor Yellow
}

# ======================== COLUMN DEFINITIONS ========================
# Name  = CSV column; for rule-usable columns this is the rule property name (user.<Name>)
# Graph = Graph property it comes from     Path  = where to read it (underscore = derived below)
# Rules = usable in dynamic rules?         AD    = feeding AD attribute under default Entra Connect rules
$columns = [System.Collections.Generic.List[object]]::new()
function Add-Column {
    param(
        [Parameter(Mandatory, Position = 0)][string]$Name,
        [Parameter(Mandatory)][string]$Graph,
        [string]$Path = $Graph,
        [string]$Rules = 'Yes',
        [string]$AD = '',
        [switch]$Multi,
        [string]$Notes = ''
    )
    $columns.Add([pscustomobject]@{
            Name = $Name; Graph = $Graph; Path = $Path; Rules = $Rules; AD = $AD; Multi = [bool]$Multi; Notes = $Notes
        })
}

$anyNote = 'Multi-value: rules need -any/-all, e.g. (user.{0} -any (_ -endsWith "@contoso.com")).'

# Boolean
Add-Column accountEnabled -Graph accountEnabled -AD 'userAccountControl (disabled flag)'
Add-Column dirSyncEnabled -Graph onPremisesSyncEnabled -AD '(set by sync)' -Notes 'Blank = cloud-only user (never synced).'
# Date/time
Add-Column employeeHireDate -Graph employeeHireDate -Rules 'Yes (preview)' -Notes 'Supports -ge/-le and system.now.'
# String
Add-Column city -Graph city -AD 'l'
Add-Column country -Graph country -AD 'co'
Add-Column companyName -Graph companyName -AD 'company'
Add-Column department -Graph department -AD 'department'
Add-Column displayName -Graph displayName -AD 'displayName'
Add-Column employeeId -Graph employeeId -AD 'employeeID'
Add-Column facsimileTelephoneNumber -Graph faxNumber -AD 'facsimileTelephoneNumber'
Add-Column givenName -Graph givenName -AD 'givenName'
Add-Column jobTitle -Graph jobTitle -AD 'title'
Add-Column mail -Graph mail -AD 'mail'
Add-Column mailNickName -Graph mailNickname -AD 'mailNickname'
Add-Column mobile -Graph mobilePhone -AD 'mobile'
Add-Column objectId -Graph id -AD '(cloud object ID)'
Add-Column onPremisesDistinguishedName -Graph onPremisesDistinguishedName -AD 'distinguishedName' -Notes 'Full AD DN, so a rule can target an OU.'
Add-Column onPremisesSecurityIdentifier -Graph onPremisesSecurityIdentifier -AD 'objectSid'
Add-Column passwordPolicies -Graph passwordPolicies -AD '(cloud-side)'
Add-Column physicalDeliveryOfficeName -Graph officeLocation -AD 'physicalDeliveryOfficeName'
Add-Column postalCode -Graph postalCode -AD 'postalCode'
Add-Column preferredLanguage -Graph preferredLanguage -AD 'preferredLanguage'
Add-Column sipProxyAddress -Graph imAddresses -AD 'msRTCSIP-PrimaryUserAddress' -Notes 'Graph has no sipProxyAddress; showing imAddresses (SIP addresses), the closest equivalent.'
Add-Column state -Graph state -AD 'st'
Add-Column streetAddress -Graph streetAddress -AD 'streetAddress'
Add-Column surname -Graph surname -AD 'sn'
Add-Column telephoneNumber -Graph businessPhones -AD 'telephoneNumber'
Add-Column usageLocation -Graph usageLocation -AD 'msExchUsageLocation (if your sync rules flow it)'
Add-Column userPrincipalName -Graph userPrincipalName -AD 'userPrincipalName'
Add-Column userType -Graph userType -AD '(cloud-side)'
# String collections
Add-Column otherMails -Graph otherMails -Multi -Notes ($anyNote -f 'otherMails')
Add-Column proxyAddresses -Graph proxyAddresses -AD 'proxyAddresses' -Multi -Notes ($anyNote -f 'proxyAddresses')
Add-Column assignedPlans.service -Graph assignedPlans -Path '_planServices' -AD '(licensing)' -Multi `
    -Notes 'Enabled plans only. Rule: user.assignedPlans -any (assignedPlan.service -eq "<service>" -and assignedPlan.capabilityStatus -eq "Enabled")'
Add-Column assignedPlans.servicePlanId -Graph assignedPlans -Path '_planIds' -AD '(licensing)' -Multi `
    -Notes 'Enabled plans only. Rule: user.assignedPlans -any (assignedPlan.servicePlanId -eq "<id>" -and assignedPlan.capabilityStatus -eq "Enabled")'
if ($IncludeMemberOf) {
    Add-Column memberOf -Graph memberOf -Path '_memberOf' -Rules 'Yes (preview)' -Multi `
        -Notes "Direct memberships, shown as Name [objectId]. Rule: user.memberOf -any (group.objectId -in ['<id>'])"
}
# extensionAttribute1-15 (Exchange schema attributes in AD)
foreach ($n in 1..15) {
    Add-Column "extensionAttribute$n" -Graph "onPremisesExtensionAttributes.extensionAttribute$n" -AD "extensionAttribute$n"
}
# Directory extensions (extension_<appId>_<name>). Synced ones are named after their AD attribute.
foreach ($ext in $userExtensions) {
    $adName = if ($ext.isSyncedFromOnPremises) { $ext.name -replace '^extension_[0-9a-fA-F]+_', '' } else { '' }
    $rules = if ($ext.isMultiValued) { 'No' } else { 'Yes' }
    $notes = "Directory extension from app '$($ext.appDisplayName)', type $($ext.dataType)."
    if ($ext.isMultiValued) { $notes += ' Multi-valued extensions are not supported in rules.' }
    Add-Column $ext.name -Graph $ext.name -Rules $rules -AD $adName -Multi:([bool]$ext.isMultiValued) -Notes $notes
}
# Reference only (NOT usable in rules)
$refNote = 'Not usable in dynamic rules.'
Add-Column ref.onPremisesSamAccountName -Graph onPremisesSamAccountName -Rules 'No' -AD 'sAMAccountName' -Notes "$refNote Join key for an AD export."
Add-Column ref.onPremisesUserPrincipalName -Graph onPremisesUserPrincipalName -Rules 'No' -AD 'userPrincipalName' -Notes $refNote
Add-Column ref.onPremisesImmutableId -Graph onPremisesImmutableId -Rules 'No' -AD 'objectGUID or mS-DS-ConsistencyGuid (base64)' -Notes $refNote
Add-Column ref.onPremisesDomainName -Graph onPremisesDomainName -Rules 'No' -AD '(AD domain)' -Notes $refNote
Add-Column ref.onPremisesLastSyncDateTime -Graph onPremisesLastSyncDateTime -Rules 'No' -AD '(last sync time)' -Notes $refNote
Add-Column ref.employeeType -Graph employeeType -Rules 'No' -Notes $refNote
Add-Column ref.employeeOrgData.division -Graph employeeOrgData.division -Rules 'No' -Notes $refNote
Add-Column ref.employeeOrgData.costCenter -Graph employeeOrgData.costCenter -Rules 'No' -Notes $refNote
if ($IncludeManager) {
    Add-Column ref.manager -Graph manager -Path '_manager' -Rules 'No' -AD 'manager' `
        -Notes 'Not a rule property. A Direct Reports for "<manager objectId>" rule groups a manager''s reports, but cannot be combined with other expressions.'
}

# Base properties to $select (top-level names only; extensions are fetched separately)
$selectProps = @(
    $columns |
        Where-Object { -not $_.Path.StartsWith('_') -and -not $_.Path.StartsWith('extension_') } |
        ForEach-Object { $_.Path.Split('.')[0] }
) + @('id', 'assignedPlans') | Select-Object -Unique
$extensionNames = @($userExtensions | ForEach-Object { $_.name })
$extensionChunkSize = 15   # keeps the URL short when the tenant has many extensions

# ======================== FETCH MEMBERS ========================
Write-Host ""
$memberPath = if ($IncludeNested) { 'transitiveMembers' } else { 'members' }
Write-Host "Fetching $(if ($IncludeNested) { 'direct and nested' } else { 'direct' }) user members..." -ForegroundColor Cyan

$usersById = @{}
$sourceGroups = @{}

foreach ($grp in $resolvedGroups) {
    $base = "v1.0/groups/$($grp.id)/$memberPath/microsoft.graph.user"
    $members = @(Invoke-GraphGetAll -Uri "${base}?`$select=$($selectProps -join ',')&`$top=999&`$count=true")

    foreach ($m in $members) {
        if (-not $usersById.ContainsKey($m.id)) {
            $usersById[$m.id] = $m
            $sourceGroups[$m.id] = [System.Collections.Generic.List[string]]::new()
        }
        $sourceGroups[$m.id].Add($grp.displayName)
    }

    for ($i = 0; $i -lt $extensionNames.Count; $i += $extensionChunkSize) {
        $chunk = @($extensionNames[$i..([Math]::Min($i + $extensionChunkSize, $extensionNames.Count) - 1)])
        $extRows = @(Invoke-GraphGetAll -Uri "${base}?`$select=id,$($chunk -join ',')&`$top=999&`$count=true")
        foreach ($r in $extRows) {
            $target = $usersById[$r.id]
            if ($null -eq $target) { continue }
            foreach ($name in $chunk) { $target[$name] = $r[$name] }
        }
    }

    Write-Host "  $($grp.displayName): $($members.Count) user(s)" -ForegroundColor White
}

if ($usersById.Count -eq 0) {
    Write-Host ""
    Write-Host "No user members found. Exiting." -ForegroundColor Yellow
    return
}

# ======================== PER-USER LOOKUPS (optional) ========================
$lookupErrors = 0
if ($IncludeMemberOf -or $IncludeManager) {
    $done = 0
    foreach ($id in @($usersById.Keys)) {
        $done++
        Write-Progress -Activity 'Fetching per-user details' -Status "$done of $($usersById.Count)" `
            -PercentComplete (100 * $done / $usersById.Count)
        $u = $usersById[$id]

        if ($IncludeMemberOf) {
            try {
                $memberOf = @(Invoke-GraphGetAll -Uri "v1.0/users/$id/memberOf/microsoft.graph.group?`$select=id,displayName&`$top=999&`$count=true")
                $u['_memberOf'] = @($memberOf | Sort-Object { $_.displayName } | ForEach-Object { "$($_.displayName) [$($_.id)]" })
            }
            catch {
                $u['_memberOf'] = "ERROR: $($_.Exception.Message)"
                $lookupErrors++
            }
        }

        if ($IncludeManager) {
            try {
                $mgr = Invoke-MgGraphRequest -Method GET -Uri "v1.0/users/$id/manager?`$select=id,displayName,userPrincipalName"
                $mgrName = if ($mgr.userPrincipalName) { $mgr.userPrincipalName } else { $mgr.displayName }
                $u['_manager'] = "$mgrName [$($mgr.id)]"
            }
            catch {
                if ($_.Exception.Message -match 'NotFound|404') {
                    $u['_manager'] = ''   # no manager set
                }
                else {
                    $u['_manager'] = "ERROR: $($_.Exception.Message)"
                    $lookupErrors++
                }
            }
        }
    }
    Write-Progress -Activity 'Fetching per-user details' -Completed
}

# ======================== BUILD ROWS ========================
foreach ($u in $usersById.Values) {
    $enabledPlans = @($u['assignedPlans'] | Where-Object { $_ -and $_['capabilityStatus'] -eq 'Enabled' })
    $u['_planServices'] = @($enabledPlans | ForEach-Object { $_['service'] } | Sort-Object -Unique)
    $u['_planIds'] = @($enabledPlans | ForEach-Object { $_['servicePlanId'] } | Sort-Object -Unique)
}

$rows = @(
    foreach ($u in ($usersById.Values | Sort-Object { $_['displayName'] })) {
        $row = [ordered]@{ SourceGroup = (@($sourceGroups[$u['id']] | Sort-Object -Unique) -join '; ') }
        foreach ($c in $columns) {
            $row[$c.Name] = ConvertTo-CellValue (Get-ValueAtPath -Object $u -Path $c.Path)
        }
        [pscustomobject]$row
    }
)

# ======================== SUMMARY ========================
$total = $rows.Count
$summary = @(
    foreach ($c in $columns) {
        $values = @($rows | ForEach-Object { $_.($c.Name) } | Where-Object { $_ -ne '' })
        $flat = if ($c.Multi) { @($values | ForEach-Object { $_ -split '; ' }) } else { $values }
        $counts = @($flat | Group-Object -NoElement |
                Sort-Object @{ Expression = 'Count'; Descending = $true }, @{ Expression = 'Name'; Descending = $false })

        # Only call it "unique" when there are too many values to list (UPNs, object IDs, ...)
        $topValues = if ($counts.Count -eq 0) { '' }
        elseif ($counts.Count -gt 5 -and $counts[0].Count -eq 1) { '(every value is unique)' }
        else { ($counts | Select-Object -First 5 | ForEach-Object { "$(Format-Short $_.Name) ($($_.Count))" }) -join ' | ' }

        [pscustomobject]@{
            Column           = $c.Name
            UsableInRules    = $c.Rules
            GraphProperty    = $c.Graph
            DefaultADSource  = $c.AD
            UsersWithValue   = $values.Count
            PercentWithValue = [int][Math]::Round(100 * $values.Count / $total, 0)
            DistinctValues   = $counts.Count
            TopValues        = $topValues
            Notes            = $c.Notes
        }
    }
)

# ======================== EXPORT ========================
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}
$label = if ($resolvedGroups.Count -eq 1) { $resolvedGroups[0].displayName } else { "$($resolvedGroups.Count)Groups" }
$label = ($label -replace '[\\/:*?"<>|\s]+', '-').Trim('-')
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$usersPath = Join-Path $OutputFolder "EntraUserAttributes_${label}_${stamp}_Users.csv"
$summaryPath = Join-Path $OutputFolder "EntraUserAttributes_${label}_${stamp}_Summary.csv"

# UTF-8 with BOM so Excel opens accented names correctly. Windows PowerShell 5.1 also needs
# -NoTypeInformation; PowerShell 7 does that by default and warns that the switch is obsolete.
$csvParams = if ($PSVersionTable.PSVersion.Major -ge 6) { @{ Encoding = 'utf8BOM' } }
else { @{ Encoding = 'UTF8'; NoTypeInformation = $true } }
$rows | Export-Csv -LiteralPath $usersPath @csvParams
$summary | Export-Csv -LiteralPath $summaryPath @csvParams

# ======================== DISPLAY ========================
$usable = @($summary | Where-Object { $_.UsableInRules -like 'Yes*' })
$populated = @($usable | Where-Object { $_.UsersWithValue -gt 0 })

Write-Host ""
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "  $total user(s) exported from $($resolvedGroups.Count) group(s)" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "  Users:   $usersPath" -ForegroundColor Green
Write-Host "  Summary: $summaryPath" -ForegroundColor Green
Write-Host ""
Write-Host "Rule-usable attributes with values, most populated first" -ForegroundColor Cyan
Write-Host "($($usable.Count - $populated.Count) of $($usable.Count) rule-usable attributes are empty for every user; see the Summary CSV):" -ForegroundColor DarkGray
$populated |
    Sort-Object @{ Expression = 'PercentWithValue'; Descending = $true }, @{ Expression = 'Column'; Descending = $false } |
    Format-Table @{ Label = 'Column'; Expression = { $_.Column } },
    @{ Label = '% set'; Expression = { $_.PercentWithValue } },
    @{ Label = 'Distinct'; Expression = { $_.DistinctValues } },
    @{ Label = 'Top values'; Expression = { $_.TopValues } } -AutoSize |
    Out-Host

if ($extensionWarning) {
    Write-Host "WARNING: $extensionWarning" -ForegroundColor Yellow
}
if ($lookupErrors -gt 0) {
    Write-Host "WARNING: $lookupErrors per-user lookup(s) failed. Those cells start with 'ERROR:'." -ForegroundColor Yellow
}
