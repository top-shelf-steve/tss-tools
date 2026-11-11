<#
.SYNOPSIS
    Exports all Entra ID Conditional Access policies to JSON, ready for review or AI-assisted analysis.

.DESCRIPTION
    Pulls every Conditional Access policy from Microsoft Graph, along with the named locations and
    authentication contexts they reference, and writes them to JSON.

    By default the export is shaped for feeding into an LLM (Claude, Copilot, etc.):
      - Object IDs (users, groups, roles, apps, service principals, named locations, authentication
        contexts) are replaced inline with "Display Name [id]" so the policies can be read without a
        separate lookup. References to objects that no longer exist are flagged as UNRESOLVED.
      - Null and empty properties are removed to cut noise and token count.
      - An "exportInfo" header explains these conventions to whoever (or whatever) reads the file.

    Output written to -OutputPath:
      ConditionalAccessPolicies.json   Everything in one file: policies, named locations, auth contexts
      policies\<Policy Name>.json      One file per policy (for smaller context windows or diffing)

    Use -SkipNameResolution and -KeepEmptyValues to get output that is close to Graph's raw response.

.PARAMETER OutputPath
    Folder to write the export to. Created if it doesn't exist.
    Defaults to .\CA-Export_<yyyyMMdd_HHmmss> in the current directory.

.PARAMETER TenantId
    Tenant ID or domain to sign in to. Useful when your account can access multiple tenants.

.PARAMETER SkipNameResolution
    Leave object IDs as raw GUIDs. Only Policy.Read.All is requested.

.PARAMETER KeepEmptyValues
    Keep null and empty properties instead of removing them.

.PARAMETER UseBeta
    Query the Graph beta endpoint, which can expose newer or preview policy conditions.

.NOTES
    Requires the Microsoft.Graph.Authentication module:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser

    Delegated permissions: Policy.Read.All, plus Directory.Read.All for name resolution.
    The signed-in account needs a role that can read CA policies, such as Global Reader,
    Security Reader or Conditional Access Administrator.

    The export describes the tenant's security posture (excluded accounts, trusted IP ranges, etc.).
    Treat it as sensitive and only share it with AI tools your organization has approved.

.EXAMPLE
    .\Get-ConditionalAccessPoliciesToJSON.ps1

.EXAMPLE
    .\Get-ConditionalAccessPoliciesToJSON.ps1 -TenantId "contoso.onmicrosoft.com" -OutputPath "C:\Exports\Contoso-CA"

.EXAMPLE
    .\Get-ConditionalAccessPoliciesToJSON.ps1 -SkipNameResolution -KeepEmptyValues
    Export without name resolution or cleanup, e.g. as a backup of the policy definitions.
#>

#Requires -Modules Microsoft.Graph.Authentication

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\CA-Export_$(Get-Date -Format 'yyyyMMdd_HHmmss')",

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [switch]$SkipNameResolution,

    [switch]$KeepEmptyValues,

    [switch]$UseBeta
)

$ErrorActionPreference = 'Stop'
$apiVersion = if ($UseBeta) { 'beta' } else { 'v1.0' }
$guidPattern = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'

#region Helper functions

function Invoke-GraphGet {
    # GET a Graph URI and parse the response into PSCustomObjects (keeps Graph's property order).
    param([string]$Uri)
    Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType Json | ConvertFrom-Json
}

function Get-GraphCollection {
    # GET a Graph collection, following @odata.nextLink until every page has been retrieved.
    param([string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    while ($Uri) {
        $response = Invoke-GraphGet -Uri $Uri
        foreach ($item in $response.value) { $items.Add($item) }
        $Uri = $response.'@odata.nextLink'
    }
    return , $items.ToArray()
}

function Test-IsEmpty {
    param($Value)
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return $Value.Length -eq 0 }
    if ($Value -is [System.Collections.ICollection]) { return $Value.Count -eq 0 }
    return $false
}

function ConvertTo-CleanObject {
    # Recursively converts parsed Graph JSON into ordered dictionaries and arrays, dropping
    # "@odata.context" annotations and (unless -KeepEmptyValues) null or empty values.
    param($InputObject)

    if ($null -eq $InputObject) { return $null }

    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $clean = [ordered]@{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            if ($prop.Name -like '*@odata.context') { continue }
            $value = ConvertTo-CleanObject -InputObject $prop.Value
            if (-not $KeepEmptyValues -and (Test-IsEmpty $value)) { continue }
            $clean[$prop.Name] = $value
        }
        return $clean
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        $array = @(foreach ($item in $InputObject) { ConvertTo-CleanObject -InputObject $item })
        return , $array
    }

    return $InputObject
}

function Get-Node {
    # Walks a path of keys through nested dictionaries. Returns $null if any step is missing.
    param($Root, [string[]]$Path)
    $node = $Root
    foreach ($key in $Path) {
        if ($node -isnot [System.Collections.IDictionary] -or -not $node.Contains($key)) { return $null }
        $node = $node[$key]
    }
    return , $node
}

function Format-Reference {
    # Turns an ID into "Display Name [id]". Built-in keywords (All, None, Office365...) pass through
    # unchanged, as do IDs whose lookup failed (so a permissions error isn't mistaken for a deleted object).
    param([string]$Value, [hashtable]$Lookup)

    if ($Lookup.Names.ContainsKey($Value)) { return "$($Lookup.Names[$Value]) [$Value]" }
    if ($Lookup.IdsAreGuids -and $Value -notmatch $guidPattern) { return $Value }
    if (-not $Lookup.Available -or $failedIds.Contains($Value)) { return $Value }

    $script:unresolvedCount++
    return "UNRESOLVED - $($Lookup.NotFound) [$Value]"
}

function Write-JsonFile {
    # Writes UTF-8 without a BOM, which some tools choke on.
    param($Object, [string]$Path)
    $json = ConvertTo-Json -InputObject $Object -Depth 50
    [System.IO.File]::WriteAllText($Path, $json, [System.Text.UTF8Encoding]::new($false))
}

#endregion

#region Connect
$scopes = @('Policy.Read.All')
if (-not $SkipNameResolution) { $scopes += 'Directory.Read.All' }

$connectParams = @{ Scopes = $scopes; NoWelcome = $true }
if ($TenantId) { $connectParams.TenantId = $TenantId }

try {
    Connect-MgGraph @connectParams
}
catch {
    throw "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
}

$context = Get-MgContext
Write-Host "Connected to tenant $($context.TenantId) as $($context.Account)" -ForegroundColor Green
#endregion

#region Fetch policies and the objects they depend on
Write-Host "Fetching Conditional Access policies (Graph $apiVersion)..." -ForegroundColor Cyan
try {
    $rawPolicies = Get-GraphCollection -Uri "$apiVersion/identity/conditionalAccess/policies"
}
catch {
    throw "Failed to read Conditional Access policies. The account needs Policy.Read.All and a role such as Global Reader, Security Reader or Conditional Access Administrator. Error: $($_.Exception.Message)"
}
Write-Host "  Found $($rawPolicies.Count) policies"

Write-Host "Fetching named locations and authentication contexts..." -ForegroundColor Cyan
$rawLocations = @()
try {
    $rawLocations = Get-GraphCollection -Uri "$apiVersion/identity/conditionalAccess/namedLocations"
}
catch {
    Write-Warning "Could not read named locations: $($_.Exception.Message)"
}

$rawAuthContexts = @()
try {
    $rawAuthContexts = Get-GraphCollection -Uri "$apiVersion/identity/conditionalAccess/authenticationContextClassReferences"
}
catch {
    Write-Warning "Could not read authentication contexts: $($_.Exception.Message)"
}

$tenantName = $null
try {
    $tenantName = (Invoke-GraphGet -Uri "$apiVersion/organization?`$select=displayName").value[0].displayName
}
catch {
    Write-Verbose "Could not read tenant display name: $($_.Exception.Message)"
}

$policies = @(@(foreach ($p in $rawPolicies) { ConvertTo-CleanObject -InputObject $p }) | Sort-Object { $_['displayName'] })
$namedLocations = @(@(foreach ($l in $rawLocations) { ConvertTo-CleanObject -InputObject $l }) | Sort-Object { $_['displayName'] })
$authContexts = @(@(foreach ($a in $rawAuthContexts) { ConvertTo-CleanObject -InputObject $a }) | Sort-Object { $_['id'] })
#endregion

#region Resolve IDs to display names
$script:unresolvedCount = 0
$failedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

if (-not $SkipNameResolution) {
    Write-Host "Resolving object IDs to display names..." -ForegroundColor Cyan

    $lookups = @{
        directory    = @{ Names = @{}; Available = $true; IdsAreGuids = $true;  NotFound = 'object not found in directory, possibly deleted' }
        roles        = @{ Names = @{}; Available = $true; IdsAreGuids = $true;  NotFound = 'unknown directory role' }
        apps         = @{ Names = @{}; Available = $true; IdsAreGuids = $true;  NotFound = 'no service principal for this app ID in tenant' }
        locations    = @{ Names = @{}; Available = $true; IdsAreGuids = $true;  NotFound = 'named location not found, possibly deleted' }
        authContexts = @{ Names = @{}; Available = $true; IdsAreGuids = $false; NotFound = 'authentication context not found' }
    }

    # Which policy properties hold which kind of ID
    $referenceRules = @(
        @{ Path = 'conditions', 'users';              Properties = 'includeUsers', 'excludeUsers', 'includeGroups', 'excludeGroups'; Lookup = 'directory' }
        @{ Path = 'conditions', 'users';              Properties = 'includeRoles', 'excludeRoles';                                  Lookup = 'roles' }
        @{ Path = 'conditions', 'applications';       Properties = 'includeApplications', 'excludeApplications';                    Lookup = 'apps' }
        @{ Path = 'conditions', 'applications';       Properties = 'includeAuthenticationContextClassReferences';                   Lookup = 'authContexts' }
        @{ Path = 'conditions', 'clientApplications'; Properties = 'includeServicePrincipals', 'excludeServicePrincipals';          Lookup = 'directory' }
        @{ Path = 'conditions', 'locations';          Properties = 'includeLocations', 'excludeLocations';                          Lookup = 'locations' }
    )

    # Collect the user/group/service principal object IDs and app IDs referenced by the policies
    $directoryIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $appIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($policy in $policies) {
        foreach ($rule in $referenceRules) {
            if ($rule.Lookup -eq 'directory') { $target = $directoryIds }
            elseif ($rule.Lookup -eq 'apps') { $target = $appIds }
            else { continue }
            $container = Get-Node -Root $policy -Path $rule.Path
            foreach ($prop in $rule.Properties) {
                foreach ($value in (Get-Node -Root $container -Path $prop)) {
                    if ($value -match $guidPattern) { [void]$target.Add($value) }
                }
            }
        }
    }

    # Users, groups and service principals: batch lookup, up to 1000 IDs per call
    $idList = @($directoryIds)
    for ($i = 0; $i -lt $idList.Count; $i += 1000) {
        $chunk = @($idList[$i..([Math]::Min($i + 999, $idList.Count - 1))])
        $body = @{ ids = $chunk; types = @('user', 'group', 'servicePrincipal') } | ConvertTo-Json -Compress
        try {
            $response = Invoke-MgGraphRequest -Method POST -Uri "$apiVersion/directoryObjects/getByIds" -Body $body -ContentType 'application/json' -OutputType Json | ConvertFrom-Json
            foreach ($obj in $response.value) {
                $label = if ($obj.userPrincipalName) { "$($obj.displayName) ($($obj.userPrincipalName))" } else { $obj.displayName }
                $lookups.directory.Names[$obj.id] = $label
            }
        }
        catch {
            Write-Warning "Could not resolve users/groups/service principals; their IDs will be left as-is. Error: $($_.Exception.Message)"
            foreach ($id in $chunk) { [void]$failedIds.Add($id) }
        }
    }
    Write-Host "  Users/groups/service principals: $($lookups.directory.Names.Count) of $($idList.Count) resolved"

    # Directory roles: CA policies reference role template IDs
    try {
        foreach ($role in (Get-GraphCollection -Uri "$apiVersion/roleManagement/directory/roleDefinitions?`$select=id,templateId,displayName")) {
            if ($role.templateId) { $lookups.roles.Names[$role.templateId] = $role.displayName }
            if ($role.id) { $lookups.roles.Names[$role.id] = $role.displayName }
        }
    }
    catch {
        Write-Warning "Could not read directory role definitions; role IDs will be left as-is. Error: $($_.Exception.Message)"
        $lookups.roles.Available = $false
    }

    # Applications: CA policies reference app (client) IDs, so look up each service principal by appId
    foreach ($appId in $appIds) {
        try {
            $sp = Invoke-GraphGet -Uri "$apiVersion/servicePrincipals(appId='$appId')?`$select=displayName"
            $lookups.apps.Names[$appId] = $sp.displayName
        }
        catch {
            $notFound = ("$($_.Exception.Response.StatusCode)" -eq 'NotFound') -or ($_.Exception.Message -match 'NotFound|\b404\b')
            if (-not $notFound) {
                Write-Warning "Could not look up app $appId; its ID will be left as-is. Error: $($_.Exception.Message)"
                [void]$failedIds.Add($appId)
            }
        }
    }
    Write-Host "  Applications: $($lookups.apps.Names.Count) of $($appIds.Count) resolved"

    # Named locations and authentication contexts were already fetched above
    $lookups.locations.Names['00000000-0000-0000-0000-000000000000'] = 'MFA Trusted IPs (legacy per-user MFA setting)'
    foreach ($loc in $namedLocations) { $lookups.locations.Names[$loc['id']] = $loc['displayName'] }
    if ($rawLocations.Count -eq 0) { $lookups.locations.Available = $false }

    foreach ($ctx in $authContexts) { $lookups.authContexts.Names[$ctx['id']] = $ctx['displayName'] }
    if ($rawAuthContexts.Count -eq 0) { $lookups.authContexts.Available = $false }

    # Rewrite the ID arrays in place
    foreach ($policy in $policies) {
        foreach ($rule in $referenceRules) {
            $container = Get-Node -Root $policy -Path $rule.Path
            if ($null -eq $container) { continue }
            $lookup = $lookups[$rule.Lookup]
            foreach ($prop in $rule.Properties) {
                if (-not $container.Contains($prop) -or $null -eq $container[$prop]) { continue }
                $container[$prop] = @(foreach ($value in $container[$prop]) { Format-Reference -Value $value -Lookup $lookup })
            }
        }
    }
}
#endregion

#region Build and write the export
$stateCounts = [ordered]@{}
foreach ($group in ($policies | Group-Object { $_['state'] } | Sort-Object Name)) { $stateCounts[$group.Name] = $group.Count }

# Notes written into the file so an LLM (or a person) knows how to read it
$notes = [System.Collections.Generic.List[string]]::new()
$notes.Add("Policy 'state' values: 'enabled' = enforced; 'enabledForReportingButNotEnforced' = report-only (evaluated and logged, not enforced); 'disabled' = off.")
if (-not $SkipNameResolution) {
    $notes.Add("Object references are shown as 'Display Name [id]'. Users also show their UPN. Role IDs are directory role template IDs; application IDs are app (client) IDs.")
    $notes.Add("'UNRESOLVED' marks a reference to an object that could not be found in the tenant, such as a deleted group or user. A bare GUID means the lookup itself failed.")
}
$notes.Add("Values such as 'All', 'None', 'GuestsOrExternalUsers', 'Office365', 'MicrosoftAdminPortals', 'AllTrusted' and 'ServicePrincipalsInMyTenant' are built-in keywords, not object IDs.")
if (-not $KeepEmptyValues) {
    $notes.Add("Null and empty properties have been removed. If a condition or control is absent from a policy, it is not configured on that policy.")
}
$notes.Add("'namedLocations' holds the full definition (IP ranges or countries) of each named location that policies reference under conditions.locations.")

$export = [ordered]@{
    exportInfo             = [ordered]@{
        tenantId        = $context.TenantId
        tenantName      = $tenantName
        exportedAtUtc   = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        source          = "Microsoft Graph $apiVersion"
        policyCount     = $policies.Count
        policiesByState = $stateCounts
        notes           = $notes.ToArray()
    }
    namedLocations         = $namedLocations
    authenticationContexts = $authContexts
    policies               = $policies
}

$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$policyFolder = Join-Path $OutputPath 'policies'
New-Item -ItemType Directory -Path $policyFolder -Force | Out-Null

$combinedFile = Join-Path $OutputPath 'ConditionalAccessPolicies.json'
Write-JsonFile -Object $export -Path $combinedFile

# One file per policy, named after the policy (with the ID appended if two policies share a name)
$usedNames = @{}
foreach ($policy in $policies) {
    $baseName = ([string]$policy['displayName'] -replace '[\\/:*?"<>|\x00-\x1F]', '_').Trim().TrimEnd('.')
    if (-not $baseName) { $baseName = $policy['id'] }
    if ($baseName.Length -gt 120) { $baseName = $baseName.Substring(0, 120).Trim() }
    if ($usedNames.ContainsKey($baseName)) { $baseName = "${baseName}_$($policy['id'])" }
    $usedNames[$baseName] = $true
    Write-JsonFile -Object $policy -Path (Join-Path $policyFolder "$baseName.json")
}
#endregion

#region Summary
$stateSummary = ($stateCounts.GetEnumerator() | ForEach-Object { "$($_.Value) $($_.Key)" }) -join ', '
$sizeKb = [Math]::Round((Get-Item $combinedFile).Length / 1KB)

Write-Host ""
Write-Host "Export complete" -ForegroundColor Green
Write-Host "  Policies:          $($policies.Count) ($stateSummary)"
Write-Host "  Named locations:   $($namedLocations.Count)"
Write-Host "  Combined file:     $combinedFile (~$sizeKb KB, roughly $([Math]::Round($sizeKb / 4))k tokens)"
Write-Host "  Per-policy files:  $policyFolder"
if ($script:unresolvedCount -gt 0) {
    Write-Host "  $($script:unresolvedCount) reference(s) could not be matched to an existing object. Search the export for 'UNRESOLVED'." -ForegroundColor Yellow
}
#endregion
