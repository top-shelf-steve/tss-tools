<#
.SYNOPSIS
    Adds an email address to the SAML certificate "Notification Email" list of every
    SAML SSO enterprise application that already has at least one notification email.

.DESCRIPTION
    The "Notification Email" list in Enterprise applications > [app] > Single sign-on >
    SAML Certificates is the notificationEmailAddresses property of the app's service
    principal. Entra ID emails those addresses 60, 30 and 7 days before the SAML signing
    certificate expires.

    Graph replaces the whole list on update, so for each app the script:
      1. Re-reads the app's current list right before changing it.
      2. Writes back the existing addresses with the new one appended. Nothing is removed
         or reordered.
      3. Reads the list back and checks that every original address and the new one are
         present. If not, the app is marked Failed.

    Apps are skipped, not changed, when:
      - the list is empty (only apps that already have an email are updated)
      - the address is already in the list (match ignores case)
      - the list already holds 5 addresses (the limit in the Entra admin center)

    Only service principals whose SSO mode is SAML (preferredSingleSignOnMode = saml) are
    considered.

    A summary prints at the end and one result object per app is written to the pipeline,
    so you can pipe it to Export-Csv. If any app failed, the script throws at the end.

.PARAMETER EmailAddress
    The address to add, e.g. certs@contoso.com.

.PARAMETER AppId
    Optional. One or more Application (client) IDs to limit the run to. Use this to try
    the change on one test app before running it against every SAML app.

.PARAMETER DryRun
    Look everything up and report what would change, without changing anything. Signs in
    with read-only scope.

.EXAMPLE
    .\Set-SAMLSSONotifyEmail.ps1 -EmailAddress "certs@contoso.com" -DryRun
    # Lists every SAML app and whether the address would be added. Changes nothing.

.EXAMPLE
    .\Set-SAMLSSONotifyEmail.ps1 -EmailAddress "certs@contoso.com" -AppId "11111111-1111-1111-1111-111111111111"
    # Adds the address to one app only.

.EXAMPLE
    .\Set-SAMLSSONotifyEmail.ps1 -EmailAddress "certs@contoso.com" | Export-Csv .\saml-notify.csv -NoTypeInformation
    # Updates every eligible SAML app and saves a per-app report.

.NOTES
    Required module: Microsoft.Graph.Authentication

    Permissions (delegated): Application.Read.All for -DryRun, Application.ReadWrite.All
    otherwise. Your account also needs a role that can edit enterprise apps, such as
    Cloud Application Administrator or Application Administrator.

    AFTER A REAL RUN
    Microsoft's guidance: when notification addresses are set through Graph or PowerShell,
    open Enterprise applications > [app] > Single sign-on and check the Notification Email
    in the SAML Certificates section. For apps that use a custom SAML signing certificate,
    opening that page is what sets up notification registration if it doesn't exist yet.
    Without that check, expiry emails might not be sent. The summary lists the apps to check.
#>

#Requires -Modules Microsoft.Graph.Authentication

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$EmailAddress,

    [Parameter(Mandatory = $false)]
    [string[]]$AppId,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$EmailAddress = $EmailAddress.Trim()
$MaxNotificationEmails = 5
$Select = 'id,appId,displayName,preferredSingleSignOnMode,notificationEmailAddresses'

# ========================
# HELPERS
# ========================
$Results = [System.Collections.Generic.List[object]]::new()

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warn', 'Error', 'Header', 'Detail')]
        [string]$Level = 'Info'
    )
    $color = switch ($Level) {
        'Success' { 'Green' }
        'Warn'    { 'Yellow' }
        'Error'   { 'Red' }
        'Header'  { 'Cyan' }
        'Detail'  { 'Gray' }
        default   { 'White' }
    }
    Write-Host $Message -ForegroundColor $color
}

function Add-Result {
    param(
        $App,
        [ValidateSet('Added', 'Failed', 'Skipped', 'DryRun')][string]$Status,
        [string]$Detail,
        [string[]]$Before,
        [string[]]$After
    )
    $Results.Add([pscustomobject]@{
        DisplayName = $App.displayName
        AppId       = $App.appId
        ObjectId    = $App.id
        Status      = $Status
        Detail      = $Detail
        Before      = ($Before -join '; ')
        After       = ($After -join '; ')
    })
    $level = switch ($Status) {
        'Added'  { 'Success' }
        'Failed' { 'Error' }
        'DryRun' { 'Info' }
        default  { 'Detail' }
    }
    Write-Log "  [$Status] $($App.displayName): $Detail" $level
}

function Get-ErrorText {
    # Pulls the Graph error message out of an Invoke-MgGraphRequest failure when there is one.
    param($ErrorRecord)
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        try { return ($ErrorRecord.ErrorDetails.Message | ConvertFrom-Json).error.message }
        catch { return $ErrorRecord.ErrorDetails.Message }
    }
    return $ErrorRecord.Exception.Message
}

function Get-EmailList {
    # Returns the app's notification addresses. PowerShell unrolls arrays returned from a
    # function (one item comes back as a plain string, none as $null), so wrap every call in @().
    param($App)
    $App.notificationEmailAddresses | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
}

function Get-ServicePrincipal {
    param([string]$ObjectId)
    Invoke-MgGraphRequest -Method GET -Uri "v1.0/servicePrincipals/$ObjectId`?`$select=$Select"
}

# ========================
# CONNECT
# ========================
Write-Log "==========================================" Header
Write-Log "  SAML certificate notification email" Header
Write-Log "  Adding: $EmailAddress" Header
if ($DryRun) { Write-Log "  DRY RUN - no changes will be made" Header }
Write-Log "==========================================" Header

$scope = if ($DryRun) { 'Application.Read.All' } else { 'Application.ReadWrite.All' }
Write-Log "`nConnecting to Microsoft Graph ($scope)..." Header
Connect-MgGraph -Scopes $scope -NoWelcome
$ctx = Get-MgContext
Write-Log "  Connected as $($ctx.Account)" Success

# ========================
# FIND SAML APPS
# ========================
Write-Log "`nFinding SAML SSO apps..." Header
$apps = [System.Collections.Generic.List[object]]::new()

if ($AppId) {
    foreach ($id in $AppId) {
        $id = $id.Trim()
        try {
            $sp = Invoke-MgGraphRequest -Method GET -Uri "v1.0/servicePrincipals(appId='$id')?`$select=$Select"
        }
        catch {
            $missing = @{ displayName = "(appId $id)"; appId = $id; id = $null }
            Add-Result $missing 'Failed' "Lookup failed: $(Get-ErrorText $_)"
            continue
        }
        if ($sp.preferredSingleSignOnMode -ne 'saml') {
            Add-Result $sp 'Skipped' "SSO mode is '$($sp.preferredSingleSignOnMode)', not SAML"
            continue
        }
        $apps.Add($sp)
    }
}
else {
    $uri = "v1.0/servicePrincipals?`$filter=preferredSingleSignOnMode eq 'saml'&`$select=$Select&`$top=999"
    while ($uri) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $uri
        foreach ($sp in $page.value) { $apps.Add($sp) }
        $uri = $page.'@odata.nextLink'
    }
}

$apps = @($apps | Sort-Object { $_.displayName })
Write-Log "  Found $($apps.Count) SAML app(s)" Success

# ========================
# UPDATE EACH APP
# ========================
Write-Log "`nProcessing..." Header
foreach ($app in $apps) {
    $before = @(Get-EmailList $app)

    # Cheap checks on the listing first, so apps that will be skipped cost no extra calls.
    if ($before.Count -eq 0) {
        Add-Result $app 'Skipped' 'No notification email registered' $before $before
        continue
    }
    if ($before -contains $EmailAddress) {
        Add-Result $app 'Skipped' 'Already in the list' $before $before
        continue
    }

    try {
        # Re-read right before writing, in case the list changed since the listing above.
        # The PATCH replaces the whole list, so it must be built from the current value.
        $current = Get-ServicePrincipal $app.id
        $before  = @(Get-EmailList $current)

        if ($before.Count -eq 0) {
            Add-Result $app 'Skipped' 'No notification email registered' $before $before
            continue
        }
        if ($before -contains $EmailAddress) {
            Add-Result $app 'Skipped' 'Already in the list' $before $before
            continue
        }
        if ($before.Count -ge $MaxNotificationEmails) {
            Add-Result $app 'Skipped' "Already has $($before.Count) addresses (limit is $MaxNotificationEmails)" $before $before
            continue
        }

        $after = [string[]]@($before + $EmailAddress)

        if ($DryRun) {
            Add-Result $app 'DryRun' "Would add (now $($before.Count) address(es))" $before $after
            continue
        }

        $body = @{ notificationEmailAddresses = $after } | ConvertTo-Json -Depth 3
        $null = Invoke-MgGraphRequest -Method PATCH -Uri "v1.0/servicePrincipals/$($app.id)" -Body $body -ContentType 'application/json'

        # Read back and check nothing was lost.
        $saved = @(Get-EmailList (Get-ServicePrincipal $app.id))
        $lost  = @($before | Where-Object { $saved -notcontains $_ })
        if ($lost.Count -gt 0 -or $saved -notcontains $EmailAddress) {
            $problem = @()
            if ($lost.Count -gt 0) { $problem += "missing after update: $($lost -join ', ')" }
            if ($saved -notcontains $EmailAddress) { $problem += 'new address not saved' }
            Add-Result $app 'Failed' "Read-back check failed ($($problem -join '; '))" $before $saved
            continue
        }

        Add-Result $app 'Added' "Added ($($saved.Count) address(es) now)" $before $saved
    }
    catch {
        Add-Result $app 'Failed' (Get-ErrorText $_) $before $before
    }
}

# ========================
# SUMMARY
# ========================
Write-Log "`n==========================================" Header
Write-Log "  Summary" Header
Write-Log "==========================================" Header
$Results | Group-Object Status | Sort-Object Name | ForEach-Object {
    Write-Log ("  {0,-8} {1}" -f $_.Name, $_.Count)
}

$added = @($Results | Where-Object Status -eq 'Added')
if ($added.Count -gt 0) {
    Write-Log "`nOpen each of these in Enterprise applications > Single sign-on and check the" Warn
    Write-Log "Notification Email in the SAML Certificates section (Microsoft's guidance after" Warn
    Write-Log "a programmatic change; for custom signing certs it sets up notification registration):" Warn
    $added | ForEach-Object { Write-Log "  - $($_.DisplayName) ($($_.AppId))" Warn }
}

$Results

$failed = @($Results | Where-Object Status -eq 'Failed')
if ($failed.Count -gt 0) {
    throw "$($failed.Count) app(s) failed. See the results above."
}
