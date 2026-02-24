# ========================
# CONFIGURATION - Update these before deploying
# ========================
$SharePointSiteUrl = ""  # e.g. https://contoso.sharepoint.com/sites/IT
$ListName = "SSO Apps"                                                    # Name of the pre-created SharePoint list

# Connect using Managed Identity (for Azure Automation Runbook)
#Connect-MgGraph -Identity
Write-Host "Connected to Microsoft Graph via Managed Identity"

# Grab all service principals with SSO configured AND assignment required
Write-Progress -Activity "SSO Apps Report" -Status "Fetching service principals..." -PercentComplete 0
$ssoApps = Get-MgServicePrincipal -All -Property "displayName,appId,preferredSingleSignOnMode,appRoleAssignmentRequired,id,keyCredentials,notes,notificationEmailAddresses" |
Where-Object {
    $_.PreferredSingleSignOnMode -in @('saml', 'oidc', 'password', 'linked') -and
    $_.AppRoleAssignmentRequired -eq $true
}

Write-Host "Found $($ssoApps.Count) SSO apps with assignment required. Fetching group assignments..."

# Enrich with group assignments
$i = 0
$total = $ssoApps.Count
$results = foreach ($app in $ssoApps) {
    $i++
    $pct = [math]::Round(($i / $total) * 100)
    Write-Progress -Activity "SSO Apps Report" -Status "Processing $i of $total - $($app.DisplayName)" -PercentComplete $pct

    $assignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $app.Id -All |
    Where-Object { $_.PrincipalType -eq 'Group' }

    # Get the latest SAML token signing certificate expiration
    $signingCertExpiry = $app.KeyCredentials |
    Where-Object { $_.Usage -eq 'Sign' -and $_.Type -eq 'AsymmetricX509Cert' } |
    Sort-Object EndDateTime -Descending |
    Select-Object -First 1 -ExpandProperty EndDateTime

    # Check if SCIM provisioning is enabled
    $scimStatus = "Disabled"
    try {
        $syncJobs = Get-MgServicePrincipalSynchronizationJob -ServicePrincipalId $app.Id -ErrorAction Stop
        if ($syncJobs | Where-Object { $_.Schedule.State -eq 'Active' }) {
            $scimStatus = "Enabled"
        }
    }
    catch {
        # No synchronization configured for this app
    }

    [PSCustomObject]@{
        AppName         = $app.DisplayName
        AppId           = $app.AppId
        SSOMode         = $app.PreferredSingleSignOnMode
        AssignedGroups  = ($assignments | ForEach-Object { $_.PrincipalDisplayName }) -join '; '
        CertExpiration  = if ($signingCertExpiry) { $signingCertExpiry.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') } else { $null }
        CertNotifyEmail = ($app.NotificationEmailAddresses) -join '; '
        SCIM            = $scimStatus
        Notes           = $app.Notes
        SPObjectId      = $app.Id
    }
}
Write-Progress -Activity "SSO Apps Report" -Completed

# ========================
# SYNC TO SHAREPOINT LIST
# ========================

# Resolve the SharePoint site and list
$siteHostAndPath = $SharePointSiteUrl -replace 'https://', '' -split '/', 2
$siteHost = $siteHostAndPath[0]
$sitePath = if ($siteHostAndPath.Length -gt 1) { "/$($siteHostAndPath[1])" } else { "/" }
$site = Get-MgSite -SiteId "${siteHost}:${sitePath}"
$list = Get-MgSiteList -SiteId $site.Id -Filter "displayName eq '$ListName'"

Write-Host "Syncing to SharePoint list '$ListName' on site $($site.DisplayName)..."

# Fetch existing list items and build a lookup by AppId
$existingItems = Get-MgSiteListItem -SiteId $site.Id -ListId $list.Id -All -Expand "fields"
$existingLookup = @{}
foreach ($item in $existingItems) {
    $appId = $item.Fields.AdditionalProperties["AppId"]
    if ($appId) {
        $existingLookup[$appId] = $item
    }
}

# Track which AppIds are still current (for stale cleanup)
$currentAppIds = @{}

# Upsert: update existing items or create new ones
$sortedResults = $results | Sort-Object AppName
$syncCount = 0
$total = @($sortedResults).Count
foreach ($row in $sortedResults) {
    $syncCount++
    Write-Progress -Activity "SharePoint Sync" -Status "Syncing $syncCount of $total - $($row.AppName)" -PercentComplete ([math]::Round(($syncCount / $total) * 100))

    $currentAppIds[$row.AppId] = $true

    $fields = @{
        "AppName"         = $row.AppName
        "AppId"           = $row.AppId
        "SSOMode"         = $row.SSOMode
        "AssignedGroups"  = $row.AssignedGroups
        "CertExpiration"  = $row.CertExpiration
        "CertNotifyEmail" = $row.CertNotifyEmail
        "SCIM"            = $row.SCIM
        "Notes"           = $row.Notes
        "SPObjectId"      = $row.SPObjectId
    }

    if ($existingLookup.ContainsKey($row.AppId)) {
        # Update existing item
        $itemId = $existingLookup[$row.AppId].Id
        Update-MgSiteListItem -SiteId $site.Id -ListId $list.Id -ListItemId $itemId -BodyParameter @{ fields = $fields }
    }
    else {
        # Create new item
        New-MgSiteListItem -SiteId $site.Id -ListId $list.Id -BodyParameter @{ fields = $fields }
    }
}

# Remove stale items (apps that no longer exist in Entra)
$staleCount = 0
foreach ($item in $existingItems) {
    $appId = $item.Fields.AdditionalProperties["AppId"]
    if ($appId -and -not $currentAppIds.ContainsKey($appId)) {
        Remove-MgSiteListItem -SiteId $site.Id -ListId $list.Id -ListItemId $item.Id
        $staleCount++
    }
}

Write-Progress -Activity "SharePoint Sync" -Completed
Write-Host "Sync complete: $syncCount items synced, $staleCount stale items removed."