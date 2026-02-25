# ========================
# CONFIGURATION - Update these before deploying
# ========================
$SharePointSiteUrl = ""  # e.g. https://contoso.sharepoint.com/sites/IT
$ListName = "Intune Apps"                                                    # Name of the pre-created SharePoint list

# Connect using Managed Identity (for Azure Automation Runbook)
#Connect-MgGraph -Identity
Write-Host "Connected to Microsoft Graph via Managed Identity"

# ========================
# APP TYPE MAPPING
# ========================
$appTypeMap = @{
    "#microsoft.graph.win32LobApp"                  = "Win32 App"
    "#microsoft.graph.officeSuiteApp"               = "M365 App"
    "#microsoft.graph.iosVppApp"                    = "iOS VPP"
    "#microsoft.graph.iosLobApp"                    = "iOS LOB"
    "#microsoft.graph.managedIOSStoreApp"           = "iOS Store App"
    "#microsoft.graph.iosStoreApp"                  = "iOS Store App"
    "#microsoft.graph.androidLobApp"                = "Android LOB"
    "#microsoft.graph.androidManagedStoreApp"       = "Android Managed Google Play"
    "#microsoft.graph.androidStoreApp"              = "Android Store App"
    "#microsoft.graph.managedAndroidStoreApp"       = "Android Store App"
    "#microsoft.graph.androidForWorkApp"            = "Android Enterprise App"
    "#microsoft.graph.microsoftStoreForBusinessApp" = "Microsoft Store App"
    "#microsoft.graph.winGetApp"                    = "WinGet App"
    "#microsoft.graph.webApp"                       = "Web Link"
    "#microsoft.graph.windowsWebApp"                = "Web Link"
    "#microsoft.graph.windowsMicrosoftEdgeApp"      = "Microsoft Edge"
    "#microsoft.graph.windowsUniversalAppX"         = "Windows AppX/MSIX"
    "#microsoft.graph.windowsMobileMSI"             = "Windows MSI"
    "#microsoft.graph.macOSLobApp"                  = "macOS LOB"
    "#microsoft.graph.macOSDmgApp"                  = "macOS DMG"
    "#microsoft.graph.macOSPkgApp"                  = "macOS PKG"
    "#microsoft.graph.macOSMicrosoftEdgeApp"        = "macOS Microsoft Edge"
    "#microsoft.graph.macOSMicrosoftDefenderApp"    = "macOS Microsoft Defender"
    "#microsoft.graph.macOSOfficeSuiteApp"          = "macOS M365 App"
    "#microsoft.graph.macOSMdatpApp"                = "macOS Defender ATP"
    "#microsoft.graph.macOsVppApp"                  = "macOS VPP"
    "#microsoft.graph.windowsAppX"                  = "Windows AppX"
    "#microsoft.graph.windowsStoreApp"              = "Windows Store App"
    "#microsoft.graph.managedApp"                   = "Managed App"
}

# ========================
# FETCH INTUNE APPS
# ========================
Write-Progress -Activity "Intune Apps Report" -Status "Fetching mobile apps..." -PercentComplete 0
$intuneApps = Get-MgDeviceAppManagementMobileApp -All
Write-Host "Found $($intuneApps.Count) Intune apps. Fetching assignments..."

# ========================
# ENRICH WITH ASSIGNMENTS
# ========================
$i = 0
$total = $intuneApps.Count
$results = foreach ($app in $intuneApps) {
    $i++
    $pct = [math]::Round(($i / $total) * 100)
    Write-Progress -Activity "Intune Apps Report" -Status "Processing $i of $total - $($app.DisplayName)" -PercentComplete $pct

    # Resolve friendly app type
    $odataType = $app.AdditionalProperties.'@odata.type'
    $appType = if ($appTypeMap.ContainsKey($odataType)) { $appTypeMap[$odataType] } else { $odataType -replace '#microsoft\.graph\.', '' }

    # Fetch assignments
    $assignments = Get-MgDeviceAppManagementMobileAppAssignment -MobileAppId $app.Id -All

    $requiredGroups = @()
    $availableGroups = @()
    $uninstallGroups = @()

    foreach ($assignment in $assignments) {
        $intent = $assignment.Intent
        $targetType = $assignment.Target.AdditionalProperties.'@odata.type'

        if ($targetType -eq '#microsoft.graph.groupAssignmentTarget' -or $targetType -eq '#microsoft.graph.exclusionGroupAssignmentTarget') {
            $groupId = $assignment.Target.AdditionalProperties.groupId
            try {
                $group = Get-MgGroup -GroupId $groupId -Property "displayName" -ErrorAction SilentlyContinue
                $groupName = $group.DisplayName
            }
            catch {
                $groupName = $groupId
            }
        }
        elseif ($targetType -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
            $groupName = "All Users"
        }
        elseif ($targetType -eq '#microsoft.graph.allDevicesAssignmentTarget') {
            $groupName = "All Devices"
        }
        else {
            $groupName = "Unknown Target"
        }

        switch ($intent) {
            'required' { $requiredGroups += $groupName }
            'available' { $availableGroups += $groupName }
            'availableWithoutEnrollment' { $availableGroups += "$groupName (No Enrollment)" }
            'uninstall' { $uninstallGroups += $groupName }
        }
    }

    [PSCustomObject]@{
        AppName     = $app.DisplayName
        AppId       = $app.Id
        AppType     = $appType
        Publisher   = $app.AdditionalProperties.publisher
        Description = $app.AdditionalProperties.description
        Required    = ($requiredGroups | Sort-Object -Unique) -join '; '
        Available   = ($availableGroups | Sort-Object -Unique) -join '; '
        Uninstall   = ($uninstallGroups | Sort-Object -Unique) -join '; '
    }
}
Write-Progress -Activity "Intune Apps Report" -Completed

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
        "AppName"     = $row.AppName
        "AppId"       = $row.AppId
        "AppType"     = $row.AppType
        "Publisher"   = $row.Publisher
        "Description" = $row.Description
        "Required"    = $row.Required
        "Available"   = $row.Available
        "Uninstall"   = $row.Uninstall
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

# Remove stale items (apps that no longer exist in Intune)
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
