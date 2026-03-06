# ========================
# CONFIGURATION - Update these before deploying
# ========================
$SharePointSiteUrl = ""              # e.g. https://contoso.sharepoint.com/sites/IT
$ListName = "AVD Groups"             # Name of the pre-created SharePoint list
$UseManagedIdentity = $false         # Set to $true for Azure Automation Runbook

# Group name prefixes to search for — add all your AVD group prefixes here
$GroupPrefixes = @(
    "avd-"
    # "avd-pooled-"
    # "avd-personal-"
    # "avd-eastus-"
    # "avd-westus-"
    # "avd-prod-"
    # "avd-dev-"
    # "avd-test-"
)

# ========================
# AUTHENTICATION
# ========================
if ($UseManagedIdentity) {
    Connect-MgGraph -Identity | Out-Null
    Write-Host "Connected via Managed Identity"
}
else {
    Connect-MgGraph -Scopes "Group.Read.All", "Sites.ReadWrite.All" | Out-Null
    Write-Host "Connected interactively"
}

# ========================
# COLLECT ENTRA GROUP DATA
# ========================
Write-Host "Searching for AVD groups matching prefixes: $($GroupPrefixes -join ', ')"

$groups = @()
$prefixIndex = 0
$totalPrefixes = $GroupPrefixes.Count

foreach ($prefix in $GroupPrefixes) {
    $prefixIndex++
    $pct = [math]::Round(($prefixIndex / $totalPrefixes) * 50)
    Write-Progress -Activity "AVD Groups Export" -Status "Searching groups with prefix '$prefix'..." -PercentComplete $pct

    $filter = "startsWith(displayName, '$prefix')"
    $matchedGroups = Get-MgGroup -Filter $filter -All -Property "Id,DisplayName,Description,GroupTypes,MembershipRule,MembershipRuleProcessingState,Mail,SecurityEnabled,MailEnabled" -ErrorAction SilentlyContinue

    foreach ($g in $matchedGroups) {
        # Avoid duplicates if a group matches multiple prefixes
        if ($groups.Id -contains $g.Id) { continue }

        $groupType = if ($g.GroupTypes -contains "DynamicMembership") { "Dynamic" } else { "Assigned" }
        $dynamicQuery = if ($groupType -eq "Dynamic" -and $g.MembershipRule) { $g.MembershipRule } else { "" }
        $members = if ($g.DisplayName -like "UG*") { "User" } elseif ($g.DisplayName -like "DG*") { "Device" } else { "" }
        $app = if ($g.DisplayName -like "*App-*") { $true } else { $false }

        $groups += [PSCustomObject]@{
            GroupName    = $g.DisplayName
            Description  = if ($g.Description) { $g.Description } else { "" }
            GroupType    = $groupType
            DynamicQuery = $dynamicQuery
            Members      = $members
            App          = $app
            GroupId      = $g.Id
        }
    }

    Write-Host "  Prefix '$prefix' — found $(@($matchedGroups).Count) group(s)"
}

Write-Host "Collected $($groups.Count) unique AVD group(s) total"

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

# Fetch existing list items and build a lookup by GroupName
$existingItems = Get-MgSiteListItem -SiteId $site.Id -ListId $list.Id -All -Expand "fields"
$existingLookup = @{}
foreach ($item in $existingItems) {
    $name = $item.Fields.AdditionalProperties["GroupName"]
    if ($name) {
        $existingLookup[$name] = $item
    }
}

# Track which GroupNames are still current (for stale cleanup)
$currentGroupNames = @{}

# Upsert: update existing items or create new ones
$sortedResults = $groups | Sort-Object GroupName
$syncCount = 0
$total = @($sortedResults).Count
foreach ($row in $sortedResults) {
    $syncCount++
    Write-Progress -Activity "SharePoint Sync" -Status "Syncing $syncCount of $total - $($row.GroupName)" -PercentComplete ([math]::Round(($syncCount / $total) * 100))

    $currentGroupNames[$row.GroupName] = $true

    $fields = @{
        "GroupName"    = $row.GroupName
        "Description"  = $row.Description
        "GroupType"    = $row.GroupType
        "DynamicQuery" = $row.DynamicQuery
        "Members"      = $row.Members
        "App"          = $row.App
    }

    if ($existingLookup.ContainsKey($row.GroupName)) {
        # Update existing item
        $itemId = $existingLookup[$row.GroupName].Id
        Update-MgSiteListItem -SiteId $site.Id -ListId $list.Id -ListItemId $itemId -BodyParameter @{ fields = $fields }
    }
    else {
        # Create new item
        New-MgSiteListItem -SiteId $site.Id -ListId $list.Id -BodyParameter @{ fields = $fields }
    }
}

# Remove stale items (groups that no longer exist or no longer match prefixes)
$staleCount = 0
foreach ($item in $existingItems) {
    $name = $item.Fields.AdditionalProperties["GroupName"]
    if ($name -and -not $currentGroupNames.ContainsKey($name)) {
        Remove-MgSiteListItem -SiteId $site.Id -ListId $list.Id -ListItemId $item.Id
        $staleCount++
    }
}

Write-Progress -Activity "SharePoint Sync" -Completed
Write-Host "Sync complete: $syncCount items synced, $staleCount stale items removed."
