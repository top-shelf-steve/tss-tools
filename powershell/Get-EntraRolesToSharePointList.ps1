# ========================
# CONFIGURATION - Update these before deploying
# ========================
$SharePointSiteUrl = ""              # e.g. https://contoso.sharepoint.com/sites/IT
$ListName = "Entra Roles"           # Name of the pre-created SharePoint list
$UseManagedIdentity = $false         # Set to $true for Azure Automation Runbook

# Sensitivity classification for Entra roles
# Add or move role display names between tiers as needed
$CriticalRoles = @(
    "Global Administrator"
    "Privileged Role Administrator"
    "Privileged Authentication Administrator"
)

$HighRoles = @(
    "Security Administrator"
    "Exchange Administrator"
    "SharePoint Administrator"
    "User Administrator"
    "Conditional Access Administrator"
    "Authentication Administrator"
    "Billing Administrator"
    "Cloud Application Administrator"
    "Application Administrator"
    "Groups Administrator"
    "Helpdesk Administrator"
    "Intune Administrator"
    "Azure Information Protection Administrator"
    "Compliance Administrator"
    "Identity Governance Administrator"
)

# ========================
# AUTHENTICATION
# ========================
if ($UseManagedIdentity) {
    Connect-MgGraph -Identity | Out-Null
    Write-Host "Connected via Managed Identity"
}
else {
    Connect-MgGraph -Scopes "RoleManagement.Read.Directory", "User.Read.All", "Sites.ReadWrite.All" | Out-Null
    Write-Host "Connected interactively"
}

# ========================
# COLLECT ENTRA ROLE DEFINITIONS
# ========================
Write-Host "Fetching Entra role definitions..."

$roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -All
Write-Host "Found $($roleDefinitions.Count) role definitions"

# ========================
# COLLECT ROLE ASSIGNMENTS
# ========================
Write-Host "Fetching active role assignments..."
$activeAssignments = Get-MgRoleManagementDirectoryRoleAssignment -All -ExpandProperty "principal"

# Attempt to fetch PIM eligible assignments (requires RoleEligibilitySchedule.Read.Directory or similar)
Write-Host "Fetching PIM eligible assignments..."
$eligibleAssignments = @()
try {
    $eligibleAssignments = Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -All -ExpandProperty "principal" -ErrorAction Stop
    Write-Host "Found $($eligibleAssignments.Count) eligible assignments"
}
catch {
    Write-Host "Could not retrieve PIM eligible assignments (PIM may not be licensed or permissions insufficient). Continuing without eligible data."
}

# Build lookup: RoleDefinitionId -> list of active assignment principals
$activeLookup = @{}
foreach ($assignment in $activeAssignments) {
    $roleId = $assignment.RoleDefinitionId
    if (-not $activeLookup.ContainsKey($roleId)) {
        $activeLookup[$roleId] = @()
    }
    $displayName = $assignment.Principal.AdditionalProperties["displayName"]
    if ($displayName) {
        $activeLookup[$roleId] += $displayName
    }
}

# Build lookup: RoleDefinitionId -> list of eligible assignment principals
$eligibleLookup = @{}
foreach ($assignment in $eligibleAssignments) {
    $roleId = $assignment.RoleDefinitionId
    if (-not $eligibleLookup.ContainsKey($roleId)) {
        $eligibleLookup[$roleId] = @()
    }
    $displayName = $assignment.Principal.AdditionalProperties["displayName"]
    if ($displayName) {
        $eligibleLookup[$roleId] += $displayName
    }
}

# ========================
# PROCESS ROLES
# ========================
$results = @()
$roleIndex = 0
$totalRoles = $roleDefinitions.Count

foreach ($role in $roleDefinitions) {
    $roleIndex++
    $pct = [math]::Round(($roleIndex / $totalRoles) * 50)
    Write-Progress -Activity "Entra Roles Export" -Status "Processing role $roleIndex of $totalRoles - $($role.DisplayName)" -PercentComplete $pct

    # Determine sensitivity
    $sensitivity = if ($CriticalRoles -contains $role.DisplayName) { "Critical" }
    elseif ($HighRoles -contains $role.DisplayName) { "High" }
    else { "Standard" }

    # Active assignments
    $activeNames = @()
    if ($activeLookup.ContainsKey($role.Id)) {
        $activeNames = $activeLookup[$role.Id] | Sort-Object -Unique
    }
    $activeAssignmentStr = $activeNames -join "; "
    $activeCount = $activeNames.Count

    # Eligible assignments
    $eligibleNames = @()
    if ($eligibleLookup.ContainsKey($role.Id)) {
        $eligibleNames = $eligibleLookup[$role.Id] | Sort-Object -Unique
    }
    $eligibleAssignmentStr = $eligibleNames -join "; "
    $eligibleCount = $eligibleNames.Count

    # Is enabled (has at least one assignment of any type)
    $isEnabled = ($activeCount -gt 0) -or ($eligibleCount -gt 0)

    $results += [PSCustomObject]@{
        RoleName            = $role.DisplayName
        Description         = if ($role.Description) { $role.Description } else { "" }
        Sensitivity         = $sensitivity
        IsBuiltIn           = $role.IsBuiltIn
        IsEnabled           = $isEnabled
        ActiveAssignments   = $activeAssignmentStr
        ActiveCount         = $activeCount
        EligibleAssignments = $eligibleAssignmentStr
        EligibleCount       = $eligibleCount
    }
}

Write-Host "Processed $($results.Count) roles"

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

# Fetch existing list items and build a lookup by RoleName
$existingItems = Get-MgSiteListItem -SiteId $site.Id -ListId $list.Id -All -Expand "fields"
$existingLookup = @{}
foreach ($item in $existingItems) {
    $name = $item.Fields.AdditionalProperties["RoleName"]
    if ($name) {
        $existingLookup[$name] = $item
    }
}

# Track which RoleNames are still current (for stale cleanup)
$currentRoleNames = @{}

# Upsert: update existing items or create new ones
$sortedResults = $results | Sort-Object RoleName
$syncCount = 0
$total = @($sortedResults).Count
foreach ($row in $sortedResults) {
    $syncCount++
    Write-Progress -Activity "SharePoint Sync" -Status "Syncing $syncCount of $total - $($row.RoleName)" -PercentComplete ([math]::Round(($syncCount / $total) * 100))

    $currentRoleNames[$row.RoleName] = $true

    $fields = @{
        "RoleName"            = $row.RoleName
        "Description"         = $row.Description
        "Sensitivity"         = $row.Sensitivity
        "IsBuiltIn"           = $row.IsBuiltIn
        "IsEnabled"           = $row.IsEnabled
        "ActiveAssignments"   = $row.ActiveAssignments
        "ActiveCount"         = $row.ActiveCount
        "EligibleAssignments" = $row.EligibleAssignments
        "EligibleCount"       = $row.EligibleCount
    }

    if ($existingLookup.ContainsKey($row.RoleName)) {
        # Update existing item
        $itemId = $existingLookup[$row.RoleName].Id
        Update-MgSiteListItem -SiteId $site.Id -ListId $list.Id -ListItemId $itemId -BodyParameter @{ fields = $fields }
    }
    else {
        # Create new item
        New-MgSiteListItem -SiteId $site.Id -ListId $list.Id -BodyParameter @{ fields = $fields }
    }
}

# Remove stale items (roles that no longer exist)
$staleCount = 0
foreach ($item in $existingItems) {
    $name = $item.Fields.AdditionalProperties["RoleName"]
    if ($name -and -not $currentRoleNames.ContainsKey($name)) {
        Remove-MgSiteListItem -SiteId $site.Id -ListId $list.Id -ListItemId $item.Id
        $staleCount++
    }
}

Write-Progress -Activity "SharePoint Sync" -Completed
Write-Host "Sync complete: $syncCount items synced, $staleCount stale items removed."
