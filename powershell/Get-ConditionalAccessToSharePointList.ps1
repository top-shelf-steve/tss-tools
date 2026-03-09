# ========================
# CONFIGURATION - Update these before deploying
# ========================
$SharePointSiteUrl = ""              # e.g. https://contoso.sharepoint.com/sites/IT
$ListName = "Conditional Access"     # Name of the pre-created SharePoint list
$UseManagedIdentity = $false         # Set to $true for Azure Automation Runbook

# ========================
# AUTHENTICATION
# ========================
if ($UseManagedIdentity) {
    Connect-MgGraph -Identity | Out-Null
    Write-Host "Connected via Managed Identity"
}
else {
    Connect-MgGraph -Scopes "Policy.Read.All", "Group.Read.All", "Application.Read.All", "Sites.ReadWrite.All" | Out-Null
    Write-Host "Connected interactively"
}

# ========================
# COLLECT CONDITIONAL ACCESS POLICIES
# ========================
Write-Host "Fetching Conditional Access policies..."

$policies = Get-MgIdentityConditionalAccessPolicy -All

Write-Host "Found $($policies.Count) Conditional Access policies"

# Build lookup for group and service principal display names
Write-Host "Building display name lookups..."

$groupIds = @()
$spIds = @()

foreach ($policy in $policies) {
    $groupIds += $policy.Conditions.Users.IncludeGroups
    $groupIds += $policy.Conditions.Users.ExcludeGroups
    $spIds += $policy.Conditions.Applications.IncludeApplications | Where-Object { $_ -notin @("All", "None", "Office365") }
    $spIds += $policy.Conditions.Applications.ExcludeApplications | Where-Object { $_ -notin @("All", "None", "Office365") }
}

$groupIds = @($groupIds | Where-Object { $_ } | Select-Object -Unique)
$spIds = @($spIds | Where-Object { $_ } | Select-Object -Unique)

$groupLookup = @{}
foreach ($id in $groupIds) {
    $group = Get-MgGroup -GroupId $id -Property "DisplayName" -ErrorAction SilentlyContinue
    if ($group) { $groupLookup[$id] = $group.DisplayName }
}

$appLookup = @{
    "All"       = "All cloud apps"
    "None"      = "None"
    "Office365" = "Office 365"
}
foreach ($id in $spIds) {
    $sp = Get-MgServicePrincipal -Filter "appId eq '$id'" -Property "DisplayName" -ErrorAction SilentlyContinue
    if ($sp) { $appLookup[$id] = $sp.DisplayName }
}

# ========================
# PROCESS POLICIES
# ========================
$results = @()
$policyIndex = 0
$totalPolicies = $policies.Count

foreach ($policy in $policies) {
    $policyIndex++
    $pct = [math]::Round(($policyIndex / $totalPolicies) * 50)
    Write-Progress -Activity "Conditional Access Export" -Status "Processing policy $policyIndex of $totalPolicies - $($policy.DisplayName)" -PercentComplete $pct

    # --- Grant controls ---
    $grantParts = @()
    if ($policy.GrantControls.BuiltInControls -contains "block") {
        $grantParts += "Block"
    }
    else {
        if ($policy.GrantControls.BuiltInControls -contains "mfa") { $grantParts += "Require MFA" }
        if ($policy.GrantControls.BuiltInControls -contains "compliantDevice") { $grantParts += "Require compliant device" }
        if ($policy.GrantControls.BuiltInControls -contains "domainJoinedDevice") { $grantParts += "Require Entra joined device" }
        if ($policy.GrantControls.BuiltInControls -contains "approvedApplication") { $grantParts += "Require approved app" }
        if ($policy.GrantControls.BuiltInControls -contains "compliantApplication") { $grantParts += "Require app protection policy" }
        if ($policy.GrantControls.BuiltInControls -contains "passwordChange") { $grantParts += "Require password change" }

        if ($policy.GrantControls.AuthenticationStrength.Id) {
            $grantParts += "Require authentication strength"
        }

        if ($grantParts.Count -gt 1 -and $policy.GrantControls.Operator) {
            $grantParts = @("($($policy.GrantControls.Operator.ToUpper())): " + ($grantParts -join ", "))
        }
    }
    $grantControls = if ($grantParts.Count -gt 0) { $grantParts -join ", " } else { "" }

    # --- Session controls ---
    $sessionParts = @()
    if ($policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled) { $sessionParts += "App enforced restrictions" }
    if ($policy.SessionControls.CloudAppSecurity.IsEnabled) { $sessionParts += "Defender for Cloud Apps" }
    if ($policy.SessionControls.SignInFrequency.IsEnabled) {
        $sessionParts += "Sign-in frequency: $($policy.SessionControls.SignInFrequency.Value) $($policy.SessionControls.SignInFrequency.Type)"
    }
    if ($policy.SessionControls.PersistentBrowser.IsEnabled) {
        $sessionParts += "Persistent browser: $($policy.SessionControls.PersistentBrowser.Mode)"
    }
    $sessionControls = if ($sessionParts.Count -gt 0) { $sessionParts -join "; " } else { "" }

    # --- Included users/groups ---
    $includeParts = @()
    if ($policy.Conditions.Users.IncludeUsers -contains "All") { $includeParts += "All users" }
    elseif ($policy.Conditions.Users.IncludeUsers -contains "GuestsOrExternalUsers") { $includeParts += "Guests/external users" }
    elseif ($policy.Conditions.Users.IncludeUsers -contains "None") { $includeParts += "None" }
    else {
        foreach ($u in $policy.Conditions.Users.IncludeUsers) {
            if ($u -and $u -notin @("All", "None", "GuestsOrExternalUsers")) { $includeParts += $u }
        }
    }
    foreach ($gid in $policy.Conditions.Users.IncludeGroups) {
        $includeParts += if ($groupLookup[$gid]) { $groupLookup[$gid] } else { $gid }
    }
    if ($policy.Conditions.Users.IncludeRoles.Count -gt 0) { $includeParts += "$($policy.Conditions.Users.IncludeRoles.Count) role(s)" }
    $assignedUsers = $includeParts -join "; "

    # --- Excluded users/groups ---
    $excludeParts = @()
    foreach ($u in $policy.Conditions.Users.ExcludeUsers) {
        if ($u -and $u -notin @("GuestsOrExternalUsers")) { $excludeParts += $u }
        elseif ($u -eq "GuestsOrExternalUsers") { $excludeParts += "Guests/external users" }
    }
    foreach ($gid in $policy.Conditions.Users.ExcludeGroups) {
        $excludeParts += if ($groupLookup[$gid]) { $groupLookup[$gid] } else { $gid }
    }
    if ($policy.Conditions.Users.ExcludeRoles.Count -gt 0) { $excludeParts += "$($policy.Conditions.Users.ExcludeRoles.Count) role(s)" }
    $excludedUsers = $excludeParts -join "; "

    # --- Assigned apps ---
    $appParts = @()
    foreach ($appId in $policy.Conditions.Applications.IncludeApplications) {
        $appParts += if ($appLookup[$appId]) { $appLookup[$appId] } else { $appId }
    }
    $assignedApps = $appParts -join "; "

    # --- Conditions summary ---
    $condParts = @()
    if ($policy.Conditions.Platforms.IncludePlatforms.Count -gt 0) {
        $condParts += "Platforms: $($policy.Conditions.Platforms.IncludePlatforms -join ', ')"
    }
    if ($policy.Conditions.Locations.IncludeLocations.Count -gt 0) {
        $condParts += "Locations: $($policy.Conditions.Locations.IncludeLocations -join ', ')"
    }
    if ($policy.Conditions.UserRiskLevels.Count -gt 0) {
        $condParts += "User risk: $($policy.Conditions.UserRiskLevels -join ', ')"
    }
    if ($policy.Conditions.SignInRiskLevels.Count -gt 0) {
        $condParts += "Sign-in risk: $($policy.Conditions.SignInRiskLevels -join ', ')"
    }
    if ($policy.Conditions.ClientAppTypes.Count -gt 0 -and $policy.Conditions.ClientAppTypes -notcontains "all") {
        $condParts += "Client apps: $($policy.Conditions.ClientAppTypes -join ', ')"
    }
    if ($policy.Conditions.Devices.DeviceFilter.Rule) {
        $condParts += "Device filter: $($policy.Conditions.Devices.DeviceFilter.Mode) - $($policy.Conditions.Devices.DeviceFilter.Rule)"
    }
    $conditions = $condParts -join "; "

    # --- Dates ---
    $created = if ($policy.CreatedDateTime) { $policy.CreatedDateTime.ToString("yyyy-MM-dd HH:mm") } else { "" }
    $modified = if ($policy.ModifiedDateTime) { $policy.ModifiedDateTime.ToString("yyyy-MM-dd HH:mm") } else { "" }

    $results += [PSCustomObject]@{
        PolicyName      = $policy.DisplayName
        State           = $policy.State
        GrantControls   = $grantControls
        SessionControls = $sessionControls
        AssignedUsers   = $assignedUsers
        ExcludedUsers   = $excludedUsers
        AssignedApps    = $assignedApps
        Conditions      = $conditions
        CreatedDate     = $created
        ModifiedDate    = $modified
    }
}

Write-Host "Processed $($results.Count) policies"

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

# Fetch existing list items and build a lookup by PolicyName
$existingItems = Get-MgSiteListItem -SiteId $site.Id -ListId $list.Id -All -Expand "fields"
$existingLookup = @{}
foreach ($item in $existingItems) {
    $name = $item.Fields.AdditionalProperties["PolicyName"]
    if ($name) {
        $existingLookup[$name] = $item
    }
}

# Track which PolicyNames are still current (for stale cleanup)
$currentPolicyNames = @{}

# Upsert: update existing items or create new ones
$sortedResults = $results | Sort-Object PolicyName
$syncCount = 0
$total = @($sortedResults).Count
foreach ($row in $sortedResults) {
    $syncCount++
    Write-Progress -Activity "SharePoint Sync" -Status "Syncing $syncCount of $total - $($row.PolicyName)" -PercentComplete ([math]::Round(($syncCount / $total) * 100))

    $currentPolicyNames[$row.PolicyName] = $true

    $fields = @{
        "PolicyName"      = $row.PolicyName
        "State"           = $row.State
        "GrantControls"   = $row.GrantControls
        "SessionControls" = $row.SessionControls
        "AssignedUsers"   = $row.AssignedUsers
        "ExcludedUsers"   = $row.ExcludedUsers
        "AssignedApps"    = $row.AssignedApps
        "Conditions"      = $row.Conditions
        "CreatedDate"     = $row.CreatedDate
        "ModifiedDate"    = $row.ModifiedDate
    }

    if ($existingLookup.ContainsKey($row.PolicyName)) {
        # Update existing item
        $itemId = $existingLookup[$row.PolicyName].Id
        Update-MgSiteListItem -SiteId $site.Id -ListId $list.Id -ListItemId $itemId -BodyParameter @{ fields = $fields }
    }
    else {
        # Create new item
        New-MgSiteListItem -SiteId $site.Id -ListId $list.Id -BodyParameter @{ fields = $fields }
    }
}

# Remove stale items (policies that no longer exist)
$staleCount = 0
foreach ($item in $existingItems) {
    $name = $item.Fields.AdditionalProperties["PolicyName"]
    if ($name -and -not $currentPolicyNames.ContainsKey($name)) {
        Remove-MgSiteListItem -SiteId $site.Id -ListId $list.Id -ListItemId $item.Id
        $staleCount++
    }
}

Write-Progress -Activity "SharePoint Sync" -Completed
Write-Host "Sync complete: $syncCount items synced, $staleCount stale items removed."
