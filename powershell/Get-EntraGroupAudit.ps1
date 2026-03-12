#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups

# ======================== CONFIGURATION ========================
# Enter group display names or object IDs to audit (leave empty for interactive prompt)
$GroupsToAudit = @(
    # "SG-MyGroup"
    # "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
)
# ===============================================================

# ======================== GRAPH CONNECTION ========================
$requiredScopes = @(
    "Group.Read.All",
    "Directory.Read.All",
    "Policy.Read.All",
    "Application.Read.All",
    "RoleManagement.Read.Directory",
    "DeviceManagementConfiguration.Read.All",
    "DeviceManagementApps.Read.All",
    "DeviceManagementManagedDevices.Read.All",
    "DeviceManagementServiceConfig.Read.All"
)

$context = Get-MgContext
if (-not $context) {
    Write-Host "Not connected to Microsoft Graph. Connecting..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes $requiredScopes
}
else {
    $missingScopes = $requiredScopes | Where-Object { $_ -notin $context.Scopes }
    if ($missingScopes) {
        Write-Host "Connected but missing required scopes: $($missingScopes -join ', '). Reconnecting..." -ForegroundColor Yellow
        Disconnect-MgGraph | Out-Null
        Connect-MgGraph -Scopes $requiredScopes
    }
    else {
        Write-Host "Already connected to Microsoft Graph as $($context.Account)" -ForegroundColor Green
    }
}

# ======================== HELPER FUNCTIONS ========================

function Invoke-MgGraphRequestAll {
    <#
    .SYNOPSIS
        Handles paginated Graph API requests, following @odata.nextLink automatically.
    #>
    param(
        [Parameter(Mandatory)][string]$Uri
    )
    $allResults = @()
    $currentUri = $Uri
    while ($currentUri) {
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $currentUri -ErrorAction Stop
        }
        catch {
            Write-Host "  Warning: Failed to query $currentUri - $($_.Exception.Message)" -ForegroundColor Yellow
            break
        }
        if ($response.value) {
            $allResults += $response.value
        }
        $currentUri = $response.'@odata.nextLink'
    }
    return $allResults
}

function Get-AssignedGroupIds {
    <#
    .SYNOPSIS
        Extracts group IDs from an array of Intune assignment objects.
    #>
    param(
        [array]$Assignments
    )
    $groupIds = @()
    foreach ($assignment in $Assignments) {
        $target = $assignment.target
        if (-not $target) { continue }
        $type = $target.'@odata.type'
        if ($type -eq '#microsoft.graph.groupAssignmentTarget' -or
            $type -eq '#microsoft.graph.exclusionGroupAssignmentTarget') {
            if ($target.groupId) {
                $groupIds += $target.groupId
            }
        }
    }
    return $groupIds
}

function Get-IntunePolicyAssignments {
    <#
    .SYNOPSIS
        Fetches all policies of a given type and their assignments.
        Returns a hashtable mapping group IDs to arrays of policy info.
    #>
    param(
        [string]$PolicyUri,
        [string]$CategoryName,
        [string]$NameProperty = "displayName"
    )
    $groupMap = @{}
    $policies = Invoke-MgGraphRequestAll -Uri $PolicyUri
    $total = $policies.Count
    $i = 0
    foreach ($policy in $policies) {
        $i++
        if ($total -gt 0) {
            Write-Progress -Activity "Entra Group Audit" -Status "$CategoryName ($i of $total)" -PercentComplete (($i / $total) * 100)
        }
        $policyName = $policy.$NameProperty
        if (-not $policyName) { $policyName = $policy.displayName }
        if (-not $policyName) { $policyName = $policy.id }

        $policyId = $policy.id
        try {
            $assignmentsUri = "$PolicyUri/$policyId/assignments"
            $assignments = Invoke-MgGraphRequestAll -Uri $assignmentsUri
        }
        catch {
            continue
        }

        $assignedGroupIds = Get-AssignedGroupIds -Assignments $assignments
        foreach ($gid in $assignedGroupIds) {
            $gidLower = $gid.ToLower()
            if (-not $groupMap.ContainsKey($gidLower)) { $groupMap[$gidLower] = @() }
            $assignmentType = "Included"
            foreach ($a in $assignments) {
                if ($a.target.groupId -eq $gid -and $a.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget') {
                    $assignmentType = "Excluded"
                    break
                }
            }
            $groupMap[$gidLower] += [PSCustomObject]@{
                Category       = $CategoryName
                PolicyName     = $policyName
                PolicyId       = $policyId
                AssignmentType = $assignmentType
            }
        }
    }
    return $groupMap
}

function Merge-GroupMap {
    <#
    .SYNOPSIS
        Merges a source group map into the master group map.
    #>
    param(
        [hashtable]$Master,
        [hashtable]$Source
    )
    foreach ($key in $Source.Keys) {
        if (-not $Master.ContainsKey($key)) { $Master[$key] = @() }
        $Master[$key] += $Source[$key]
    }
}

# ======================== GROUP INPUT ========================

if ($GroupsToAudit.Count -eq 0) {
    Write-Host ""
    Write-Host "Enter group names or object IDs to audit (one per line, blank line to finish):" -ForegroundColor Cyan
    while ($true) {
        $input = Read-Host "  Group"
        if ([string]::IsNullOrWhiteSpace($input)) { break }
        $GroupsToAudit += $input.Trim()
    }
}

if ($GroupsToAudit.Count -eq 0) {
    Write-Host "No groups specified. Exiting." -ForegroundColor Yellow
    return
}

# ======================== RESOLVE GROUPS ========================

Write-Host ""
Write-Host "Resolving groups..." -ForegroundColor Cyan
$resolvedGroups = @()
foreach ($groupInput in $GroupsToAudit) {
    $group = $null
    # Try as object ID first (GUID pattern)
    if ($groupInput -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
        $group = Get-MgGroup -GroupId $groupInput -Property "id,displayName,membershipRule,groupTypes,assignedLicenses" -ErrorAction SilentlyContinue
    }
    # Try as display name
    if (-not $group) {
        $groups = Get-MgGroup -Filter "displayName eq '$($groupInput -replace "'","''")'" -Property "id,displayName,membershipRule,groupTypes,assignedLicenses" -ErrorAction SilentlyContinue
        if ($groups -is [array]) { $group = $groups[0] } else { $group = $groups }
    }
    if ($group) {
        $resolvedGroups += $group
        Write-Host "  Resolved: $($group.DisplayName) ($($group.Id))" -ForegroundColor Green
    }
    else {
        Write-Host "  NOT FOUND: $groupInput" -ForegroundColor Red
    }
}

if ($resolvedGroups.Count -eq 0) {
    Write-Host "No valid groups found. Exiting." -ForegroundColor Red
    return
}

# Build lookup set of group IDs to check against
$auditGroupIds = @{}
foreach ($g in $resolvedGroups) {
    $auditGroupIds[$g.Id.ToLower()] = $g.DisplayName
}

# ======================== FETCH ALL ASSIGNMENTS ========================

Write-Host ""
Write-Host "Scanning all assignment sources. This may take a few minutes..." -ForegroundColor Cyan
Write-Host ""

$masterGroupMap = @{}

# --- 1. Intune Device Configuration Profiles ---
Write-Host "  [1/18] Device Configuration Profiles..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceManagement/deviceConfigurations" -CategoryName "Device Configuration Profile"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 2. Settings Catalog (Configuration Policies) ---
Write-Host "  [2/18] Settings Catalog Policies..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceManagement/configurationPolicies" -CategoryName "Settings Catalog Policy" -NameProperty "name"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 3. Compliance Policies ---
Write-Host "  [3/18] Compliance Policies..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceManagement/deviceCompliancePolicies" -CategoryName "Compliance Policy"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 4. Mobile Apps ---
Write-Host "  [4/18] Intune Apps..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceAppManagement/mobileApps" -CategoryName "Intune App"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 5. App Configuration Policies (Managed Devices) ---
Write-Host "  [5/18] App Configuration Policies..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceAppManagement/mobileAppConfigurations" -CategoryName "App Configuration Policy"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 6. App Protection Policies - Android ---
Write-Host "  [6/18] App Protection Policies (Android)..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceAppManagement/androidManagedAppProtections" -CategoryName "App Protection Policy (Android)"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 7. App Protection Policies - iOS ---
Write-Host "  [7/18] App Protection Policies (iOS)..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceAppManagement/iosManagedAppProtections" -CategoryName "App Protection Policy (iOS)"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 8. App Protection Policies - Windows ---
Write-Host "  [8/18] App Protection Policies (Windows)..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceAppManagement/mdmWindowsInformationProtectionPolicies" -CategoryName "App Protection Policy (Windows)"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 9. Device Management Scripts (PowerShell) ---
Write-Host "  [9/18] Device Management Scripts..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceManagement/deviceManagementScripts" -CategoryName "Device Management Script"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 10. Proactive Remediations (Device Health Scripts) ---
Write-Host "  [10/18] Proactive Remediations..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceManagement/deviceHealthScripts" -CategoryName "Proactive Remediation"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 11. Windows Autopilot Deployment Profiles ---
Write-Host "  [11/18] Autopilot Deployment Profiles..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceManagement/windowsAutopilotDeploymentProfiles" -CategoryName "Autopilot Deployment Profile"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 12. Enrollment Configurations ---
Write-Host "  [12/18] Enrollment Configurations..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceManagement/deviceEnrollmentConfigurations" -CategoryName "Enrollment Configuration"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 13. Windows Feature Update Profiles ---
Write-Host "  [13/18] Windows Feature Update Profiles..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceManagement/windowsFeatureUpdateProfiles" -CategoryName "Windows Feature Update Profile"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 14. Windows Quality Update Profiles ---
Write-Host "  [14/18] Windows Quality Update Profiles..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceManagement/windowsQualityUpdateProfiles" -CategoryName "Windows Quality Update Profile"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 15. Group Policy Configurations ---
Write-Host "  [15/18] Group Policy Configurations..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceManagement/groupPolicyConfigurations" -CategoryName "Group Policy Configuration"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 16. Endpoint Security Intents ---
Write-Host "  [16/18] Endpoint Security Policies..." -ForegroundColor White
$map = Get-IntunePolicyAssignments -PolicyUri "/beta/deviceManagement/intents" -CategoryName "Endpoint Security Policy"
Merge-GroupMap -Master $masterGroupMap -Source $map

# --- 17. Conditional Access Policies ---
Write-Host "  [17/18] Conditional Access Policies..." -ForegroundColor White
try {
    $caPolicies = Invoke-MgGraphRequestAll -Uri "/v1.0/identity/conditionalAccess/policies"
    foreach ($ca in $caPolicies) {
        $caName = $ca.displayName
        $caId = $ca.id

        # Check included groups
        $includeGroups = @()
        if ($ca.conditions.users.includeGroups) { $includeGroups = $ca.conditions.users.includeGroups }
        foreach ($gid in $includeGroups) {
            $gidLower = $gid.ToLower()
            if (-not $masterGroupMap.ContainsKey($gidLower)) { $masterGroupMap[$gidLower] = @() }
            $masterGroupMap[$gidLower] += [PSCustomObject]@{
                Category       = "Conditional Access Policy"
                PolicyName     = $caName
                PolicyId       = $caId
                AssignmentType = "Included"
            }
        }

        # Check excluded groups
        $excludeGroups = @()
        if ($ca.conditions.users.excludeGroups) { $excludeGroups = $ca.conditions.users.excludeGroups }
        foreach ($gid in $excludeGroups) {
            $gidLower = $gid.ToLower()
            if (-not $masterGroupMap.ContainsKey($gidLower)) { $masterGroupMap[$gidLower] = @() }
            $masterGroupMap[$gidLower] += [PSCustomObject]@{
                Category       = "Conditional Access Policy"
                PolicyName     = $caName
                PolicyId       = $caId
                AssignmentType = "Excluded"
            }
        }
    }
}
catch {
    Write-Host "  Warning: Could not fetch Conditional Access Policies - $($_.Exception.Message)" -ForegroundColor Yellow
}

# --- 18. Enterprise App Role Assignments, Licensing, Directory Roles, Admin Units ---
Write-Host "  [18/18] Entra ID assignments (App Roles, Licensing, Directory Roles, Admin Units)..." -ForegroundColor White
foreach ($group in $resolvedGroups) {
    $gidLower = $group.Id.ToLower()

    # App role assignments (enterprise apps)
    try {
        $appRoles = Invoke-MgGraphRequestAll -Uri "/v1.0/groups/$($group.Id)/appRoleAssignments"
        foreach ($role in $appRoles) {
            $appName = $role.resourceDisplayName
            if (-not $appName) { $appName = $role.resourceId }
            if (-not $masterGroupMap.ContainsKey($gidLower)) { $masterGroupMap[$gidLower] = @() }
            $masterGroupMap[$gidLower] += [PSCustomObject]@{
                Category       = "Enterprise App Role Assignment"
                PolicyName     = $appName
                PolicyId       = $role.resourceId
                AssignmentType = "Assigned"
            }
        }
    }
    catch {
        Write-Host "    Warning: Could not fetch app role assignments for $($group.DisplayName)" -ForegroundColor Yellow
    }

    # Group-based licensing
    if ($group.AssignedLicenses -and $group.AssignedLicenses.Count -gt 0) {
        $skuNames = @()
        foreach ($lic in $group.AssignedLicenses) {
            $skuNames += $lic.SkuId
        }
        if (-not $masterGroupMap.ContainsKey($gidLower)) { $masterGroupMap[$gidLower] = @() }
        $masterGroupMap[$gidLower] += [PSCustomObject]@{
            Category       = "Group-Based License"
            PolicyName     = "License SKUs: $($skuNames -join ', ')"
            PolicyId       = "-"
            AssignmentType = "Assigned"
        }
    }

    # Directory role assignments
    try {
        $roleAssignments = Invoke-MgGraphRequestAll -Uri "/v1.0/roleManagement/directory/roleAssignments?`$filter=principalId eq '$($group.Id)'"
        foreach ($ra in $roleAssignments) {
            $roleDef = $null
            try {
                $roleDef = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/roleManagement/directory/roleDefinitions/$($ra.roleDefinitionId)" -ErrorAction SilentlyContinue
            }
            catch { }
            $roleName = if ($roleDef) { $roleDef.displayName } else { $ra.roleDefinitionId }
            if (-not $masterGroupMap.ContainsKey($gidLower)) { $masterGroupMap[$gidLower] = @() }
            $masterGroupMap[$gidLower] += [PSCustomObject]@{
                Category       = "Azure AD Directory Role"
                PolicyName     = $roleName
                PolicyId       = $ra.roleDefinitionId
                AssignmentType = "Assigned"
            }
        }
    }
    catch {
        Write-Host "    Warning: Could not fetch directory role assignments for $($group.DisplayName)" -ForegroundColor Yellow
    }

    # Administrative unit membership
    try {
        $memberOf = Invoke-MgGraphRequestAll -Uri "/v1.0/groups/$($group.Id)/memberOf"
        foreach ($member in $memberOf) {
            if ($member.'@odata.type' -eq '#microsoft.graph.administrativeUnit') {
                if (-not $masterGroupMap.ContainsKey($gidLower)) { $masterGroupMap[$gidLower] = @() }
                $masterGroupMap[$gidLower] += [PSCustomObject]@{
                    Category       = "Administrative Unit"
                    PolicyName     = $member.displayName
                    PolicyId       = $member.id
                    AssignmentType = "Member"
                }
            }
        }
    }
    catch {
        Write-Host "    Warning: Could not fetch admin unit memberships for $($group.DisplayName)" -ForegroundColor Yellow
    }
}

Write-Progress -Activity "Entra Group Audit" -Completed

# ======================== DISPLAY RESULTS ========================

Write-Host ""
Write-Host ("=" * 100) -ForegroundColor Cyan
Write-Host "  ENTRA GROUP AUDIT RESULTS" -ForegroundColor Cyan
Write-Host ("=" * 100) -ForegroundColor Cyan

$allResults = @()

foreach ($group in $resolvedGroups) {
    $gidLower = $group.Id.ToLower()
    $findings = @()
    if ($masterGroupMap.ContainsKey($gidLower)) {
        $findings = $masterGroupMap[$gidLower]
    }

    Write-Host ""
    Write-Host ("  " + $group.DisplayName) -ForegroundColor White
    Write-Host ("  " + $group.Id) -ForegroundColor DarkGray

    $groupType = if ($group.GroupTypes -contains "DynamicMembership") { "Dynamic" } else { "Assigned" }
    Write-Host "  Type: $groupType" -ForegroundColor DarkGray

    if ($findings.Count -eq 0) {
        Write-Host "  Status: NO ASSIGNMENTS FOUND - likely safe to delete" -ForegroundColor Green
        Write-Host ""
    }
    else {
        Write-Host "  Status: ASSIGNED ($($findings.Count) assignment(s) found) - DO NOT DELETE" -ForegroundColor Red
        Write-Host ""

        # Group findings by category
        $byCategory = $findings | Group-Object -Property Category | Sort-Object Name
        foreach ($cat in $byCategory) {
            Write-Host "    $($cat.Name) ($($cat.Count)):" -ForegroundColor Yellow
            foreach ($finding in $cat.Group) {
                $typeColor = switch ($finding.AssignmentType) {
                    "Excluded" { "DarkYellow" }
                    "Included" { "White" }
                    default { "White" }
                }
                $truncatedName = $finding.PolicyName
                if ($truncatedName.Length -gt 70) { $truncatedName = $truncatedName.Substring(0, 67) + "..." }
                Write-Host "      - $truncatedName [$($finding.AssignmentType)]" -ForegroundColor $typeColor
            }
        }
        Write-Host ""
    }

    # Build export data
    if ($findings.Count -eq 0) {
        $allResults += [PSCustomObject]@{
            GroupName      = $group.DisplayName
            GroupId        = $group.Id
            GroupType      = $groupType
            Status         = "SAFE TO DELETE"
            Category       = "-"
            AssignedTo     = "-"
            AssignmentType = "-"
        }
    }
    else {
        foreach ($finding in $findings) {
            $allResults += [PSCustomObject]@{
                GroupName      = $group.DisplayName
                GroupId        = $group.Id
                GroupType      = $groupType
                Status         = "IN USE"
                Category       = $finding.Category
                AssignedTo     = $finding.PolicyName
                AssignmentType = $finding.AssignmentType
            }
        }
    }
}

# ======================== SUMMARY ========================

$safeGroups = $resolvedGroups | Where-Object { -not $masterGroupMap.ContainsKey($_.Id.ToLower()) -or $masterGroupMap[$_.Id.ToLower()].Count -eq 0 }
$assignedGroups = $resolvedGroups | Where-Object { $masterGroupMap.ContainsKey($_.Id.ToLower()) -and $masterGroupMap[$_.Id.ToLower()].Count -gt 0 }

Write-Host ("=" * 100) -ForegroundColor Cyan
Write-Host "  SUMMARY" -ForegroundColor Cyan
Write-Host ("=" * 100) -ForegroundColor Cyan
Write-Host "  Total groups audited:   $($resolvedGroups.Count)" -ForegroundColor White
Write-Host "  Safe to delete:         $($safeGroups.Count)" -ForegroundColor Green
Write-Host "  Still in use:           $($assignedGroups.Count)" -ForegroundColor Red
Write-Host ""

if ($safeGroups.Count -gt 0) {
    Write-Host "  Groups safe to delete:" -ForegroundColor Green
    foreach ($g in $safeGroups) {
        Write-Host "    - $($g.DisplayName)" -ForegroundColor Green
    }
    Write-Host ""
}

if ($assignedGroups.Count -gt 0) {
    Write-Host "  Groups still in use:" -ForegroundColor Red
    foreach ($g in $assignedGroups) {
        $count = $masterGroupMap[$g.Id.ToLower()].Count
        Write-Host "    - $($g.DisplayName) ($count assignment(s))" -ForegroundColor Red
    }
    Write-Host ""
}

# ======================== EXPORT ========================

Write-Host "Would you like to export the results to CSV?" -ForegroundColor Cyan
$exportChoice = Read-Host "  Export? (Y/N)"

if ($exportChoice -match '^[Yy]') {
    Add-Type -AssemblyName System.Windows.Forms
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select folder to save Entra Group Audit report"
    $folderDialog.ShowNewFolderButton = $true

    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $fileName = "EntraGroupAudit_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').csv"
        $filePath = Join-Path $folderDialog.SelectedPath $fileName
        $allResults | Export-Csv -Path $filePath -NoTypeInformation
        Write-Host "Report saved to $filePath" -ForegroundColor Green
    }
    else {
        Write-Host "Export cancelled." -ForegroundColor Yellow
    }
}
