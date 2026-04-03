#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups

<#
.SYNOPSIS
    Exports job titles and office locations for members of one or more Entra groups.
.DESCRIPTION
    Resolves members of the specified groups, retrieves DisplayName, OfficeLocation,
    and JobTitle from Microsoft Graph, and adds a combined Location-Title column
    (e.g. "GA-IT Engineer"). Exports results to CSV.
#>

# ======================== CONFIGURATION ========================
# Enter group display names or object IDs (leave empty for interactive prompt)
$GroupsToQuery = @(
    # "SG-MyGroup"
    # "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
)
# ===============================================================

# ======================== GRAPH CONNECTION ========================
$context = Get-MgContext
if (-not $context) {
    Write-Host "Not connected to Microsoft Graph. Connecting..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All" -NoWelcome
    $context = Get-MgContext
    if (-not $context) {
        Write-Host "Failed to connect to Microsoft Graph. Exiting." -ForegroundColor Red
        exit 1
    }
}
else {
    Write-Host "Connected as: $($context.Account)" -ForegroundColor Green
}

# ======================== GROUP INPUT ========================
if ($GroupsToQuery.Count -eq 0) {
    Write-Host ""
    Write-Host "Enter group names or object IDs to query (one per line, blank line to finish):" -ForegroundColor Cyan
    while ($true) {
        $entry = Read-Host "  Group"
        if ([string]::IsNullOrWhiteSpace($entry)) { break }
        $GroupsToQuery += $entry.Trim()
    }
}

if ($GroupsToQuery.Count -eq 0) {
    Write-Host "No groups specified. Exiting." -ForegroundColor Yellow
    return
}

# ======================== RESOLVE GROUPS ========================
Write-Host ""
Write-Host "Resolving groups..." -ForegroundColor Cyan

$resolvedGroups = @()
foreach ($groupInput in $GroupsToQuery) {
    $group = $null
    if ($groupInput -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
        $group = Get-MgGroup -GroupId $groupInput -Property "id,displayName" -ErrorAction SilentlyContinue
    }
    if (-not $group) {
        $found = Get-MgGroup -Filter "displayName eq '$($groupInput -replace "'","''")'" -Property "id,displayName" -ErrorAction SilentlyContinue
        if ($found -is [array]) { $group = $found[0] } else { $group = $found }
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

# ======================== FETCH MEMBERS ========================
Write-Host ""
Write-Host "Fetching members..." -ForegroundColor Cyan

$allResults = @()
$seenUserIds = @{}

foreach ($group in $resolvedGroups) {
    Write-Host "  Processing: $($group.DisplayName)" -ForegroundColor White

    try {
        $members = @(Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop)
    }
    catch {
        Write-Host "  Error fetching members for $($group.DisplayName): $($_.Exception.Message)" -ForegroundColor Red
        continue
    }

    Write-Host "  Found $($members.Count) member(s)" -ForegroundColor DarkGray

    foreach ($member in $members) {
        # Skip duplicates when querying multiple groups
        if ($seenUserIds.ContainsKey($member.Id)) { continue }
        $seenUserIds[$member.Id] = $true

        try {
            $user = Get-MgUser -UserId $member.Id `
                -Property DisplayName, OfficeLocation, JobTitle `
                -ErrorAction Stop

            $location  = if ($user.OfficeLocation) { $user.OfficeLocation.Trim() } else { '' }
            $jobTitle  = if ($user.JobTitle) { $user.JobTitle.Trim() } else { '' }

            $combined = switch ($true) {
                { $location -and $jobTitle } { "$location-$jobTitle" }
                { $location }                { $location }
                { $jobTitle }                { $jobTitle }
                default                      { '' }
            }

            $allResults += [PSCustomObject]@{
                SourceGroup       = $group.DisplayName
                DisplayName       = if ($user.DisplayName) { $user.DisplayName } else { '-' }
                OfficeLocation    = if ($location) { $location } else { '-' }
                JobTitle          = if ($jobTitle) { $jobTitle } else { '-' }
                LocationTitle     = if ($combined) { $combined } else { '-' }
            }
        }
        catch {
            # Non-user objects (devices, service principals) will fail — skip silently
        }
    }
}

if ($allResults.Count -eq 0) {
    Write-Host ""
    Write-Host "No user records found. Exiting." -ForegroundColor Yellow
    return
}

$allResults = $allResults | Sort-Object SourceGroup, DisplayName

# ======================== DISPLAY SUMMARY ========================
Write-Host ""
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "  RESULTS  ($($allResults.Count) user(s))" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan
$allResults | Format-Table -AutoSize

# ======================== EXPORT ========================
Write-Host "Would you like to export the results to CSV?" -ForegroundColor Cyan
$exportChoice = Read-Host "  Export? (Y/N)"

if ($exportChoice -match '^[Yy]') {
    Add-Type -AssemblyName System.Windows.Forms
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select folder to save job titles report"
    $folderDialog.ShowNewFolderButton = $true

    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $fileName = "EntraUserJobTitles_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').csv"
        $filePath = Join-Path $folderDialog.SelectedPath $fileName
        $allResults | Export-Csv -Path $filePath -NoTypeInformation
        Write-Host "Report saved to: $filePath" -ForegroundColor Green
    }
    else {
        Write-Host "Export cancelled." -ForegroundColor Yellow
    }
}
