# Check if already connected to Microsoft Graph with required scopes
$requiredScopes = @("User.Read.All", "AuditLog.Read.All")
$context = Get-MgContext
if (-not $context) {
    Write-Host "Not connected to Microsoft Graph. Connecting..."
    Connect-MgGraph -Scopes $requiredScopes
}
else {
    $missingScopes = $requiredScopes | Where-Object { $_ -notin $context.Scopes }
    if ($missingScopes) {
        Write-Host "Connected but missing required scopes: $($missingScopes -join ', '). Reconnecting..."
        Disconnect-MgGraph | Out-Null
        Connect-MgGraph -Scopes $requiredScopes
    }
    else {
        Write-Host "Already connected to Microsoft Graph as $($context.Account)"
    }
}

# Fetch all guest users
Write-Progress -Activity "Guest User Audit" -Status "Fetching guest users..." -PercentComplete 0
$guests = Get-MgUser -Filter "userType eq 'Guest'" -All -Property "id,displayName,mail,userPrincipalName,createdDateTime,signInActivity,accountEnabled"

if ($guests.Count -eq 0) {
    Write-Host "No guest users found in the directory." -ForegroundColor Yellow
    return
}

Write-Host "Found $($guests.Count) guest users. Processing sign-in data..."

# Fetch audit logs for "Disable account" events to determine when users were disabled
Write-Progress -Activity "Guest User Audit" -Status "Fetching disable account audit logs..." -PercentComplete 5
$disableAuditLogs = @{}
try {
    $auditLogs = Get-MgAuditLogDirectoryAudit -Filter "activityDisplayName eq 'Disable account'" -All
    foreach ($log in $auditLogs) {
        $targetUser = $log.TargetResources | Where-Object { $_.Type -eq 'User' } | Select-Object -First 1
        if ($targetUser) {
            $targetId = $targetUser.Id
            # Keep the most recent disable event per user
            if (-not $disableAuditLogs.ContainsKey($targetId) -or $log.ActivityDateTime -gt $disableAuditLogs[$targetId]) {
                $disableAuditLogs[$targetId] = $log.ActivityDateTime
            }
        }
    }
}
catch {
    Write-Host "Note: Could not retrieve disable audit logs. 'Disabled Date' column will show 'Unknown'." -ForegroundColor Yellow
}

# Process each guest user
$cutoffDate = (Get-Date).AddDays(-30)
$i = 0
$total = $guests.Count
$results = foreach ($guest in $guests) {
    $i++
    $pct = [math]::Round(($i / $total) * 100)
    Write-Progress -Activity "Guest User Audit" -Status "Processing $i of $total - $($guest.DisplayName)" -PercentComplete $pct

    $lastSignIn = $guest.SignInActivity.LastSignInDateTime
    $lastNonInteractive = $guest.SignInActivity.LastNonInteractiveSignInDateTime

    # Use the most recent of interactive or non-interactive sign-in
    $mostRecentSignIn = @($lastSignIn, $lastNonInteractive) | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1

    if ($mostRecentSignIn) {
        $daysSinceSignIn = [math]::Round(((Get-Date) - $mostRecentSignIn).TotalDays)
        $inactive = $mostRecentSignIn -lt $cutoffDate
        $lastSignInDisplay = $mostRecentSignIn.ToString("yyyy-MM-dd HH:mm")
    }
    else {
        $daysSinceSignIn = "N/A"
        $inactive = $true
        $lastSignInDisplay = "Never"
    }

    $accountEnabled = $guest.AccountEnabled
    $disabledDate = if (-not $accountEnabled -and $disableAuditLogs.ContainsKey($guest.Id)) {
        $disableAuditLogs[$guest.Id].ToString("yyyy-MM-dd HH:mm")
    }
    elseif (-not $accountEnabled) {
        "Unknown"
    }
    else {
        "N/A"
    }

    [PSCustomObject]@{
        DisplayName    = $guest.DisplayName
        Email          = $guest.Mail
        UPN            = $guest.UserPrincipalName
        CreatedDate    = if ($guest.CreatedDateTime) { $guest.CreatedDateTime.ToString("yyyy-MM-dd") } else { "Unknown" }
        LastSignIn     = $lastSignInDisplay
        DaysSinceLogin = $daysSinceSignIn
        Inactive       = $inactive
        AccountStatus  = if ($accountEnabled) { "Enabled" } else { "Disabled" }
        DisabledDate   = $disabledDate
    }
}
Write-Progress -Activity "Guest User Audit" -Completed

# Sort results: inactive users first, then by display name
$results = $results | Sort-Object @{Expression = "Inactive"; Descending = $true }, DisplayName

# Display results in console with color coding
Write-Host ""
Write-Host ("{0,-30} {1,-35} {2,-20} {3,-10} {4,-12} {5}" -f "Display Name", "Email", "Last Sign-In", "Days Ago", "Account", "Status") -ForegroundColor Cyan
Write-Host ("-" * 130) -ForegroundColor Cyan

foreach ($user in $results) {
    $status = if ($user.Inactive) { "INACTIVE" } else { "Active" }

    # Color logic: Magenta if inactive AND disabled, Red if inactive only, Green if active
    $color = if ($user.Inactive -and $user.AccountStatus -eq "Disabled") {
        "Magenta"
    }
    elseif ($user.Inactive) {
        "Red"
    }
    else {
        "Green"
    }

    $line = "{0,-30} {1,-35} {2,-20} {3,-10} {4,-12} {5}" -f `
    ($user.DisplayName.Substring(0, [math]::Min(29, $user.DisplayName.Length))),
    $(if ($user.Email) { $user.Email.Substring(0, [math]::Min(34, $user.Email.Length)) } else { "N/A" }),
    $user.LastSignIn,
    $user.DaysSinceLogin,
    $user.AccountStatus,
    $status

    Write-Host $line -ForegroundColor $color
}

# Summary
$inactiveUsers = $results | Where-Object { $_.Inactive }
$activeCount = ($results | Where-Object { -not $_.Inactive }).Count
Write-Host ""
Write-Host "Summary: $($results.Count) total guests | $activeCount active | $($inactiveUsers.Count) inactive (30+ days or never signed in)" -ForegroundColor Yellow

# Export inactive users to CSV
if ($inactiveUsers.Count -gt 0) {
    Write-Host ""
    Write-Host "Select a folder to save the inactive guest users report..."

    Add-Type -AssemblyName System.Windows.Forms
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select folder to save inactive guest users report"
    $folderDialog.ShowNewFolderButton = $true

    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $fileName = "InactiveGuests_$(Get-Date -Format 'yyyy-MM-dd').csv"
        $filePath = Join-Path $folderDialog.SelectedPath $fileName
        $inactiveUsers | Select-Object DisplayName, Email, UPN, CreatedDate, LastSignIn, DaysSinceLogin, AccountStatus, DisabledDate |
        Export-Csv -Path $filePath -NoTypeInformation
        Write-Host "Inactive guest report saved to $filePath" -ForegroundColor Green
    }
    else {
        Write-Host "Save cancelled." -ForegroundColor Yellow
    }
}
else {
    Write-Host "No inactive guest users found. No CSV export needed." -ForegroundColor Green
}
