#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups

<#
.SYNOPSIS
    Menu-driven tool for querying Microsoft Graph user information.
.DESCRIPTION
    Connects to Microsoft Graph and provides a menu system to look up
    various user attributes by email address. Supports single or multiple
    email lookups per query.
#>

# ── Connection ────────────────────────────────────────────────────────
function Connect-GraphIfNeeded {
    $context = Get-MgContext
    if (-not $context) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All" -NoWelcome
        $context = Get-MgContext
        if (-not $context) {
            Write-Host "Failed to connect to Microsoft Graph. Exiting." -ForegroundColor Red
            exit 1
        }
    }
    Write-Host "Connected as: $($context.Account)`n" -ForegroundColor Green
}

# ── Email extraction ──────────────────────────────────────────────────
function Select-EmailsFromText {
    param([string]$Text)
    $pattern = '[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}'
    $found = [regex]::Matches($Text, $pattern)
    return $found | ForEach-Object { $_.Value.ToLower() } | Select-Object -Unique
}

function Read-MultiLineInput {
    param([string]$Prompt = "Paste your text (enter a blank line when done)")
    Write-Host $Prompt -ForegroundColor Yellow
    $lines = @()
    while ($true) {
        $line = Read-Host
        if ([string]::IsNullOrWhiteSpace($line)) { break }
        $lines += $line
    }
    return $lines -join "`n"
}

# ── Input helper ──────────────────────────────────────────────────────
function Get-EmailAddressInput {
    Write-Host ""
    Write-Host "  C. Comma-separated emails" -ForegroundColor White
    Write-Host "  P. Paste messy text (auto-extract emails)" -ForegroundColor White
    Write-Host ""
    $method = Read-Host "Input method (C/P)"

    switch ($method.ToUpper()) {
        'C' {
            Write-Host "Enter one or more email addresses (comma-separated):" -ForegroundColor Yellow
            $raw = Read-Host "Email(s)"
            if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
            return $raw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
        }
        'P' {
            $text = Read-MultiLineInput "Paste your text below (enter a blank line when done):"
            if ([string]::IsNullOrWhiteSpace($text)) { return @() }
            $emails = @(Select-EmailsFromText -Text $text)
            if ($emails.Count -eq 0) {
                Write-Host "No email addresses found in the pasted text.`n" -ForegroundColor Red
                return @()
            }
            Write-Host "`nFound $($emails.Count) email(s):" -ForegroundColor Green
            $emails | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
            Write-Host ""
            return $emails
        }
        default {
            Write-Host "Invalid input method.`n" -ForegroundColor Red
            return @()
        }
    }
}

# ── Menu options ──────────────────────────────────────────────────────
function Get-EnabledStatus {
    $emails = Get-EmailAddressInput
    if ($emails.Count -eq 0) { Write-Host "No emails provided.`n" -ForegroundColor Red; return }

    $results = foreach ($email in $emails) {
        try {
            $user = Get-MgUser -Filter "mail eq '$email' or userPrincipalName eq '$email'" `
                -Property DisplayName, Mail, UserPrincipalName, AccountEnabled `
                -ErrorAction Stop | Select-Object -First 1
            if ($user) {
                [PSCustomObject]@{
                    Email          = $email
                    DisplayName    = $user.DisplayName
                    UPN            = $user.UserPrincipalName
                    AccountEnabled = if ($user.AccountEnabled) { "Enabled" } else { "Disabled" }
                }
            }
            else {
                [PSCustomObject]@{
                    Email          = $email
                    DisplayName    = "NOT FOUND"
                    UPN            = "N/A"
                    AccountEnabled = "N/A"
                }
            }
        }
        catch {
            [PSCustomObject]@{
                Email          = $email
                DisplayName    = "ERROR"
                UPN            = "N/A"
                AccountEnabled = $_.Exception.Message
            }
        }
    }

    $results | Format-Table -AutoSize
}

function Search-GroupByName {
    Write-Host "Enter a group name or prefix to search for:" -ForegroundColor Yellow
    $search = Read-Host "Group name"
    if ([string]::IsNullOrWhiteSpace($search)) {
        Write-Host "No search term provided.`n" -ForegroundColor Red
        return
    }

    Write-Host "`nSearching for groups starting with '$search'..." -ForegroundColor Cyan

    try {
        $groups = Get-MgGroup -Filter "startsWith(displayName, '$search')" `
            -Property DisplayName, Id, Mail, Description, GroupTypes, MailEnabled, SecurityEnabled `
            -All -ErrorAction Stop
    }
    catch {
        Write-Host "Error searching groups: $($_.Exception.Message)`n" -ForegroundColor Red
        return
    }

    if ($groups.Count -eq 0) {
        Write-Host "No groups found matching '$search'.`n" -ForegroundColor Red
        return
    }

    Write-Host "Found $($groups.Count) group(s):`n" -ForegroundColor Green

    $results = foreach ($group in $groups) {
        $type = if ($group.GroupTypes -contains 'Unified') { 'Microsoft 365' }
        elseif ($group.SecurityEnabled -and $group.MailEnabled) { 'Mail-enabled Security' }
        elseif ($group.SecurityEnabled) { 'Security' }
        elseif ($group.MailEnabled) { 'Distribution' }
        else { 'Other' }

        [PSCustomObject]@{
            DisplayName = $group.DisplayName
            Type        = $type
            Mail        = if ($group.Mail) { $group.Mail } else { '-' }
            Description = if ($group.Description) { $group.Description.Substring(0, [Math]::Min(50, $group.Description.Length)) } else { '-' }
            ObjectId    = $group.Id
        }
    }

    $results | Sort-Object DisplayName | Format-Table -AutoSize
}

function Get-GroupMembers {
    Write-Host "Enter a group name or prefix to search for:" -ForegroundColor Yellow
    $search = Read-Host "Group name"
    if ([string]::IsNullOrWhiteSpace($search)) {
        Write-Host "No search term provided.`n" -ForegroundColor Red
        return
    }

    Write-Host "`nSearching for groups starting with '$search'..." -ForegroundColor Cyan

    try {
        $groups = @(Get-MgGroup -Filter "startsWith(displayName, '$search')" `
            -Property DisplayName, Id -All -ErrorAction Stop)
    }
    catch {
        Write-Host "Error searching groups: $($_.Exception.Message)`n" -ForegroundColor Red
        return
    }

    if ($groups.Count -eq 0) {
        Write-Host "No groups found matching '$search'.`n" -ForegroundColor Red
        return
    }

    # If multiple groups found, let the user pick one
    $selectedGroup = $null
    if ($groups.Count -eq 1) {
        $selectedGroup = $groups[0]
        Write-Host "Found group: $($selectedGroup.DisplayName)`n" -ForegroundColor Green
    }
    else {
        Write-Host "`nFound $($groups.Count) group(s):`n" -ForegroundColor Green
        for ($i = 0; $i -lt $groups.Count; $i++) {
            Write-Host "  $($i + 1). $($groups[$i].DisplayName)" -ForegroundColor White
        }
        Write-Host ""
        $pick = Read-Host "Select a group (1-$($groups.Count))"
        $index = 0
        if (-not [int]::TryParse($pick, [ref]$index) -or $index -lt 1 -or $index -gt $groups.Count) {
            Write-Host "Invalid selection.`n" -ForegroundColor Red
            return
        }
        $selectedGroup = $groups[$index - 1]
    }

    Write-Host "Fetching members of '$($selectedGroup.DisplayName)'..." -ForegroundColor Cyan

    try {
        $members = @(Get-MgGroupMember -GroupId $selectedGroup.Id -All -ErrorAction Stop)
    }
    catch {
        Write-Host "Error fetching members: $($_.Exception.Message)`n" -ForegroundColor Red
        return
    }

    if ($members.Count -eq 0) {
        Write-Host "No members found in this group.`n" -ForegroundColor Red
        return
    }

    # Resolve full user details for each member
    $memberResults = foreach ($member in $members) {
        try {
            $user = Get-MgUser -UserId $member.Id `
                -Property DisplayName, UserPrincipalName, OfficeLocation, AccountEnabled `
                -ErrorAction Stop
            [PSCustomObject]@{
                DisplayName       = if ($user.DisplayName) { $user.DisplayName } else { '-' }
                UserPrincipalName = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { '-' }
                Location          = if ($user.OfficeLocation) { $user.OfficeLocation } else { '-' }
                AccountEnabled    = if ($user.AccountEnabled) { "Enabled" } else { "Disabled" }
            }
        }
        catch {
            [PSCustomObject]@{
                DisplayName       = $member.Id
                UserPrincipalName = '-'
                Location          = '-'
                AccountEnabled    = '-'
            }
        }
    }

    $memberResults = $memberResults | Sort-Object DisplayName

    Write-Host "`nFound $($memberResults.Count) member(s) in '$($selectedGroup.DisplayName)'.`n" -ForegroundColor Green

    # Ask user how to output
    Write-Host "  D. Display to terminal" -ForegroundColor White
    Write-Host "  C. Copy UPNs to clipboard" -ForegroundColor White
    Write-Host "  B. Both" -ForegroundColor White
    Write-Host ""
    $output = Read-Host "Output method (D/C/B)"

    switch ($output.ToUpper()) {
        'D' {
            $memberResults | Format-Table -AutoSize
        }
        'C' {
            $upns = ($memberResults | Where-Object { $_.UserPrincipalName -ne '-' } | ForEach-Object { $_.UserPrincipalName }) -join "`n"
            $upns | Set-Clipboard
            Write-Host "UPNs copied to clipboard.`n" -ForegroundColor Green
        }
        'B' {
            $memberResults | Format-Table -AutoSize
            $upns = ($memberResults | Where-Object { $_.UserPrincipalName -ne '-' } | ForEach-Object { $_.UserPrincipalName }) -join "`n"
            $upns | Set-Clipboard
            Write-Host "UPNs also copied to clipboard.`n" -ForegroundColor Green
        }
        default {
            Write-Host "Invalid selection. Displaying to terminal by default.`n" -ForegroundColor Yellow
            $memberResults | Format-Table -AutoSize
        }
    }
}

# ── Menu ──────────────────────────────────────────────────────────────
function Show-Menu {
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "       Graph User Information Tool        " -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  1. Enabled Status"
    Write-Host "  2. Group Search"
    Write-Host "  3. Group Members"
    Write-Host ""
    Write-Host "  ── Utilities ──────────────────────────" -ForegroundColor DarkGray
    Write-Host "  E. Extract Emails from Text (copy to clipboard)"
    Write-Host ""
    Write-Host "  Q. Quit"
    Write-Host ""
}

# ── Main loop ─────────────────────────────────────────────────────────
Connect-GraphIfNeeded

do {
    Show-Menu
    $choice = Read-Host "Select an option"

    switch ($choice.ToUpper()) {
        '1' { Get-EnabledStatus }
        '2' { Search-GroupByName }
        '3' { Get-GroupMembers }
        'E' {
            $text = Read-MultiLineInput "Paste your text below (enter a blank line when done):"
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $emails = @(Select-EmailsFromText -Text $text)
                if ($emails.Count -gt 0) {
                    $csv = $emails -join ', '
                    $csv | Set-Clipboard
                    Write-Host "`nFound $($emails.Count) email(s) — copied to clipboard:" -ForegroundColor Green
                    $emails | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
                    Write-Host ""
                }
                else {
                    Write-Host "No email addresses found in the pasted text.`n" -ForegroundColor Red
                }
            }
        }
        'Q' { Write-Host "Goodbye!" -ForegroundColor Green }
        default { Write-Host "Invalid selection. Try again.`n" -ForegroundColor Red }
    }
} while ($choice.ToUpper() -ne 'Q')
