#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups

<#
.SYNOPSIS
    Picks a random winner from a Microsoft Graph group.
.DESCRIPTION
    Connects to Microsoft Graph, searches for a group by name,
    fetches its members, and randomly selects a lucky winner.
    Set $DefaultGroupName below to skip the search step.
#>

# ── Default group (set this to skip the search) ──────────────────────
# Leave empty to always be prompted for a group name.
$DefaultGroupName = ""

# ── Connection ────────────────────────────────────────────────────────
function Connect-GraphIfNeeded {
    $context = Get-MgContext
    if (-not $context) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes "Group.Read.All", "GroupMember.Read.All" -NoWelcome
        $context = Get-MgContext
        if (-not $context) {
            Write-Host "Failed to connect to Microsoft Graph. Exiting." -ForegroundColor Red
            exit 1
        }
    }
    Write-Host "Connected as: $($context.Account)`n" -ForegroundColor Green
}

# ── Group lookup ──────────────────────────────────────────────────────
function Find-GroupByName {
    param([string]$Search)

    Write-Host "`nSearching for groups starting with '$Search'..." -ForegroundColor Cyan

    try {
        $groups = @(Get-MgGroup -Filter "startsWith(displayName, '$Search')" `
                -Property DisplayName, Id -All -ErrorAction Stop)
    }
    catch {
        Write-Host "Error searching groups: $($_.Exception.Message)`n" -ForegroundColor Red
        return $null
    }

    if ($groups.Count -eq 0) {
        Write-Host "No groups found matching '$Search'.`n" -ForegroundColor Red
        return $null
    }

    if ($groups.Count -eq 1) {
        Write-Host "Found group: $($groups[0].DisplayName)`n" -ForegroundColor Green
        return $groups[0]
    }

    Write-Host "`nFound $($groups.Count) group(s):`n" -ForegroundColor Green
    for ($i = 0; $i -lt $groups.Count; $i++) {
        Write-Host "  $($i + 1). $($groups[$i].DisplayName)" -ForegroundColor White
    }
    Write-Host ""
    $pick = Read-Host "Select a group (1-$($groups.Count))"
    $index = 0
    if (-not [int]::TryParse($pick, [ref]$index) -or $index -lt 1 -or $index -gt $groups.Count) {
        Write-Host "Invalid selection.`n" -ForegroundColor Red
        return $null
    }
    return $groups[$index - 1]
}

function Select-Group {
    Write-Host "Enter a group name or prefix to search for:" -ForegroundColor Yellow
    $search = Read-Host "Group name"
    if ([string]::IsNullOrWhiteSpace($search)) {
        Write-Host "No search term provided.`n" -ForegroundColor Red
        return $null
    }
    return Find-GroupByName -Search $search
}

# ── Winner selection ──────────────────────────────────────────────────
function Show-LuckyWinner {
    param([object]$Group)

    Write-Host "Fetching members of '$($Group.DisplayName)'..." -ForegroundColor Cyan

    try {
        $members = @(Get-MgGroupMember -GroupId $Group.Id -All -ErrorAction Stop)
    }
    catch {
        Write-Host "Error fetching members: $($_.Exception.Message)`n" -ForegroundColor Red
        return
    }

    if ($members.Count -eq 0) {
        Write-Host "No members found in this group.`n" -ForegroundColor Red
        return
    }

    # Resolve display names
    $people = foreach ($member in $members) {
        try {
            $user = Get-MgUser -UserId $member.Id -Property DisplayName -ErrorAction Stop
            $user.DisplayName
        }
        catch { $null }
    }
    $people = @($people | Where-Object { $_ })

    if ($people.Count -eq 0) {
        Write-Host "Could not resolve any members.`n" -ForegroundColor Red
        return
    }

    Write-Host "`nGroup '$($Group.DisplayName)' has $($people.Count) member(s).`n" -ForegroundColor Green

    # Drumroll
    Write-Host "  Spinning the wheel..." -ForegroundColor DarkGray
    $spinChars = @('|', '/', '-', '\')
    for ($i = 0; $i -lt 12; $i++) {
        $randomName = $people | Get-Random
        Write-Host "`r  $($spinChars[$i % 4]) $randomName     " -NoNewline -ForegroundColor DarkGray
        Start-Sleep -Milliseconds 150
    }
    Write-Host "`r                                                          " -NoNewline

    # Pick the actual winner
    $winner = $people | Get-Random

    Write-Host ""
    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Yellow
    Write-Host "  *                                          *" -ForegroundColor Yellow
    Write-Host "  *        AND THE LUCKY WINNER IS...        *" -ForegroundColor Yellow
    Write-Host "  *                                          *" -ForegroundColor Yellow
    Write-Host "  *   >> $($winner.PadRight(33)) <<" -ForegroundColor Magenta
    Write-Host "  *                                          *" -ForegroundColor Yellow
    Write-Host "  *      Congratulations! You've been        *" -ForegroundColor Cyan
    Write-Host "  *           chosen by destiny!             *" -ForegroundColor Cyan
    Write-Host "  *                                          *" -ForegroundColor Yellow
    Write-Host "  ============================================" -ForegroundColor Yellow
    Write-Host ""
}

# ── Main ──────────────────────────────────────────────────────────────
Connect-GraphIfNeeded

# Resolve default group once at startup if configured
$defaultGroup = $null
if (-not [string]::IsNullOrWhiteSpace($DefaultGroupName)) {
    Write-Host "Default group configured: '$DefaultGroupName'" -ForegroundColor Cyan
    $defaultGroup = Find-GroupByName -Search $DefaultGroupName
    if (-not $defaultGroup) {
        Write-Host "Could not find default group. You can still search manually.`n" -ForegroundColor Yellow
    }
}

do {
    if ($defaultGroup) {
        Write-Host "  D. Use default group ($($defaultGroup.DisplayName))" -ForegroundColor White
        Write-Host "  S. Search for a different group" -ForegroundColor White
        Write-Host ""
        $method = Read-Host "Choice (D/S)"

        if ($method.ToUpper() -eq 'S') {
            $group = Select-Group
        }
        else {
            $group = $defaultGroup
        }
    }
    else {
        $group = Select-Group
    }

    if ($group) {
        Show-LuckyWinner -Group $group
    }

    Write-Host ""
    $again = Read-Host "Pick another winner? (Y/N)"
} while ($again -and $again.ToUpper() -eq 'Y')

Write-Host "Thanks for playing! Goodbye!" -ForegroundColor Green
