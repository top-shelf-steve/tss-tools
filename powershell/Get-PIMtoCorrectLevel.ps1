#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.Governance, Microsoft.Graph.Users

<#
.SYNOPSIS
    Activates an eligible Entra PIM role assignment for the signed-in user.
.DESCRIPTION
    Lists the current user's eligible PIM role assignments, prompts for a selection,
    collects ticket info + justification + duration, and submits a selfActivate
    request via Microsoft Graph. Works with policies that require ticket information.
    If the role policy also requires approval, the request will be submitted in a
    PendingApproval state and an approver must still action it.
#>

# ======================== CONFIGURATION ========================
# Optional: pre-select a role by display name (leave empty for interactive prompt)
$RoleDisplayName = ""

# Ticket system label used when submitting TicketInfo
$DefaultTicketSystem = "ServiceNow"

# Max duration defaults to 2 hours — policy may cap this lower
$DefaultDurationHours = 2
# ===============================================================

# ======================== GRAPH CONNECTION ========================
$requiredScopes = @(
    "RoleAssignmentSchedule.ReadWrite.Directory",
    "RoleEligibilitySchedule.Read.Directory",
    "User.Read"
)

$context = Get-MgContext
$needsReconnect = -not $context -or ($requiredScopes | Where-Object { $_ -notin $context.Scopes })

if ($needsReconnect) {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    $context = Get-MgContext
    if (-not $context) {
        Write-Host "Failed to connect to Microsoft Graph. Exiting." -ForegroundColor Red
        exit 1
    }
}
Write-Host "Connected as: $($context.Account)" -ForegroundColor Green

# ======================== RESOLVE CURRENT USER ========================
$me = Get-MgUser -UserId $context.Account -Property Id, DisplayName, UserPrincipalName -ErrorAction Stop
Write-Host "Signed-in user: $($me.DisplayName) ($($me.Id))" -ForegroundColor DarkGray

# ======================== FETCH ELIGIBLE ROLES ========================
Write-Host ""
Write-Host "Fetching eligible PIM role assignments..." -ForegroundColor Cyan

$eligible = @(
    Get-MgRoleManagementDirectoryRoleEligibilitySchedule `
        -Filter "principalId eq '$($me.Id)'" `
        -ExpandProperty "roleDefinition" `
        -All `
        -ErrorAction Stop
)

if ($eligible.Count -eq 0) {
    Write-Host "No eligible PIM role assignments found for this user. Exiting." -ForegroundColor Yellow
    return
}

# ======================== SELECT ROLE ========================
$target = $null

if ($RoleDisplayName) {
    $target = $eligible | Where-Object { $_.RoleDefinition.DisplayName -eq $RoleDisplayName } | Select-Object -First 1
    if (-not $target) {
        Write-Host "Configured role '$RoleDisplayName' not found in eligible assignments." -ForegroundColor Red
    }
}

if (-not $target) {
    Write-Host ""
    Write-Host "Eligible roles:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $eligible.Count; $i++) {
        $scopeLabel = if ($eligible[$i].DirectoryScopeId -eq "/") { "Directory" } else { $eligible[$i].DirectoryScopeId }
        Write-Host ("  [{0}] {1}  ({2})" -f ($i + 1), $eligible[$i].RoleDefinition.DisplayName, $scopeLabel)
    }

    $choice = Read-Host "Select role number"
    if (-not ($choice -as [int]) -or [int]$choice -lt 1 -or [int]$choice -gt $eligible.Count) {
        Write-Host "Invalid selection. Exiting." -ForegroundColor Red
        return
    }
    $target = $eligible[[int]$choice - 1]
}

Write-Host ""
Write-Host "Activating: $($target.RoleDefinition.DisplayName)" -ForegroundColor Green

# ======================== COLLECT ACTIVATION INPUT ========================
$ticketNumber = Read-Host "Ticket number"
if ([string]::IsNullOrWhiteSpace($ticketNumber)) {
    Write-Host "Ticket number is required by policy. Exiting." -ForegroundColor Red
    return
}

$ticketSystem = Read-Host "Ticket system [$DefaultTicketSystem]"
if ([string]::IsNullOrWhiteSpace($ticketSystem)) { $ticketSystem = $DefaultTicketSystem }

$justification = Read-Host "Justification"
if ([string]::IsNullOrWhiteSpace($justification)) {
    Write-Host "Justification is required. Exiting." -ForegroundColor Red
    return
}

$durationInput = Read-Host "Duration in hours [$DefaultDurationHours]"
$hours = if ([string]::IsNullOrWhiteSpace($durationInput)) { $DefaultDurationHours } else { [int]$durationInput }

# ======================== SUBMIT ACTIVATION ========================
$params = @{
    Action           = "selfActivate"
    PrincipalId      = $me.Id
    RoleDefinitionId = $target.RoleDefinitionId
    DirectoryScopeId = $target.DirectoryScopeId
    Justification    = $justification
    ScheduleInfo     = @{
        StartDateTime = Get-Date
        Expiration    = @{
            Type     = "AfterDuration"
            Duration = "PT${hours}H"
        }
    }
    TicketInfo       = @{
        TicketNumber = $ticketNumber
        TicketSystem = $ticketSystem
    }
}

Write-Host ""
Write-Host "Submitting activation request..." -ForegroundColor Cyan

try {
    $request = New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $params -ErrorAction Stop
}
catch {
    Write-Host "Activation request failed: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# ======================== RESULT ========================
Write-Host ""
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "  ACTIVATION RESULT" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "  Role:    $($target.RoleDefinition.DisplayName)"
Write-Host "  Request: $($request.Id)"
Write-Host "  Status:  $($request.Status)" -ForegroundColor Yellow

switch ($request.Status) {
    "Provisioned"     { Write-Host "  Role is now active." -ForegroundColor Green }
    "Granted"         { Write-Host "  Role is now active." -ForegroundColor Green }
    "PendingApproval" { Write-Host "  Awaiting approver action. Check the PIM portal or Teams app." -ForegroundColor Yellow }
    default           { Write-Host "  Review the request in the Entra PIM portal for details." -ForegroundColor DarkGray }
}
