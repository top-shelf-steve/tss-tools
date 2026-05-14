<#
.SYNOPSIS
    Adds a user to an Entra ID security group and assigns them to a personal AVD session host.

.DESCRIPTION
    Performs two actions for onboarding a user to a personal Azure Virtual Desktop machine:
      1. Adds the user as a member of a specified Entra ID group (used for AVD app-group / desktop assignment).
      2. Assigns the user to a specific personal session host in a host pool.

    Resource group, host pool, and Entra group all have defaults defined in the
    DEFAULTS region below — override any of them at call time via parameters.

.PARAMETER UserPrincipalName
    UPN of the user to add (e.g. jdoe@contoso.com).

.PARAMETER SessionHostName
    The session host (VM) name as it appears in the host pool, e.g. "avd-prod-01.contoso.local".
    Use Get-AzWvdSessionHost to list valid names if unsure.

.PARAMETER EntraGroupId
    Object ID of the Entra group to add the user to. Overrides the default below.

.PARAMETER EntraGroupName
    Display name of the Entra group. Used only if -EntraGroupId is not supplied.
    Overrides the default below.

.PARAMETER ResourceGroupName
    Resource group containing the AVD host pool. Overrides the default below.

.PARAMETER HostPoolName
    Name of the AVD host pool containing the personal session host. Overrides the default below.

.PARAMETER SubscriptionId
    Optional Azure subscription ID. If omitted, uses the current Az context.

.EXAMPLE
    .\Add-UserToPersonalAVDMachine.ps1 -UserPrincipalName "jdoe@contoso.com" -SessionHostName "avd-prod-01.contoso.local"
    # Uses the default Entra group, resource group, and host pool.

.EXAMPLE
    .\Add-UserToPersonalAVDMachine.ps1 -UserPrincipalName "jdoe@contoso.com" -SessionHostName "avd-prod-01.contoso.local" -EntraGroupName "AVD-Special-Users"
    # Adds to a different Entra group by display name.

.EXAMPLE
    .\Add-UserToPersonalAVDMachine.ps1 -UserPrincipalName "jdoe@contoso.com" -SessionHostName "avd-dev-05.contoso.local" -ResourceGroupName "rg-avd-dev" -HostPoolName "hp-avd-dev"
    # Overrides resource group and host pool for a dev environment.

.NOTES
    Required modules: Az.Accounts, Az.DesktopVirtualization, Microsoft.Graph.Groups, Microsoft.Graph.Users
#>

#Requires -Modules Az.Accounts, Az.DesktopVirtualization, Microsoft.Graph.Groups, Microsoft.Graph.Users

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $true)]
    [string]$SessionHostName,

    [Parameter(Mandatory = $false)]
    [string]$EntraGroupId,

    [Parameter(Mandatory = $false)]
    [string]$EntraGroupName,

    [Parameter(Mandatory = $false)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $false)]
    [string]$HostPoolName,

    [Parameter(Mandatory = $false)]
    [string]$SubscriptionId
)

$ErrorActionPreference = 'Stop'

# ========================
# DEFAULTS — edit these to match your most-common values.
# Anything passed via parameters overrides what's set here.
# ========================
$DefaultEntraGroupId     = ""                          # e.g. "00000000-0000-0000-0000-000000000000"
$DefaultEntraGroupName   = "AVD-Personal-Users"        # used only if DefaultEntraGroupId is blank
$DefaultResourceGroup    = "rg-avd-prod"
$DefaultHostPoolName     = "hp-avd-prod"

# Resolve effective values: parameter > default
if (-not $EntraGroupId)     { $EntraGroupId     = $DefaultEntraGroupId }
if (-not $EntraGroupName)   { $EntraGroupName   = $DefaultEntraGroupName }
if (-not $ResourceGroupName){ $ResourceGroupName= $DefaultResourceGroup }
if (-not $HostPoolName)     { $HostPoolName     = $DefaultHostPoolName }

if (-not $EntraGroupId -and -not $EntraGroupName) {
    throw "No Entra group specified. Set a default in the DEFAULTS region or pass -EntraGroupId / -EntraGroupName."
}

# ========================
# AUTHENTICATION
# ========================
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  Add User to Personal AVD Machine" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

Write-Host "`nAuthenticating to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "GroupMember.ReadWrite.All", "User.Read.All" -NoWelcome
Write-Host "  Connected to Graph" -ForegroundColor Green

Write-Host "`nAuthenticating to Azure..." -ForegroundColor Cyan
$azContext = Get-AzContext
if (-not $azContext) {
    Connect-AzAccount | Out-Null
    $azContext = Get-AzContext
}
if ($SubscriptionId) {
    Set-AzContext -SubscriptionId $SubscriptionId | Out-Null
    $azContext = Get-AzContext
}
Write-Host "  Subscription: $($azContext.Subscription.Name) ($($azContext.Subscription.Id))" -ForegroundColor Green

# ========================
# RESOLVE USER
# ========================
Write-Host "`nResolving user..." -ForegroundColor Cyan
$user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
Write-Host "  Found user: $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor Green
Write-Host "  Object ID: $($user.Id)" -ForegroundColor Gray

# ========================
# RESOLVE ENTRA GROUP
# ========================
Write-Host "`nResolving Entra group..." -ForegroundColor Cyan
if ($EntraGroupId) {
    $group = Get-MgGroup -GroupId $EntraGroupId -ErrorAction Stop
}
else {
    $escaped = $EntraGroupName.Replace("'", "''")
    $group = Get-MgGroup -Filter "displayName eq '$escaped'" -ErrorAction Stop | Select-Object -First 1
    if (-not $group) {
        throw "No Entra group found with displayName '$EntraGroupName'."
    }
}
Write-Host "  Found group: $($group.DisplayName)" -ForegroundColor Green
Write-Host "  Object ID: $($group.Id)" -ForegroundColor Gray

# ========================
# ADD USER TO GROUP
# ========================
Write-Host "`nAdding user to group..." -ForegroundColor Cyan
$existingMember = Get-MgGroupMember -GroupId $group.Id -All |
    Where-Object { $_.Id -eq $user.Id }

if ($existingMember) {
    Write-Host "  User is already a member of '$($group.DisplayName)'. Skipping." -ForegroundColor Yellow
}
else {
    New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.Id
    Write-Host "  Added $($user.UserPrincipalName) to '$($group.DisplayName)'" -ForegroundColor Green
}

# ========================
# VALIDATE HOST POOL & SESSION HOST
# ========================
Write-Host "`nValidating host pool and session host..." -ForegroundColor Cyan
$hostPool = Get-AzWvdHostPool -ResourceGroupName $ResourceGroupName -Name $HostPoolName -ErrorAction Stop
Write-Host "  Host pool: $($hostPool.Name) ($($hostPool.HostPoolType))" -ForegroundColor Green

if ($hostPool.HostPoolType -ne "Personal") {
    Write-Warning "Host pool '$HostPoolName' is type '$($hostPool.HostPoolType)', not 'Personal'. User assignment only applies to personal host pools."
}

$sessionHost = Get-AzWvdSessionHost -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName -Name $SessionHostName -ErrorAction Stop
Write-Host "  Session host: $($sessionHost.Name.Split('/')[-1])" -ForegroundColor Green
if ($sessionHost.AssignedUser) {
    Write-Host "  Currently assigned to: $($sessionHost.AssignedUser)" -ForegroundColor Gray
}

# ========================
# ASSIGN USER TO SESSION HOST
# ========================
Write-Host "`nAssigning user to session host..." -ForegroundColor Cyan
$updated = Update-AzWvdSessionHost `
    -ResourceGroupName $ResourceGroupName `
    -HostPoolName $HostPoolName `
    -Name $SessionHostName `
    -AssignedUser $UserPrincipalName `
    -ErrorAction Stop

Write-Host "  Assigned $UserPrincipalName -> $SessionHostName" -ForegroundColor Green

# ========================
# SUMMARY
# ========================
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "  Done." -ForegroundColor Green
Write-Host "  User:          $($user.UserPrincipalName)" -ForegroundColor White
Write-Host "  Entra Group:   $($group.DisplayName)" -ForegroundColor White
Write-Host "  Host Pool:     $HostPoolName ($ResourceGroupName)" -ForegroundColor White
Write-Host "  Session Host:  $SessionHostName" -ForegroundColor White
Write-Host "  Assigned User: $($updated.AssignedUser)" -ForegroundColor White
Write-Host "==========================================" -ForegroundColor Cyan
