<#
.SYNOPSIS
    Renews an Azure Virtual Desktop host pool registration token and stores it in Azure Key Vault.

.DESCRIPTION
    Generates a new registration token for a specified AVD host pool and updates
    an existing Azure Key Vault secret with the new token value.

.PARAMETER HostPoolName
    The name of the AVD host pool to renew the token for.

.PARAMETER ResourceGroupName
    The resource group containing the host pool.

.PARAMETER KeyVaultName
    The name of the Azure Key Vault containing the secret to update.

.PARAMETER SecretName
    The name of the existing Key Vault secret to update with the new token.

.PARAMETER SubscriptionId
    Azure subscription ID. If omitted, uses the current Az context.

.PARAMETER UseManagedIdentity
    Authenticate using Managed Identity instead of interactive login.
    Use this when running as an Azure Automation Runbook.

.PARAMETER ExpirationHours
    Number of hours until the new token expires. Defaults to 24. Maximum is 720 (30 days).

.EXAMPLE
    .\Renew-AVDHostPoolToken.ps1 -HostPoolName "hp-avd-prod" -ResourceGroupName "rg-avd-prod" -KeyVaultName "kv-avd-prod" -SecretName "avd-host-pool-token"

.EXAMPLE
    .\Renew-AVDHostPoolToken.ps1 -HostPoolName "hp-avd-prod" -ResourceGroupName "rg-avd-prod" -KeyVaultName "kv-avd-prod" -SecretName "avd-host-pool-token" -ExpirationHours 48

.EXAMPLE
    .\Renew-AVDHostPoolToken.ps1 -HostPoolName "hp-avd-prod" -ResourceGroupName "rg-avd-prod" -KeyVaultName "kv-avd-prod" -SecretName "avd-host-pool-token" -UseManagedIdentity

.NOTES
    Required modules: Az.Accounts, Az.DesktopVirtualization, Az.KeyVault
#>

#Requires -Modules Az.Accounts, Az.DesktopVirtualization, Az.KeyVault

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$HostPoolName,

    [Parameter(Mandatory = $true)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $true)]
    [string]$KeyVaultName,

    [Parameter(Mandatory = $true)]
    [string]$SecretName,

    [Parameter(Mandatory = $false)]
    [string]$SubscriptionId,

    [Parameter(Mandatory = $false)]
    [switch]$UseManagedIdentity,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 720)]
    [int]$ExpirationHours = 24
)

# ========================
# AUTHENTICATION
# ========================
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  AVD Host Pool Token Renewal" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

Write-Host "`nAuthenticating..." -ForegroundColor Cyan
if ($UseManagedIdentity) {
    Connect-AzAccount -Identity | Out-Null
    Write-Host "  Connected via Managed Identity" -ForegroundColor Green
}
else {
    Connect-AzAccount | Out-Null
    Write-Host "  Connected interactively" -ForegroundColor Green
}

if ($SubscriptionId) {
    Set-AzContext -SubscriptionId $SubscriptionId | Out-Null
}

$context = Get-AzContext
Write-Host "  Subscription: $($context.Subscription.Name) ($($context.Subscription.Id))" -ForegroundColor Green

# ========================
# VALIDATE HOST POOL
# ========================
Write-Host "`nValidating host pool..." -ForegroundColor Cyan

$hostPool = Get-AzWvdHostPool -ResourceGroupName $ResourceGroupName -Name $HostPoolName -ErrorAction Stop
if (-not $hostPool) {
    Write-Error "Host pool '$HostPoolName' not found in resource group '$ResourceGroupName'."
    exit 1
}
Write-Host "  Found host pool: $($hostPool.Name)" -ForegroundColor Green
Write-Host "  Type: $($hostPool.HostPoolType)" -ForegroundColor Gray
Write-Host "  Resource Group: $ResourceGroupName" -ForegroundColor Gray

# ========================
# VALIDATE KEY VAULT SECRET EXISTS
# ========================
Write-Host "`nValidating Key Vault secret..." -ForegroundColor Cyan

try {
    $existingSecret = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name $SecretName -ErrorAction Stop
    Write-Host "  Found existing secret: $SecretName" -ForegroundColor Green
    Write-Host "  Key Vault: $KeyVaultName" -ForegroundColor Gray
}
catch {
    Write-Error "Secret '$SecretName' not found in Key Vault '$KeyVaultName'. Ensure the secret already exists."
    exit 1
}

# ========================
# RENEW REGISTRATION TOKEN
# ========================
Write-Host "`nRenewing registration token..." -ForegroundColor Cyan

$expirationTime = (Get-Date).ToUniversalTime().AddHours($ExpirationHours)
Write-Host "  Token expiration: $($expirationTime.ToString('yyyy-MM-dd HH:mm:ss')) UTC" -ForegroundColor Gray

$token = New-AzWvdRegistrationInfo -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName -ExpirationTime $expirationTime -ErrorAction Stop

if (-not $token.Token) {
    Write-Error "Failed to generate registration token. The token value is empty."
    exit 1
}

Write-Host "  Registration token generated successfully" -ForegroundColor Green

# ========================
# UPDATE KEY VAULT SECRET
# ========================
Write-Host "`nUpdating Key Vault secret..." -ForegroundColor Cyan

$secureToken = ConvertTo-SecureString -String $token.Token -AsPlainText -Force
Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name $SecretName -SecretValue $secureToken -ErrorAction Stop | Out-Null

Write-Host "  Secret '$SecretName' updated successfully in Key Vault '$KeyVaultName'" -ForegroundColor Green

# ========================
# SUMMARY
# ========================
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "  Token renewal complete!" -ForegroundColor Green
Write-Host "  Host Pool:   $HostPoolName" -ForegroundColor White
Write-Host "  Key Vault:   $KeyVaultName" -ForegroundColor White
Write-Host "  Secret:      $SecretName" -ForegroundColor White
Write-Host "  Expires:     $($expirationTime.ToString('yyyy-MM-dd HH:mm:ss')) UTC ($ExpirationHours hours)" -ForegroundColor White
Write-Host "==========================================" -ForegroundColor Cyan
