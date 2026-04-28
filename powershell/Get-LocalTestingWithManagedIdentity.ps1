<#
.SYNOPSIS
    Connects to Azure using managed identity when available, falling back to
    interactive sign-in for local development.

.DESCRIPTION
    Mirrors the behavior of DefaultAzureCredential from the Azure SDK: on an
    Azure host (VM, App Service, Function, Automation, etc.) it authenticates
    as the assigned managed identity; on a developer machine it prompts for
    interactive sign-in. The same script runs unchanged in both places.

.PARAMETER TenantId
    Optional tenant to sign in against. Useful when the signed-in user has
    guest access to multiple tenants.

.PARAMETER SubscriptionId
    Optional subscription to select after sign-in.

.PARAMETER UserAssignedIdentityClientId
    Client ID of a user-assigned managed identity. Omit to use the
    system-assigned identity.

.PARAMETER ResourceUrl
    Optional resource to fetch a token for after connecting (e.g.
    "https://vault.azure.net"). Returns the context plus a token.

.EXAMPLE
    .\Get-LocalTestingWithManagedIdentity.ps1

.EXAMPLE
    .\Get-LocalTestingWithManagedIdentity.ps1 -ResourceUrl "https://vault.azure.net"

.EXAMPLE
    .\Get-LocalTestingWithManagedIdentity.ps1 -UserAssignedIdentityClientId "00000000-0000-0000-0000-000000000000"
#>
[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$SubscriptionId,
    [string]$UserAssignedIdentityClientId,
    [string]$ResourceUrl,
    [switch]$ForceInteractive
)

if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
    throw "Az.Accounts module is required. Install with: Install-Module Az -Scope CurrentUser"
}

Import-Module Az.Accounts -ErrorAction Stop

function Test-AzureHostedEnvironment {
    # App Service, Functions, Container Apps, Arc — all expose IDENTITY_ENDPOINT
    if ($env:IDENTITY_ENDPOINT -or $env:MSI_ENDPOINT) { return $true }

    # Azure VM / VMSS — probe IMDS with a short timeout so we don't hang on laptops
    try {
        $imds = [System.Net.HttpWebRequest]::Create('http://169.254.169.254/metadata/instance?api-version=2021-02-01')
        $imds.Headers.Add('Metadata', 'true')
        $imds.Timeout = 1000
        $imds.ReadWriteTimeout = 1000
        $response = $imds.GetResponse()
        $response.Close()
        return $true
    }
    catch {
        return $false
    }
}

$connectParams = @{ ErrorAction = 'Stop' }
if ($TenantId)       { $connectParams['Tenant']       = $TenantId }
if ($SubscriptionId) { $connectParams['Subscription'] = $SubscriptionId }

$context = $null
$authMethod = $null

$tryMi = -not $ForceInteractive -and (Test-AzureHostedEnvironment)

if ($tryMi) {
    try {
        $miParams = $connectParams.Clone()
        $miParams['Identity'] = $true
        if ($UserAssignedIdentityClientId) {
            $miParams['AccountId'] = $UserAssignedIdentityClientId
        }

        Write-Verbose "Azure host detected — attempting managed identity sign-in."
        $context = Connect-AzAccount @miParams -WarningAction SilentlyContinue
        $authMethod = if ($UserAssignedIdentityClientId) { 'UserAssignedManagedIdentity' } else { 'SystemAssignedManagedIdentity' }
    }
    catch {
        Write-Warning "Managed identity sign-in failed: $($_.Exception.Message). Falling back to interactive."
        $context = Connect-AzAccount @connectParams
        $authMethod = 'Interactive'
    }
}
else {
    Write-Verbose "No Azure host detected (or -ForceInteractive supplied) — using interactive sign-in."
    $context = Connect-AzAccount @connectParams
    $authMethod = 'Interactive'
}

$accountInfo = Get-AzContext
Write-Host "Signed in via $authMethod as $($accountInfo.Account.Id) on tenant $($accountInfo.Tenant.Id)" -ForegroundColor Green

$result = [pscustomobject]@{
    AuthMethod   = $authMethod
    Account      = $accountInfo.Account.Id
    TenantId     = $accountInfo.Tenant.Id
    Subscription = $accountInfo.Subscription.Name
    Token        = $null
}

if ($ResourceUrl) {
    $token = Get-AzAccessToken -ResourceUrl $ResourceUrl -ErrorAction Stop
    $result.Token = $token.Token
    Write-Host "Acquired token for $ResourceUrl (expires $($token.ExpiresOn))" -ForegroundColor Green
}

return $result
