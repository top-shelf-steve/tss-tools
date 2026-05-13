<#
.SYNOPSIS
    Adds a service principal as an owner of an app registration.
    By default, makes an app registration its own owner (so the app can manage
    its own credentials when authenticated as itself).

.PARAMETER AppId
    The Application (client) ID of the app registration that should be self-owned.
    The script resolves both the application object and its service principal from this.

.PARAMETER OwnerAppId
    Optional. The Application (client) ID of a *different* service principal to add
    as owner. If omitted, the app's own service principal is used (self-ownership).

.PARAMETER UpdateSecret
    If set, generates a new client secret on the target app registration after
    ensuring ownership. The plaintext secret value is written to the host once
    and returned on the pipeline — capture it immediately, it cannot be retrieved later.

.PARAMETER SecretDisplayName
    Optional display name for the new secret. Defaults to "Rotated <yyyy-MM-dd HH:mm>".

.PARAMETER SecretValidDays
    Optional lifetime for the new secret, in days. Default 365.

.EXAMPLE
    .\Add-OwnerToServicePrincpal.ps1 -AppId "11111111-1111-1111-1111-111111111111"
    # Makes the app a self-owner.

.EXAMPLE
    .\Add-OwnerToServicePrincpal.ps1 -AppId "<target-app>" -OwnerAppId "<owner-app>"
    # Adds the OwnerApp's service principal as owner of the target app.

.EXAMPLE
    .\Add-OwnerToServicePrincpal.ps1 -AppId "<app>" -UpdateSecret
    # Ensures self-ownership and generates a fresh client secret valid for 1 year.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AppId,

    [Parameter(Mandatory = $false)]
    [string]$OwnerAppId,

    [Parameter(Mandatory = $false)]
    [switch]$UpdateSecret,

    [Parameter(Mandatory = $false)]
    [string]$SecretDisplayName,

    [Parameter(Mandatory = $false)]
    [int]$SecretValidDays = 365
)

$ErrorActionPreference = 'Stop'

# Ensure the Graph SDK module is available
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Applications)) {
    Write-Host "Installing Microsoft.Graph.Applications module..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph.Applications

Connect-MgGraph -Scopes "Application.ReadWrite.All" -NoWelcome

# Resolve the target application (the one receiving the owner)
$application = Get-MgApplication -Filter "appId eq '$AppId'"
if (-not $application) {
    throw "No app registration found with AppId '$AppId'."
}
$applicationObjectId = $application.Id
Write-Host "Target application: $($application.DisplayName) (ObjectId: $applicationObjectId)"

# Resolve the owner service principal
$ownerLookupAppId = if ($OwnerAppId) { $OwnerAppId } else { $AppId }
$ownerSp = Get-MgServicePrincipal -Filter "appId eq '$ownerLookupAppId'"
if (-not $ownerSp) {
    throw "No service principal found with AppId '$ownerLookupAppId'. " +
          "If this is a brand-new app registration, you may need to create the " +
          "enterprise application (service principal) first."
}
$ownerServicePrincipalObjectId = $ownerSp.Id
Write-Host "Owner service principal: $($ownerSp.DisplayName) (ObjectId: $ownerServicePrincipalObjectId)"

# Skip the owner-add if the SP is already an owner (the API returns 400 on duplicates)
$existingOwners = Get-MgApplicationOwner -ApplicationId $applicationObjectId -All
if ($existingOwners.Id -contains $ownerServicePrincipalObjectId) {
    Write-Host "Service principal is already an owner. Skipping ownership step." -ForegroundColor Green
}
else {
    $ownerRef = @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$ownerServicePrincipalObjectId"
    }
    New-MgApplicationOwnerByRef -ApplicationId $applicationObjectId -BodyParameter $ownerRef
    Write-Host "Owner added successfully." -ForegroundColor Green
}

if ($UpdateSecret) {
    if (-not $SecretDisplayName) {
        $SecretDisplayName = "Rotated $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    }

    $passwordCredential = @{
        displayName   = $SecretDisplayName
        endDateTime   = (Get-Date).AddDays($SecretValidDays).ToUniversalTime()
    }

    Write-Host "Generating new client secret '$SecretDisplayName' (valid $SecretValidDays days)..."
    $newSecret = Add-MgApplicationPassword -ApplicationId $applicationObjectId -PasswordCredential $passwordCredential

    Write-Host ""
    Write-Host "=== NEW CLIENT SECRET (copy now — it will not be shown again) ===" -ForegroundColor Yellow
    Write-Host "KeyId       : $($newSecret.KeyId)"
    Write-Host "DisplayName : $($newSecret.DisplayName)"
    Write-Host "Expires     : $($newSecret.EndDateTime)"
    Write-Host "SecretText  : $($newSecret.SecretText)" -ForegroundColor Cyan
    Write-Host ""

    # Also emit on the pipeline so callers can capture it programmatically
    [pscustomobject]@{
        AppId       = $AppId
        KeyId       = $newSecret.KeyId
        DisplayName = $newSecret.DisplayName
        EndDateTime = $newSecret.EndDateTime
        SecretText  = $newSecret.SecretText
    }
}
