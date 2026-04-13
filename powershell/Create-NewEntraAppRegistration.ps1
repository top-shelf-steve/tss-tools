<#
.SYNOPSIS
    Creates a new Entra ID (Azure AD) App Registration and its associated Enterprise Application.

.DESCRIPTION
    Uses the Microsoft.Graph PowerShell module to create a new App Registration and a
    corresponding Service Principal (Enterprise Application), optionally setting a redirect
    URI, platform type, and/or an owner.

    Note: New-MgApplication only creates the App Registration object. This script also calls
    New-MgServicePrincipal to create the Enterprise Application, mirroring portal behavior.

.PARAMETER DisplayName
    The display name for the new App Registration.

.PARAMETER RedirectUri
    Optional. A redirect URI to associate with the app.

.PARAMETER PlatformType
    Optional. The redirect URI platform type. Defaults to "Web".
    - Web         : Standard web app (server-side). Sets Web.RedirectUris.
    - SPA         : Single Page Application. Sets Spa.RedirectUris.
    - PublicClient: Mobile/desktop app. Sets PublicClient.RedirectUris.

.PARAMETER OwnerUPN
    Optional. The UPN or Object ID of a user to add as an owner of the app.

.EXAMPLE
    .\Create-NewEntraAppRegistration.ps1 -DisplayName "MyApp"

.EXAMPLE
    .\Create-NewEntraAppRegistration.ps1 -DisplayName "MyApp" `
        -RedirectUri "https://myapp.example.com/auth/callback" `
        -OwnerUPN "alice@contoso.com"

.EXAMPLE
    .\Create-NewEntraAppRegistration.ps1 -DisplayName "MySPA" `
        -RedirectUri "https://myapp.example.com" `
        -PlatformType SPA

.EXAMPLE
    .\Create-NewEntraAppRegistration.ps1 -DisplayName "MyMobileApp" `
        -RedirectUri "myapp://auth" `
        -PlatformType PublicClient
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$DisplayName,

    [Parameter(Mandatory = $false)]
    [string]$RedirectUri,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Web", "SPA", "PublicClient")]
    [string]$PlatformType = "Web",

    [Parameter(Mandatory = $false)]
    [string]$OwnerUPN
)

#region Prerequisites
$requiredScopes = @("Application.ReadWrite.All", "User.Read.All")

try {
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
}
catch {
    Write-Error "Microsoft.Graph modules not found. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    exit 1
}

try {
    Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}
#endregion

#region Build App Registration Body
$appBody = @{
    DisplayName    = $DisplayName
    SignInAudience = "AzureADMyOrg"
}

if ($RedirectUri) {
    switch ($PlatformType) {
        "Web" {
            $appBody.Web = @{ RedirectUris = @($RedirectUri) }
        }
        "SPA" {
            $appBody.Spa = @{ RedirectUris = @($RedirectUri) }
        }
        "PublicClient" {
            $appBody.PublicClient = @{ RedirectUris = @($RedirectUri) }
        }
    }
}
#endregion

#region Create the App Registration
try {
    Write-Host "Creating App Registration: '$DisplayName'..." -ForegroundColor Cyan
    $newApp = New-MgApplication @appBody -ErrorAction Stop
    Write-Host "App Registration created." -ForegroundColor Green
    Write-Host "  Display Name    : $($newApp.DisplayName)"
    Write-Host "  App (Client) ID : $($newApp.AppId)"
    Write-Host "  Object ID       : $($newApp.Id)"
    if ($RedirectUri) {
        Write-Host "  Platform        : $PlatformType"
        Write-Host "  Redirect URI    : $RedirectUri"
    }
}
catch {
    Write-Error "Failed to create App Registration: $_"
    exit 1
}
#endregion

#region Create the Enterprise Application (Service Principal)
try {
    Write-Host "`nCreating Enterprise Application (Service Principal)..." -ForegroundColor Cyan
    $sp = New-MgServicePrincipal -AppId $newApp.AppId -ErrorAction Stop
    Write-Host "Enterprise Application created." -ForegroundColor Green
    Write-Host "  Service Principal ID : $($sp.Id)"
}
catch {
    Write-Warning "App Registration was created, but failed to create the Enterprise Application: $_"
}
#endregion

#region Add Owner
if ($OwnerUPN) {
    try {
        Write-Host "`nLooking up user '$OwnerUPN'..." -ForegroundColor Cyan
        $owner = Get-MgUser -UserId $OwnerUPN -ErrorAction Stop

        $ownerRef = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($owner.Id)"
        }

        New-MgApplicationOwnerByRef -ApplicationId $newApp.Id -BodyParameter $ownerRef -ErrorAction Stop
        Write-Host "Owner added: $($owner.DisplayName) ($($owner.UserPrincipalName))" -ForegroundColor Green
    }
    catch {
        Write-Warning "App was created, but failed to add owner '$OwnerUPN': $_"
    }
}
#endregion

Write-Host "`nDone." -ForegroundColor Green
