# Add delegated permissions to Graph Command Line Tools for a specific user
# Usage: .\Add-GraphPermissionsForUser.ps1 -UserPrincipalName "user@domain.com"
# Optionally specify permissions: -Permissions @("Chat.Read", "Mail.Read")

param (
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string[]]$Permissions = @(
        "User.Read.All"
        #"Group.ReadWrite.All"
        #"Mail.Read"
        #"Mail.ReadWrite"
        #"Calendars.Read"
        #"Calendars.ReadWrite"
        #"Files.Read.All"
        #"Files.ReadWrite.All"
        #"Directory.Read.All"
    )
)

Connect-MgGraph -Scopes "Application.ReadWrite.All", "DelegatedPermissionGrant.ReadWrite.All"

# Resolve the user
$user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
if (-not $user) {
    Write-Error "Could not find user: $UserPrincipalName"
    exit 1
}
Write-Host "Found user: $($user.DisplayName) ($($user.Id))" -ForegroundColor Cyan

# Graph Command Line Tools service principal
$servicePrincipalId = "b9384443-0c1a-4a9b-a5e5-a9662106ff7b"
$appServicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $servicePrincipalId -ErrorAction Stop

if (-not $appServicePrincipal) {
    Write-Error "Could not find Graph Command Line Tools service principal."
    exit 1
}

# Microsoft Graph service principal
$graphResourceId = "00000003-0000-0000-c000-000000000000"
$graphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '$graphResourceId'"
$graphPermissions = $graphServicePrincipal.Oauth2PermissionScopes

# Grant each permission for the specific user
foreach ($permission in $Permissions) {
    $scope = $graphPermissions | Where-Object { $_.Value -eq $permission }

    if (-not $scope) {
        Write-Host "Permission not found in Graph API: $permission" -ForegroundColor Red
        continue
    }

    try {
        $params = @{
            clientId    = $appServicePrincipal.Id
            consentType = "Principal"
            principalId = $user.Id
            resourceId  = $graphServicePrincipal.Id
            scope       = $permission
        }

        New-MgOauth2PermissionGrant @params -ErrorAction Stop
        Write-Host "Granted: $permission -> $($user.DisplayName)" -ForegroundColor Green
    }
    catch {
        Write-Host "Error granting $permission (may already exist): $_" -ForegroundColor Yellow
    }
}

# Show current user-specific grants for this app
Write-Host "`nCurrent delegated permissions for $($user.DisplayName):" -ForegroundColor Cyan
$userGrants = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($appServicePrincipal.Id)' and principalId eq '$($user.Id)'"
if ($userGrants) {
    $userGrants | Select-Object -ExpandProperty Scope | Sort-Object -Unique
} else {
    Write-Host "No user-specific grants found." -ForegroundColor Yellow
}
