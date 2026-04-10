# Add delegated permissions to Graph Command Line Tools Enterprise application
# This script adds the required delegated permissions to enable the app to function properly

Connect-MgGraph -Scopes "Application.ReadWrite.All", "DelegatedPermissionGrant.ReadWrite.All"

# Graph Command Line Tools service principal ID (tenant-specific object ID)
$servicePrincipalId = ""

# Look up the service principal by object ID
$appServicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $servicePrincipalId

if (-not $appServicePrincipal) {
    Write-Error "Could not find Graph Command Line Tools service principal in this tenant."
    exit 1
}

# Get the service principal for Microsoft Graph API
$graphResourceId = "00000003-0000-0000-c000-000000000000"
$graphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '$graphResourceId'"

# Define the delegated permissions to add
$delegatedPermissions = @(
    "User.Read.All"
    #"Group.ReadWrite.All"
    #"Mail.Read"
    #"Mail.ReadWrite"
    #"Calendars.Read"
    #"Calendars.ReadWrite"
    #"Files.Read.All"
    #"Files.ReadWrite.All"
    #"Directory.Read.All"
    #"Directory.ReadWrite.All"
)

# Check for an existing AllPrincipals grant for this app
$existingGrant = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($appServicePrincipal.Id)' and consentType eq 'AllPrincipals' and resourceId eq '$($graphServicePrincipal.Id)'" |
    Select-Object -First 1

$currentScopes = @()
if ($existingGrant) {
    $currentScopes = $existingGrant.Scope -split "\s+" | Where-Object { $_ }
    Write-Host "Found existing grant. Current scopes: $($currentScopes -join ', ')" -ForegroundColor Cyan
}

# Merge new permissions with existing ones, removing any that are substrings of a new permission
# e.g. remove "User.Read" if "User.Read.All" is being added
$filteredCurrent = $currentScopes | Where-Object {
    $existing = $_
    -not ($delegatedPermissions | Where-Object { $_ -like "$existing.*" })
}
$mergedScopes = ($filteredCurrent + $delegatedPermissions | Select-Object -Unique) -join " "

if ($existingGrant) {
    # Update the existing grant
    Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $existingGrant.Id -Scope $mergedScopes
    Write-Host "Updated existing grant with: $($delegatedPermissions -join ', ')" -ForegroundColor Green
} else {
    # Create a new grant
    $params = @{
        clientId    = $appServicePrincipal.Id
        consentType = "AllPrincipals"
        principalId = $null
        resourceId  = $graphServicePrincipal.Id
        scope       = $mergedScopes
    }
    New-MgOauth2PermissionGrant @params
    Write-Host "Created new grant with: $($delegatedPermissions -join ', ')" -ForegroundColor Green
}

# Display final permissions
Write-Host "`nCurrent admin-consented permissions:" -ForegroundColor Cyan
$finalGrant = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($appServicePrincipal.Id)' and consentType eq 'AllPrincipals' and resourceId eq '$($graphServicePrincipal.Id)'"
$finalGrant.Scope -split "\s+" | Where-Object { $_ } | Sort-Object
