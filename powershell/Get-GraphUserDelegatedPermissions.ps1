# Get all user delegated permissions assigned to the Graph Command Line Tools enterprise app
# Displays both admin-consented (AllPrincipals) grants and per-user (Principal) grants

Connect-MgGraph -Scopes "Application.Read.All", "DelegatedPermissionGrant.ReadWrite.All"

# Graph Command Line Tools service principal ID (tenant-specific object ID)
$servicePrincipalId = ""

$appServicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $servicePrincipalId -ErrorAction Stop

if (-not $appServicePrincipal) {
    Write-Error "Could not find the Graph Command Line Tools service principal in this tenant."
    exit 1
}

Write-Host "Found: $($appServicePrincipal.DisplayName) (Object ID: $($appServicePrincipal.Id))" -ForegroundColor Cyan

# Retrieve all OAuth2 permission grants for this client app
$allGrants = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($appServicePrincipal.Id)'" -All

if (-not $allGrants) {
    Write-Host "`nNo delegated permission grants found for this application." -ForegroundColor Yellow
    exit 0
}

# Separate admin-consented grants from per-user grants
$adminGrants = $allGrants | Where-Object { $_.ConsentType -eq "AllPrincipals" }
$userGrants  = $allGrants | Where-Object { $_.ConsentType -eq "Principal" }

# --- Admin-consented permissions (apply to all users) ---
Write-Host "`n=== Admin-Consented Permissions (AllPrincipals) ===" -ForegroundColor Green

if ($adminGrants) {
    foreach ($grant in $adminGrants) {
        $scopes = $grant.Scope -split "\s+" | Where-Object { $_ } | Sort-Object
        Write-Host "  Resource: $($grant.ResourceId)"
        $scopes | ForEach-Object { Write-Host "    - $_" }
    }
} else {
    Write-Host "  None found." -ForegroundColor Yellow
}

# --- Per-user permissions ---
Write-Host "`n=== Per-User Delegated Permissions (Principal) ===" -ForegroundColor Green

if ($userGrants) {
    # Group grants by user for a clean display
    $userGrants | Group-Object -Property PrincipalId | ForEach-Object {
        $userId = $_.Name
        try {
            $user = Get-MgUser -UserId $userId -Property DisplayName, UserPrincipalName -ErrorAction Stop
            $displayName = "$($user.DisplayName) ($($user.UserPrincipalName))"
        } catch {
            $displayName = "Unknown user (ID: $userId)"
        }

        Write-Host "`n  User: $displayName" -ForegroundColor Cyan
        foreach ($grant in $_.Group) {
            $scopes = $grant.Scope -split "\s+" | Where-Object { $_ } | Sort-Object
            $scopes | ForEach-Object { Write-Host "    - $_" }
        }
    }
} else {
    Write-Host "  None found." -ForegroundColor Yellow
}

# --- All unique delegated permissions across all grants ---
$allScopes = $allGrants | ForEach-Object { $_.Scope -split "\s+" | Where-Object { $_ } } | Sort-Object -Unique

Write-Host "`n=== All Delegated Permissions (Combined) ===" -ForegroundColor Green
$allScopes | ForEach-Object { Write-Host "  - $_" }

$scopesFormatted = ($allScopes | ForEach-Object { "$_," }) -join "`n"
Set-Clipboard -Value $scopesFormatted
Write-Host "`nCopied to clipboard:" -ForegroundColor Cyan
Write-Host $scopesFormatted

# --- Summary ---
Write-Host "`n=== Summary ===" -ForegroundColor Green
Write-Host "  Admin-consented unique scopes : $(($adminGrants | ForEach-Object { $_.Scope -split "\s+" | Where-Object { $_ } } | Select-Object -Unique).Count)"
Write-Host "  Per-user grants               : $(($userGrants | Measure-Object).Count)"
Write-Host "  Total unique scopes           : $($allScopes.Count)"
