# Batch create Entra App Registrations interactively
# Loops until you're done, then copies a summary of all created apps to clipboard
# Usage: . .\People.ps1; .\Create-BatchEntraAppRegistrations.ps1

#region Prerequisites
try {
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
}
catch {
    Write-Error "Microsoft.Graph modules not found. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    exit 1
}

try {
    Connect-MgGraph -Scopes "Application.ReadWrite.All", "User.Read.All" -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}
#endregion

$createdApps = @()

do {
    Write-Host "`n--- New App Registration ---" -ForegroundColor Cyan

    $displayName = Read-Host "Display Name"
    if (-not $displayName) {
        Write-Host "Display name is required, skipping." -ForegroundColor Yellow
        continue
    }

    $ownerUPN = Read-Host "Owner UPN (leave blank to skip)"

    $redirectUri = Read-Host "Redirect URI (leave blank to skip)"
    $platformType = "Web"
    if ($redirectUri) {
        $platformInput = Read-Host "Platform type [Web/SPA/PublicClient] (default: Web)"
        if ($platformInput -in "SPA", "PublicClient") { $platformType = $platformInput }
    }

    # Check for existing app with same name
    $existingApp = Get-MgApplication -Filter "displayName eq '$displayName'" -Top 1
    if ($existingApp) {
        Write-Host "An App Registration named '$displayName' already exists. Skipping." -ForegroundColor Yellow
        continue
    }

    # Build app body
    $appBody = @{
        DisplayName    = $displayName
        SignInAudience = "AzureADMyOrg"
    }

    if ($redirectUri) {
        switch ($platformType) {
            "Web"          { $appBody.Web = @{ RedirectUris = @($redirectUri) } }
            "SPA"          { $appBody.Spa = @{ RedirectUris = @($redirectUri) } }
            "PublicClient" { $appBody.PublicClient = @{ RedirectUris = @($redirectUri) } }
        }
    }

    # Create app registration
    try {
        Write-Host "Creating App Registration: '$displayName'..." -ForegroundColor Cyan
        $newApp = New-MgApplication @appBody -ErrorAction Stop
        Write-Host "App Registration created." -ForegroundColor Green
        Write-Host "  Display Name    : $($newApp.DisplayName)"
        Write-Host "  App (Client) ID : $($newApp.AppId)"
        Write-Host "  Object ID       : $($newApp.Id)"
    }
    catch {
        Write-Error "Failed to create App Registration: $_"
        continue
    }

    # Create enterprise application
    try {
        $sp = New-MgServicePrincipal -AppId $newApp.AppId -ErrorAction Stop
        Write-Host "Enterprise Application created. SP ID: $($sp.Id)" -ForegroundColor Green
    }
    catch {
        Write-Warning "App Registration created, but failed to create Enterprise Application: $_"
    }

    # Add owner(s)
    if ($ownerUPN) {
        $owners = $ownerUPN -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        foreach ($upn in $owners) {
            try {
                $owner = Get-MgUser -UserId $upn -ErrorAction Stop
                $ownerRef = @{
                    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($owner.Id)"
                }
                New-MgApplicationOwnerByRef -ApplicationId $newApp.Id -BodyParameter $ownerRef -ErrorAction Stop
                Write-Host "Owner added: $($owner.DisplayName) ($($owner.UserPrincipalName))" -ForegroundColor Green
            }
            catch {
                Write-Warning "Failed to add owner '$upn': $_"
            }
        }
    }

    # Track the created app
    $createdApps += [PSCustomObject]@{
        DisplayName = $newApp.DisplayName
        ClientId    = $newApp.AppId
    }

    Write-Host "`nApp #$($createdApps.Count) done." -ForegroundColor Green
    $another = Read-Host "Create another? (y/n)"

} while ($another -eq "y")

# --- Summary ---
if ($createdApps.Count -gt 0) {
    Write-Host "`n=== Created App Registrations ===" -ForegroundColor Green
    $createdApps | Format-Table -AutoSize

    $clipboardText = ($createdApps | ForEach-Object { "$($_.DisplayName) — $($_.ClientId)" }) -join "`n"
    Set-Clipboard -Value $clipboardText
    Write-Host "Copied to clipboard:" -ForegroundColor Cyan
    Write-Host $clipboardText
}
