<#
.SYNOPSIS
    Adds a user as an owner to one or more existing Entra ID (Azure AD) App Registrations.

.DESCRIPTION
    Uses the Microsoft.Graph PowerShell module to look up one or more App Registrations
    by display name, verify a user exists, and add that user as an owner on each app.

    - Verifies each App Registration exists (by exact display name). Skips any that
      don't exist or are ambiguous (more than one match).
    - Verifies the user exists (by UPN or Object ID).
    - Skips adding the owner if they are already an owner of the app.

.PARAMETER DisplayName
    One or more App Registration display names. Accepts pipeline input.

.PARAMETER OwnerUPN
    The UPN or Object ID of the user to add as an owner.

.EXAMPLE
    .\Add-OwnerToAppRegistration.ps1 -DisplayName "MyApp" -OwnerUPN "alice@contoso.com"

.EXAMPLE
    .\Add-OwnerToAppRegistration.ps1 -DisplayName "App1","App2","App3" -OwnerUPN "alice@contoso.com"

.EXAMPLE
    "App1","App2" | .\Add-OwnerToAppRegistration.ps1 -OwnerUPN "alice@contoso.com"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
    [string[]]$DisplayName,

    [Parameter(Mandatory = $true, Position = 1)]
    [string]$OwnerUPN
)

begin {
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

    #region Look up the user once
    try {
        Write-Host "Looking up user '$OwnerUPN'..." -ForegroundColor Cyan
        $owner = Get-MgUser -UserId $OwnerUPN -ErrorAction Stop
        Write-Host "Found user: $($owner.DisplayName) ($($owner.UserPrincipalName))" -ForegroundColor Green
    }
    catch {
        Write-Error "User '$OwnerUPN' not found in Entra: $_"
        exit 1
    }

    $ownerRef = @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($owner.Id)"
    }
    #endregion

    $results = @()
}

process {
    foreach ($name in $DisplayName) {
        Write-Host "`n--- $name ---" -ForegroundColor Cyan

        # Find the app registration(s) matching this display name
        try {
            $apps = @(Get-MgApplication -Filter "displayName eq '$name'" -All -ErrorAction Stop)
        }
        catch {
            Write-Warning "Failed to query App Registration '$name': $_"
            $results += [PSCustomObject]@{ DisplayName = $name; Status = "QueryFailed" }
            continue
        }

        if ($apps.Count -eq 0) {
            Write-Warning "No App Registration found with display name '$name'. Skipping."
            $results += [PSCustomObject]@{ DisplayName = $name; Status = "NotFound" }
            continue
        }

        if ($apps.Count -gt 1) {
            Write-Warning "Multiple App Registrations found with display name '$name' ($($apps.Count) matches). Skipping to avoid ambiguity."
            $results += [PSCustomObject]@{ DisplayName = $name; Status = "Ambiguous" }
            continue
        }

        $app = $apps[0]
        Write-Host "Found: $($app.DisplayName)  AppId=$($app.AppId)  ObjectId=$($app.Id)"

        # Check if user is already an owner
        try {
            $existingOwners = Get-MgApplicationOwner -ApplicationId $app.Id -All -ErrorAction Stop
            if ($existingOwners.Id -contains $owner.Id) {
                Write-Host "$($owner.UserPrincipalName) is already an owner. Skipping." -ForegroundColor Yellow
                $results += [PSCustomObject]@{ DisplayName = $name; Status = "AlreadyOwner" }
                continue
            }
        }
        catch {
            Write-Warning "Failed to read existing owners for '$name', attempting add anyway: $_"
        }

        # Add owner
        try {
            New-MgApplicationOwnerByRef -ApplicationId $app.Id -BodyParameter $ownerRef -ErrorAction Stop
            Write-Host "Owner added: $($owner.DisplayName) ($($owner.UserPrincipalName))" -ForegroundColor Green
            $results += [PSCustomObject]@{ DisplayName = $name; Status = "Added" }
        }
        catch {
            Write-Warning "Failed to add owner to '$name': $_"
            $results += [PSCustomObject]@{ DisplayName = $name; Status = "AddFailed" }
        }
    }
}

end {
    if ($results.Count -gt 0) {
        Write-Host "`n=== Summary ===" -ForegroundColor Green
        $results | Format-Table -AutoSize
    }
    Write-Host "Done." -ForegroundColor Green
}
