<#
.SYNOPSIS
    Lists all Entra ID (Azure AD) App Registrations a given user is assigned as an owner of.

.DESCRIPTION
    Uses the Microsoft.Graph PowerShell module to look up a user by email/UPN (or Object ID),
    then enumerates the directory objects they own and filters the results down to
    App Registrations (microsoft.graph.application).

    Outputs one PSCustomObject per owned App Registration with DisplayName, AppId,
    ObjectId, SignInAudience and CreatedDateTime. Optionally writes a CSV.

.PARAMETER UserUPN
    The user's email / UPN, or their directory Object ID.

.PARAMETER OutputCsv
    Optional path to a CSV file. If provided, results are also written there.

.EXAMPLE
    .\Get-AppRegOwnedByUser.ps1 -UserUPN "alice@contoso.com"

.EXAMPLE
    .\Get-AppRegOwnedByUser.ps1 -UserUPN "alice@contoso.com" -OutputCsv ".\alice-apps.csv"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$UserUPN,

    [Parameter(Mandatory = $false)]
    [string]$OutputCsv
)

#region Prerequisites
$requiredScopes = @("Application.Read.All", "User.Read.All", "Directory.Read.All")

try {
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
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

#region Look up the user
try {
    Write-Host "Looking up user '$UserUPN'..." -ForegroundColor Cyan
    $user = Get-MgUser -UserId $UserUPN -ErrorAction Stop
    Write-Host "Found user: $($user.DisplayName) ($($user.UserPrincipalName))  Id=$($user.Id)" -ForegroundColor Green
}
catch {
    Write-Error "User '$UserUPN' not found in Entra: $_"
    exit 1
}
#endregion

#region Enumerate owned objects and filter to applications
try {
    Write-Host "Retrieving objects owned by $($user.UserPrincipalName)..." -ForegroundColor Cyan
    $ownedObjects = @(Get-MgUserOwnedObject -UserId $user.Id -All -ErrorAction Stop)
}
catch {
    Write-Error "Failed to retrieve owned objects: $_"
    exit 1
}

$appType = "#microsoft.graph.application"
$ownedApps = @($ownedObjects | Where-Object {
    $_.AdditionalProperties["@odata.type"] -eq $appType
})

Write-Host "Found $($ownedApps.Count) App Registration(s) owned by $($user.UserPrincipalName)." -ForegroundColor Green
#endregion

#region Project results
$results = foreach ($obj in $ownedApps) {
    $props = $obj.AdditionalProperties
    [PSCustomObject]@{
        DisplayName      = $props["displayName"]
        AppId            = $props["appId"]
        ObjectId         = $obj.Id
        SignInAudience   = $props["signInAudience"]
        CreatedDateTime  = $props["createdDateTime"]
    }
}

$results = @($results | Sort-Object DisplayName)
#endregion

if ($results.Count -gt 0) {
    $results | Format-Table -AutoSize

    if ($OutputCsv) {
        try {
            $results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
            Write-Host "Wrote $($results.Count) row(s) to $OutputCsv" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to write CSV to '$OutputCsv': $_"
        }
    }
}
else {
    Write-Host "No App Registrations found for this user." -ForegroundColor Yellow
}

# Emit results to the pipeline so callers can consume them
$results
