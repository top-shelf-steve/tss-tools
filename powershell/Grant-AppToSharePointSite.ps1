#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Grants an app registration (service principal) Sites.Selected access to a
    specific SharePoint site, at a chosen permission level.

.DESCRIPTION
    Sites.Selected is a two-step model:
      1. The app registration is granted the "Sites.Selected" application
         permission in Entra (admin consent). This grants no access on its own.
      2. A site-scoped permission entry is created on each target site, naming
         the app and the role it should have. This script performs step 2.

    The caller running this script needs delegated "Sites.FullControl.All"
    in order to write the permission entry. Once granted, the app itself only
    needs Sites.Selected.

    Use -List to view existing app permissions on the site, or -Remove with
    -PermissionId to revoke an existing grant.

.PARAMETER SiteUrl
    Full URL of the SharePoint site, e.g.
    "https://contoso.sharepoint.com/sites/Finance".

.PARAMETER AppId
    Client (Application) ID of the app registration to grant.

.PARAMETER AppDisplayName
    Display name of the app registration. Used only for labelling the
    permission entry; does not need to match exactly but should be recognisable.

.PARAMETER Role
    Permission level to grant. One of: read, write, manage, fullcontrol.
    Defaults to "write".

.PARAMETER List
    List existing app permissions on the site instead of granting.

.PARAMETER Remove
    Remove an existing permission entry. Requires -PermissionId.

.PARAMETER PermissionId
    The ID of the permission entry to remove. Get this from -List output.

.EXAMPLE
    .\Grant-AppToSharePointSite.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/Finance" `
        -AppId "11111111-2222-3333-4444-555555555555" `
        -AppDisplayName "Finance Reporting Runbook" `
        -Role write

.EXAMPLE
    .\Grant-AppToSharePointSite.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -List

.EXAMPLE
    .\Grant-AppToSharePointSite.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/Finance" `
        -Remove -PermissionId "aTowaS50fG1zLnNwLmV4dHw..."

.NOTES
    Required Graph permissions for the user running the script (delegated):
      * Sites.FullControl.All     - to write permission entries on the site

    The target app only needs the Sites.Selected application permission
    (admin-consented in Entra). No tenant-wide site access is conferred.
#>

[CmdletBinding(DefaultParameterSetName = "Grant")]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true, ParameterSetName = "Grant")]
    [string]$AppId,

    [Parameter(Mandatory = $true, ParameterSetName = "Grant")]
    [string]$AppDisplayName,

    [Parameter(ParameterSetName = "Grant")]
    [ValidateSet("read", "write", "manage", "fullcontrol")]
    [string]$Role = "write",

    [Parameter(Mandatory = $true, ParameterSetName = "List")]
    [switch]$List,

    [Parameter(Mandatory = $true, ParameterSetName = "Remove")]
    [switch]$Remove,

    [Parameter(Mandatory = $true, ParameterSetName = "Remove")]
    [string]$PermissionId
)

# ======================== AUTHENTICATION ========================
Connect-MgGraph -Scopes "Sites.FullControl.All" -NoWelcome
Write-Host "Connected as $((Get-MgContext).Account)" -ForegroundColor Green

# ======================== RESOLVE SITE ID ========================
$uri      = [uri]$SiteUrl
$hostname = $uri.Host
$path     = $uri.AbsolutePath

Write-Host "Resolving site '$SiteUrl'..."
$site = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/sites/${hostname}:${path}"

if (-not $site.id) {
    throw "Could not resolve site '$SiteUrl'."
}
Write-Host "  Site ID: $($site.id)" -ForegroundColor DarkGray

# ======================== EXECUTE ========================
switch ($PSCmdlet.ParameterSetName) {

    "List" {
        Write-Host "Listing app permissions on site..."
        $perms = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/permissions"

        if (-not $perms.value -or $perms.value.Count -eq 0) {
            Write-Host "  (no app permissions granted on this site)" -ForegroundColor Yellow
            break
        }

        $perms.value | ForEach-Object {
            [pscustomobject]@{
                PermissionId = $_.id
                Roles        = ($_.roles -join ", ")
                AppId        = $_.grantedToIdentitiesV2.application.id
                AppName      = $_.grantedToIdentitiesV2.application.displayName
            }
        } | Format-Table -AutoSize
    }

    "Remove" {
        Write-Host "Removing permission '$PermissionId' from site..."
        Invoke-MgGraphRequest -Method DELETE `
            -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/permissions/$PermissionId" | Out-Null
        Write-Host "Permission removed." -ForegroundColor Green
    }

    "Grant" {
        Write-Host "Granting '$Role' on site to app '$AppDisplayName' ($AppId)..."

        $body = @{
            roles = @($Role)
            grantedToIdentities = @(
                @{
                    application = @{
                        id          = $AppId
                        displayName = $AppDisplayName
                    }
                }
            )
        } | ConvertTo-Json -Depth 5

        $result = Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/permissions" `
            -Body $body -ContentType "application/json"

        Write-Host "Grant complete." -ForegroundColor Green
        [pscustomobject]@{
            PermissionId = $result.id
            Roles        = ($result.roles -join ", ")
            AppId        = $result.grantedToIdentitiesV2.application.id
            AppName      = $result.grantedToIdentitiesV2.application.displayName
            SiteUrl      = $SiteUrl
        } | Format-List
    }
}
