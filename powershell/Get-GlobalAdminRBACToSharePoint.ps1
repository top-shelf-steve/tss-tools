#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Captures PIM activations into Global Administrator (direct role and via the
    PIM-GlobalAdmin group) over the last N days, exports a CSV, and uploads it
    twice to a Teams-channel-backed SharePoint folder: once as a fixed-name
    "Latest" file (used as the Excel workbook's data source) and once as a
    dated copy in an Archive/ subfolder for permanent audit history.

.DESCRIPTION
    Designed to run as an Azure Automation runbook on a 30-day schedule under a
    system-assigned managed identity. Set $UseManagedIdentity = $false to run
    locally with interactive Graph auth for testing.

    Two activation sources are queried so we cover both paths users take to GA:
      * PIM-for-Group : selfActivate requests against the PIM-GlobalAdmin group
                        (users granted GA transitively via group membership).
      * PIM-for-Role  : selfActivate requests against the Global Administrator
                        role definition (users granted GA directly via PIM).

    Each request includes Justification and TicketInfo (TicketNumber +
    TicketSystem) as first-class fields, so we don't have to scrape the audit
    log.

    Output layout in the Teams channel folder:
      <ChannelFolder>/PIM-GlobalAdmin-Activations-Latest.csv      <- Excel data source
      <ChannelFolder>/Archive/PIM-GlobalAdmin-Activations-YYYY-MM-DD.csv

    The Latest file is overwritten every run so the Excel workbook's link
    never breaks. The Archive copy is dated and immutable.

.NOTES
    Required Graph permissions on the managed identity (application):
      * PrivilegedAccess.Read.AzureADGroup
      * RoleManagement.Read.Directory
      * User.Read.All
      * Sites.Selected            (granted on the target site only)

    Schedule: every 30 days via the Automation Account schedule.
    Automation Account: <PLACEHOLDER-AutomationAccountName>
#>

# ======================== CONFIGURATION ========================

# Object ID of the "PIM-GlobalAdmin" group in Entra
$PimGlobalAdminGroupId = "<PLACEHOLDER-PIM-GlobalAdmin-GroupObjectId>"

# Built-in Entra role template ID for Global Administrator (do not change)
$GlobalAdministratorRoleId = "62e90394-69f5-4237-9190-012177145e10"

# Lookback window in days. Match this to the Automation schedule interval.
$DaysBack = 30

# SharePoint destination - point at the Team's site and the channel's folder
$SharePointSiteUrl = "https://<PLACEHOLDER-tenant>.sharepoint.com/sites/<PLACEHOLDER-teamsite>"
$LibraryPath       = "Shared Documents/<PLACEHOLDER-ChannelName>"
$ArchiveSubfolder  = "Archive"
$FilePrefix        = "PIM-GlobalAdmin-Activations"
$LatestFileName    = "$FilePrefix-Latest.csv"

# Automation Account this runbook is deployed to (reference only)
$AutomationAccountName = "<PLACEHOLDER-AutomationAccountName>"

# Auth mode - set to $true when deploying to the Automation runbook
$UseManagedIdentity = $false

# ======================== AUTHENTICATION ========================
$requiredScopes = @(
    "PrivilegedAccess.Read.AzureADGroup",
    "RoleManagement.Read.Directory",
    "User.Read.All",
    "Sites.ReadWrite.All"
)

if ($UseManagedIdentity) {
    Connect-MgGraph -Identity -NoWelcome
    Write-Host "Connected via managed identity" -ForegroundColor Green
}
else {
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    Write-Host "Connected interactively as $((Get-MgContext).Account)" -ForegroundColor Green
}

# ======================== DATE WINDOW ========================
$cutoff    = (Get-Date).ToUniversalTime().AddDays(-$DaysBack)
$cutoffIso = $cutoff.ToString("yyyy-MM-ddTHH:mm:ssZ")
Write-Host "Capturing selfActivate requests from $cutoffIso onward (last $DaysBack days)"

# ======================== HELPERS ========================
function Invoke-GraphPaged {
    param([string]$Uri)
    $items = @()
    while ($Uri) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri
        if ($resp.value) { $items += $resp.value }
        $Uri = $resp.'@odata.nextLink'
    }
    return $items
}

function Resolve-Principal {
    param([string]$PrincipalId, [hashtable]$Cache)
    if ($Cache.ContainsKey($PrincipalId)) { return $Cache[$PrincipalId] }
    try {
        $u = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$PrincipalId`?`$select=id,displayName,userPrincipalName"
        $Cache[$PrincipalId] = [pscustomobject]@{
            DisplayName       = $u.displayName
            UserPrincipalName = $u.userPrincipalName
        }
    }
    catch {
        $Cache[$PrincipalId] = [pscustomobject]@{
            DisplayName       = "(unresolved)"
            UserPrincipalName = $PrincipalId
        }
    }
    return $Cache[$PrincipalId]
}

function Send-FileToSharePoint {
    param(
        [string]$SiteId,
        [string]$DriveRelativePath,   # e.g. "ChannelName/Archive/file.csv"
        [string]$LocalPath,
        [string]$ContentType = "text/csv"
    )
    $encodedPath = ($DriveRelativePath -split '/' | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join '/'
    $uri   = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/${encodedPath}:/content"
    $bytes = [System.IO.File]::ReadAllBytes($LocalPath)
    Invoke-MgGraphRequest -Method PUT -Uri $uri -Body $bytes -ContentType $ContentType | Out-Null
}

# ======================== FETCH PIM-FOR-GROUP ACTIVATIONS ========================
Write-Host "Fetching PIM-for-Group activations for group $PimGlobalAdminGroupId..."
$groupFilter = "groupId eq '$PimGlobalAdminGroupId' and action eq 'selfActivate' and createdDateTime ge $cutoffIso"
$groupUri    = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/assignmentScheduleRequests?`$filter=" + [System.Uri]::EscapeDataString($groupFilter)
$groupRequests = Invoke-GraphPaged -Uri $groupUri
Write-Host "  Found $($groupRequests.Count) group activation request(s)"

# ======================== FETCH PIM-FOR-ROLE ACTIVATIONS ========================
Write-Host "Fetching PIM-for-Role activations for Global Administrator..."
$roleFilter = "action eq 'selfActivate' and roleDefinitionId eq '$GlobalAdministratorRoleId' and createdDateTime ge $cutoffIso"
$roleUri    = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests?`$filter=" + [System.Uri]::EscapeDataString($roleFilter)
$roleRequests = Invoke-GraphPaged -Uri $roleUri
Write-Host "  Found $($roleRequests.Count) role activation request(s)"

# ======================== NORMALIZE INTO ROWS ========================
$principalCache = @{}
$rows = @()

foreach ($r in $groupRequests) {
    $p = Resolve-Principal -PrincipalId $r.principalId -Cache $principalCache
    $rows += [pscustomobject]@{
        ActivationDateTimeUtc = $r.createdDateTime
        ActivationSource      = "PIM-for-Group"
        Target                = "PIM-GlobalAdmin (group)"
        UserDisplayName       = $p.DisplayName
        UserPrincipalName     = $p.UserPrincipalName
        PrincipalId           = $r.principalId
        TicketNumber          = $r.ticketInfo.ticketNumber
        TicketSystem          = $r.ticketInfo.ticketSystem
        Justification         = $r.justification
        Status                = $r.status
        StartDateTime         = $r.scheduleInfo.startDateTime
        Expiration            = $r.scheduleInfo.expiration.endDateTime
        RequestId             = $r.id
    }
}

foreach ($r in $roleRequests) {
    $p = Resolve-Principal -PrincipalId $r.principalId -Cache $principalCache
    $rows += [pscustomobject]@{
        ActivationDateTimeUtc = $r.createdDateTime
        ActivationSource      = "PIM-for-Role"
        Target                = "Global Administrator (role)"
        UserDisplayName       = $p.DisplayName
        UserPrincipalName     = $p.UserPrincipalName
        PrincipalId           = $r.principalId
        TicketNumber          = $r.ticketInfo.ticketNumber
        TicketSystem          = $r.ticketInfo.ticketSystem
        Justification         = $r.justification
        Status                = $r.status
        StartDateTime         = $r.scheduleInfo.startDateTime
        Expiration            = $r.scheduleInfo.expiration.endDateTime
        RequestId             = $r.id
    }
}

$rows = $rows | Sort-Object ActivationDateTimeUtc
Write-Host "Total activations captured: $($rows.Count)"

# ======================== WRITE CSV ========================
$snapshotDate    = (Get-Date).ToString("yyyy-MM-dd")
$archiveFileName = "$FilePrefix-$snapshotDate.csv"
$tempPath        = Join-Path -Path $env:TEMP -ChildPath $archiveFileName
$rows | Export-Csv -Path $tempPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV written: $tempPath"

# ======================== UPLOAD TO SHAREPOINT ========================
$siteHostAndPath = $SharePointSiteUrl -replace 'https://','' -split '/',2
$siteHost = $siteHostAndPath[0]
$sitePath = if ($siteHostAndPath.Length -gt 1) { "/$($siteHostAndPath[1])" } else { "/" }
$site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/${siteHost}:${sitePath}"

# Default drive root == "Documents" library, so strip the "Shared Documents/" prefix
$channelRelativePath = $LibraryPath
if ($channelRelativePath -match '^Shared Documents/?(.*)') { $channelRelativePath = $matches[1] }
$channelRelativePath = $channelRelativePath.Trim('/')

$latestDrivePath  = "$channelRelativePath/$LatestFileName".TrimStart('/')
$archiveDrivePath = "$channelRelativePath/$ArchiveSubfolder/$archiveFileName".TrimStart('/')

Write-Host "Uploading Latest snapshot..."
Send-FileToSharePoint -SiteId $site.id -DriveRelativePath $latestDrivePath -LocalPath $tempPath
Write-Host "  -> $SharePointSiteUrl/$LibraryPath/$LatestFileName" -ForegroundColor Green

Write-Host "Uploading Archive copy..."
Send-FileToSharePoint -SiteId $site.id -DriveRelativePath $archiveDrivePath -LocalPath $tempPath
Write-Host "  -> $SharePointSiteUrl/$LibraryPath/$ArchiveSubfolder/$archiveFileName" -ForegroundColor Green

# ======================== CLEANUP ========================
Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
