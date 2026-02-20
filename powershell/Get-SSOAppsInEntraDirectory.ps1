# Check if already connected to Microsoft Graph
$context = Get-MgContext
if (-not $context) {
    Write-Host "Not connected to Microsoft Graph. Connecting..."
    Connect-MgGraph -Scopes "Application.Read.All", "Directory.Read.All"
} else {
    Write-Host "Already connected to Microsoft Graph as $($context.Account)"
}

# Grab all service principals with SSO configured AND assignment required
Write-Progress -Activity "SSO Apps Report" -Status "Fetching service principals..." -PercentComplete 0
$ssoApps = Get-MgServicePrincipal -All -Property "displayName,appId,preferredSingleSignOnMode,appRoleAssignmentRequired,id" |
    Where-Object {
        $_.PreferredSingleSignOnMode -in @('saml', 'oidc', 'password', 'linked') -and
        $_.AppRoleAssignmentRequired -eq $true
    }

Write-Host "Found $($ssoApps.Count) SSO apps with assignment required. Fetching group assignments..."

# Enrich with group assignments
$i = 0
$total = $ssoApps.Count
$results = foreach ($app in $ssoApps) {
    $i++
    $pct = [math]::Round(($i / $total) * 100)
    Write-Progress -Activity "SSO Apps Report" -Status "Processing $i of $total - $($app.DisplayName)" -PercentComplete $pct

    $assignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $app.Id -All |
        Where-Object { $_.PrincipalType -eq 'Group' }

    [PSCustomObject]@{
        AppName        = $app.DisplayName
        AppId          = $app.AppId
        SSOMode        = $app.PreferredSingleSignOnMode
        AssignedGroups = ($assignments | ForEach-Object { $_.PrincipalDisplayName }) -join '; '
        SPObjectId     = $app.Id
    }
}
Write-Progress -Activity "SSO Apps Report" -Completed

# Prompt for save location
Add-Type -AssemblyName System.Windows.Forms
$saveDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveDialog.Filter = "CSV files (*.csv)|*.csv"
$saveDialog.FileName = "SSO_Apps_Report.csv"
$saveDialog.Title = "Save SSO Apps Report"

if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $results | Sort-Object AppName | Export-Csv -Path $saveDialog.FileName -NoTypeInformation
    Write-Host "Report saved to $($saveDialog.FileName)"
} else {
    Write-Host "Save cancelled."
}