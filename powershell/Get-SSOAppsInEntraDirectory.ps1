# Check if already connected to Microsoft Graph
$context = Get-MgContext
if (-not $context) {
    Write-Host "Not connected to Microsoft Graph. Connecting..."
    Connect-MgGraph -Scopes "Application.Read.All", "Directory.Read.All"
} else {
    Write-Host "Already connected to Microsoft Graph as $($context.Account)"
}

# Grab all service principals with SSO configured AND assignment required
$ssoApps = Get-MgServicePrincipal -All -Property "displayName,appId,preferredSingleSignOnMode,appRoleAssignmentRequired,id" |
    Where-Object { 
        $_.PreferredSingleSignOnMode -in @('saml', 'oidc', 'password', 'linked') -and
        $_.AppRoleAssignmentRequired -eq $true
    }

# Enrich with group assignments
$results = foreach ($app in $ssoApps) {
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