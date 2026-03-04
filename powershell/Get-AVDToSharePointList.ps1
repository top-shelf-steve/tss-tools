# ========================
# CONFIGURATION - Update these before deploying
# ========================
$SharePointSiteUrl = ""              # e.g. https://contoso.sharepoint.com/sites/IT
$ListName = "AVD Session Hosts"      # Name of the pre-created SharePoint list
$ResourceGroupNames = @(
    # "rg-avd-prod"
    # "rg-avd-dev"
)
$SubscriptionId = ""                 # Optional — uses current Az context if blank
$UseManagedIdentity = $false         # Set to $true for Azure Automation Runbook

# ========================
# AUTHENTICATION
# ========================
if ($UseManagedIdentity) {
    Connect-AzAccount -Identity | Out-Null
    Connect-MgGraph -Identity | Out-Null
    Write-Host "Connected via Managed Identity"
}
else {
    Connect-AzAccount | Out-Null
    Connect-MgGraph -Scopes "Sites.ReadWrite.All" | Out-Null
    Write-Host "Connected interactively"
}

if ($SubscriptionId) {
    Set-AzContext -SubscriptionId $SubscriptionId | Out-Null
}

# ========================
# COLLECT AVD DATA
# ========================
Write-Host "Collecting AVD data from resource groups: $($ResourceGroupNames -join ', ')"

# --- Host Pools ---
Write-Progress -Activity "AVD Export" -Status "Fetching host pools..." -PercentComplete 5
$hostPools = @()
foreach ($rg in $ResourceGroupNames) {
    $hps = Get-AzWvdHostPool -ResourceGroupName $rg -ErrorAction SilentlyContinue
    foreach ($hp in $hps) {
        $hostPools += [PSCustomObject]@{
            Name              = $hp.Name
            Description       = $hp.Description
            ResourceGroup     = $rg
            HostPoolType      = $hp.HostPoolType
            LoadBalancerType  = $hp.LoadBalancerType
            MaxSessionLimit   = $hp.MaxSessionLimit
            StartVMOnConnect  = $hp.StartVMOnConnect
            Id                = $hp.Id
        }
    }
}
Write-Host "  Found $($hostPools.Count) host pool(s)"

# --- Application Groups & Access Groups ---
Write-Progress -Activity "AVD Export" -Status "Fetching application groups..." -PercentComplete 15
$appGroupsByHostPool = @{}  # HostPoolName -> list of app group names
$appNamesByHostPool = @{}  # HostPoolName -> list of published application display names
$accessGroupsByHostPool = @{}  # HostPoolName -> list of Entra group names

foreach ($rg in $ResourceGroupNames) {
    $ags = Get-AzWvdApplicationGroup -ResourceGroupName $rg -ErrorAction SilentlyContinue
    foreach ($ag in $ags) {
        $hpName = ($ag.HostPoolArmPath -split '/')[-1]

        # Track app group names and friendly names per host pool
        if (-not $appGroupsByHostPool.ContainsKey($hpName)) { $appGroupsByHostPool[$hpName] = @() }
        $appGroupsByHostPool[$hpName] += $ag.Name
        # Get published application display names from this app group
        if (-not $appNamesByHostPool.ContainsKey($hpName)) { $appNamesByHostPool[$hpName] = @() }
        try {
            $apps = Get-AzWvdApplication -ResourceGroupName $rg -GroupName $ag.Name -ErrorAction SilentlyContinue
            foreach ($app in $apps) {
                $displayName = if ($app.FriendlyName) { $app.FriendlyName } else { $app.Name.Split('/')[-1] }
                if ($displayName -and $displayName -notin $appNamesByHostPool[$hpName]) {
                    $appNamesByHostPool[$hpName] += $displayName
                }
            }
        }
        catch {}

        # Get Desktop Virtualization User role assignments
        if (-not $accessGroupsByHostPool.ContainsKey($hpName)) { $accessGroupsByHostPool[$hpName] = @() }
        try {
            $assignments = Get-AzRoleAssignment -Scope $ag.Id -ErrorAction SilentlyContinue |
                Where-Object { $_.RoleDefinitionName -eq "Desktop Virtualization User" }
            foreach ($a in $assignments) {
                if ($a.DisplayName -and $a.DisplayName -notin $accessGroupsByHostPool[$hpName]) {
                    $accessGroupsByHostPool[$hpName] += $a.DisplayName
                }
            }
        }
        catch {}
    }
}

# --- Workspaces (build lookup: host pool -> workspace name) ---
Write-Progress -Activity "AVD Export" -Status "Fetching workspaces..." -PercentComplete 25
$workspaceByHostPool = @{}

foreach ($rg in $ResourceGroupNames) {
    $wsList = Get-AzWvdWorkspace -ResourceGroupName $rg -ErrorAction SilentlyContinue
    foreach ($ws in $wsList) {
        if ($ws.ApplicationGroupReference) {
            foreach ($agRef in $ws.ApplicationGroupReference) {
                # Find which host pool this app group belongs to
                $agName = ($agRef -split '/')[-1]
                $agRg = ($agRef -split '/')[4]
                try {
                    $agObj = Get-AzWvdApplicationGroup -ResourceGroupName $agRg -Name $agName -ErrorAction SilentlyContinue
                    if ($agObj) {
                        $hpName = ($agObj.HostPoolArmPath -split '/')[-1]
                        $workspaceByHostPool[$hpName] = $ws.Name
                    }
                }
                catch {}
            }
        }
    }
}

# --- Session Hosts ---
Write-Progress -Activity "AVD Export" -Status "Fetching session hosts..." -PercentComplete 35
$results = @()
$totalHPs = $hostPools.Count
$hpIndex = 0

foreach ($hp in $hostPools) {
    $hpIndex++
    Write-Host "  Collecting session hosts for: $($hp.Name)"

    $shList = Get-AzWvdSessionHost -ResourceGroupName $hp.ResourceGroup -HostPoolName $hp.Name -ErrorAction SilentlyContinue
    if (-not $shList) { continue }

    $shIndex = 0
    $totalSH = @($shList).Count
    foreach ($sh in $shList) {
        $shIndex++
        $pct = 35 + (60 * (($hpIndex - 1) / $totalHPs) + (60 / $totalHPs) * ($shIndex / $totalSH))
        $fullName = $sh.Name.Split('/')[-1]
        $hostName = $fullName.Split('.')[0]
        Write-Progress -Activity "AVD Export" -Status "Processing $hostName ($shIndex/$totalSH in $($hp.Name))" -PercentComplete ([math]::Min($pct, 95))

        # VM details
        $vmSize = "N/A"; $osDiskSize = "N/A"; $osType = "N/A"
        $privateIp = "N/A"; $subnetName = "N/A"; $vnetName = "N/A"

        try {
            $vm = Get-AzVM -ResourceGroupName $hp.ResourceGroup -Name $hostName -ErrorAction SilentlyContinue
            if (-not $vm) {
                $vmResourceId = $sh.ResourceId
                if ($vmResourceId) {
                    $vmRg = ($vmResourceId -split '/')[4]
                    $vm = Get-AzVM -ResourceGroupName $vmRg -Name $hostName -ErrorAction SilentlyContinue
                }
            }
            if ($vm) {
                $vmSize = $vm.HardwareProfile.VmSize
                $osDiskSize = "$($vm.StorageProfile.OsDisk.DiskSizeGB) GB"
                $osType = $vm.StorageProfile.OsDisk.OsType

                $nicId = $vm.NetworkProfile.NetworkInterfaces[0].Id
                if ($nicId) {
                    try {
                        $nic = Get-AzNetworkInterface -ResourceId $nicId -ErrorAction Stop
                        if ($nic) {
                            $ipConfig = $nic.IpConfigurations[0]
                            $privateIp = $ipConfig.PrivateIpAddress
                            if ($ipConfig.Subnet.Id) {
                                $subnetParts = $ipConfig.Subnet.Id -split '/'
                                $vnetName = $subnetParts[8]
                                $subnetName = $subnetParts[10]
                            }
                        }
                    }
                    catch {
                        Write-Host "    Warning: Could not read NIC for $hostName" -ForegroundColor Yellow
                    }
                }
            }
        }
        catch {
            Write-Host "    Warning: Could not get VM details for $hostName" -ForegroundColor Yellow
        }

        $assignedUser = if ($sh.AssignedUser) { $sh.AssignedUser } else { "" }

        $results += [PSCustomObject]@{
            HostName         = $hostName
            HostPool         = $hp.Name
            HostPoolDesc     = $hp.Description
            HostPoolType     = $hp.HostPoolType
            LoadBalancerType = $hp.LoadBalancerType
            MaxSessionLimit  = $hp.MaxSessionLimit
            StartVMOnConnect = if ($hp.StartVMOnConnect) { "Yes" } else { "No" }
            Status           = $sh.Status
            AssignedUser     = $assignedUser
            OSVersion        = $sh.OSVersion
            VMSize           = $vmSize
            OSDiskSizeGB     = $osDiskSize
            OSType           = $osType
            PrivateIP        = $privateIp
            Subnet           = $subnetName
            VNet             = $vnetName
            ResourceGroup    = $hp.ResourceGroup
            Workspace        = if ($workspaceByHostPool.ContainsKey($hp.Name)) { $workspaceByHostPool[$hp.Name] } else { "" }
            AppGroups        = if ($appGroupsByHostPool.ContainsKey($hp.Name)) { $appGroupsByHostPool[$hp.Name] -join '; ' } else { "" }
            FriendlyName = if ($appNamesByHostPool.ContainsKey($hp.Name)) { $appNamesByHostPool[$hp.Name] -join '; ' } else { "" }
            AccessGroups     = if ($accessGroupsByHostPool.ContainsKey($hp.Name)) { $accessGroupsByHostPool[$hp.Name] -join '; ' } else { "" }
        }
    }
}

Write-Progress -Activity "AVD Export" -Completed
Write-Host "Collected $($results.Count) session host(s) total"

# ========================
# SYNC TO SHAREPOINT LIST
# ========================

# Resolve the SharePoint site and list
$siteHostAndPath = $SharePointSiteUrl -replace 'https://', '' -split '/', 2
$siteHost = $siteHostAndPath[0]
$sitePath = if ($siteHostAndPath.Length -gt 1) { "/$($siteHostAndPath[1])" } else { "/" }
$site = Get-MgSite -SiteId "${siteHost}:${sitePath}"
$list = Get-MgSiteList -SiteId $site.Id -Filter "displayName eq '$ListName'"

Write-Host "Syncing to SharePoint list '$ListName' on site $($site.DisplayName)..."

# Fetch existing list items and build a lookup by HostName
$existingItems = Get-MgSiteListItem -SiteId $site.Id -ListId $list.Id -All -Expand "fields"
$existingLookup = @{}
foreach ($item in $existingItems) {
    $name = $item.Fields.AdditionalProperties["HostName"]
    if ($name) {
        $existingLookup[$name] = $item
    }
}

# Track which HostNames are still current (for stale cleanup)
$currentHostNames = @{}

# Upsert: update existing items or create new ones
$sortedResults = $results | Sort-Object HostPool, HostName
$syncCount = 0
$total = @($sortedResults).Count
foreach ($row in $sortedResults) {
    $syncCount++
    Write-Progress -Activity "SharePoint Sync" -Status "Syncing $syncCount of $total - $($row.HostName)" -PercentComplete ([math]::Round(($syncCount / $total) * 100))

    $currentHostNames[$row.HostName] = $true

    $fields = @{
        "HostName"         = $row.HostName
        "HostPool"         = $row.HostPool
        "HostPoolDesc"     = $row.HostPoolDesc
        "HostPoolType"     = $row.HostPoolType
        "LoadBalancerType" = $row.LoadBalancerType
        "MaxSessionLimit"  = $row.MaxSessionLimit
        "StartVMOnConnect" = $row.StartVMOnConnect
        "Status"           = $row.Status
        "AssignedUser"     = $row.AssignedUser
        "OSVersion"        = $row.OSVersion
        "VMSize"           = $row.VMSize
        "OSDiskSizeGB"     = $row.OSDiskSizeGB
        "OSType"           = $row.OSType
        "PrivateIP"        = $row.PrivateIP
        "Subnet"           = $row.Subnet
        "VNet"             = $row.VNet
        "ResourceGroup"    = $row.ResourceGroup
        "Workspace"        = $row.Workspace
        "AppGroups"        = $row.AppGroups
        "FriendlyName" = $row.FriendlyName
        "AccessGroups"     = $row.AccessGroups
    }

    if ($existingLookup.ContainsKey($row.HostName)) {
        # Update existing item
        $itemId = $existingLookup[$row.HostName].Id
        Update-MgSiteListItem -SiteId $site.Id -ListId $list.Id -ListItemId $itemId -BodyParameter @{ fields = $fields }
    }
    else {
        # Create new item
        New-MgSiteListItem -SiteId $site.Id -ListId $list.Id -BodyParameter @{ fields = $fields }
    }
}

# Remove stale items (session hosts that no longer exist)
$staleCount = 0
foreach ($item in $existingItems) {
    $name = $item.Fields.AdditionalProperties["HostName"]
    if ($name -and -not $currentHostNames.ContainsKey($name)) {
        Remove-MgSiteListItem -SiteId $site.Id -ListId $list.Id -ListItemId $item.Id
        $staleCount++
    }
}

Write-Progress -Activity "SharePoint Sync" -Completed
Write-Host "Sync complete: $syncCount items synced, $staleCount stale items removed."
