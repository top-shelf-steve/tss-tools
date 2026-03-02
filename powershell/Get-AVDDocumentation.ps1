<#
.SYNOPSIS
    Generates comprehensive Markdown documentation for Azure Virtual Desktop environments.

.DESCRIPTION
    Queries Azure for all AVD-related resources across specified resource groups and produces
    a Markdown file documenting workspaces, host pools, application groups, session hosts
    (with VM specs and networking), VNet/subnet topology, and optionally storage accounts
    used for FSLogix profiles.

.PARAMETER ResourceGroupNames
    An array of resource group names containing AVD resources.

.PARAMETER SubscriptionId
    Azure subscription ID. If omitted, uses the current Az context.

.PARAMETER UseManagedIdentity
    Authenticate using Managed Identity instead of interactive login.
    Use this when running as an Azure Automation Runbook.

.PARAMETER OutputPath
    Path for the generated Markdown file. Defaults to ./AVD-Documentation.md.

.PARAMETER IncludeStorageAccounts
    Also document storage accounts in the specified resource groups (useful for FSLogix profile containers).

.EXAMPLE
    .\Get-AVDDocumentation.ps1 -ResourceGroupNames "rg-avd-prod", "rg-avd-dev"

.EXAMPLE
    .\Get-AVDDocumentation.ps1 -ResourceGroupNames "rg-avd-prod" -SubscriptionId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -IncludeStorageAccounts

.EXAMPLE
    .\Get-AVDDocumentation.ps1 -ResourceGroupNames "rg-avd-prod" -UseManagedIdentity -OutputPath "/output/avd-docs.md"

.NOTES
    Required modules: Az.Accounts, Az.DesktopVirtualization, Az.Compute, Az.Network, Az.Storage (if -IncludeStorageAccounts)
#>

#Requires -Modules Az.Accounts, Az.DesktopVirtualization, Az.Compute, Az.Network

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string[]]$ResourceGroupNames,

    [Parameter(Mandatory = $false)]
    [string]$SubscriptionId,

    [Parameter(Mandatory = $false)]
    [switch]$UseManagedIdentity,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "./AVD-Documentation.md",

    [Parameter(Mandatory = $false)]
    [switch]$IncludeStorageAccounts
)

# ============================================================================
# DATA COLLECTION FUNCTIONS
# ============================================================================

function Get-AVDWorkspaceInfo {
    param ([string[]]$ResourceGroupNames)

    $workspaces = @()
    foreach ($rg in $ResourceGroupNames) {
        try {
            $ws = Get-AzWvdWorkspace -ResourceGroupName $rg -ErrorAction SilentlyContinue
            foreach ($w in $ws) {
                $appGroupNames = @()
                if ($w.ApplicationGroupReference) {
                    $appGroupNames = $w.ApplicationGroupReference | ForEach-Object { ($_ -split '/')[-1] }
                }
                $workspaces += [PSCustomObject]@{
                    Name              = $w.Name
                    FriendlyName      = $w.FriendlyName
                    ResourceGroup     = $rg
                    Location          = $w.Location
                    ApplicationGroups = $appGroupNames -join ", "
                }
            }
        }
        catch {
            Write-Host "  Warning: Could not get workspaces from $rg : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    return $workspaces
}

function Get-AVDHostPoolInfo {
    param ([string[]]$ResourceGroupNames)

    $hostPools = @()
    foreach ($rg in $ResourceGroupNames) {
        try {
            $hps = Get-AzWvdHostPool -ResourceGroupName $rg -ErrorAction SilentlyContinue
            foreach ($hp in $hps) {
                $hostPools += [PSCustomObject]@{
                    Name                  = $hp.Name
                    FriendlyName          = $hp.FriendlyName
                    ResourceGroup         = $rg
                    Location              = $hp.Location
                    HostPoolType          = $hp.HostPoolType
                    LoadBalancerType      = $hp.LoadBalancerType
                    MaxSessionLimit       = $hp.MaxSessionLimit
                    ValidationEnvironment = $hp.ValidationEnvironment
                    PreferredAppGroupType = $hp.PreferredAppGroupType
                    StartVMOnConnect      = $hp.StartVMOnConnect
                    Id                    = $hp.Id
                }
            }
        }
        catch {
            Write-Host "  Warning: Could not get host pools from $rg : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    return $hostPools
}

function Get-AVDAppGroupInfo {
    param ([string[]]$ResourceGroupNames)

    $appGroups = @()
    foreach ($rg in $ResourceGroupNames) {
        try {
            $ags = Get-AzWvdApplicationGroup -ResourceGroupName $rg -ErrorAction SilentlyContinue
            foreach ($ag in $ags) {
                $hostPoolName = ($ag.HostPoolArmPath -split '/')[-1]
                $workspaceName = ""
                if ($ag.WorkspaceArmPath) {
                    $workspaceName = ($ag.WorkspaceArmPath -split '/')[-1]
                }

                # Get Desktop Virtualization User role assignments
                $accessGroups = @()
                try {
                    $assignments = Get-AzRoleAssignment -Scope $ag.Id -ErrorAction SilentlyContinue |
                        Where-Object { $_.RoleDefinitionName -eq "Desktop Virtualization User" }
                    foreach ($assignment in $assignments) {
                        $accessGroups += $assignment.DisplayName
                    }
                }
                catch {
                    Write-Host "    Warning: Could not get role assignments for $($ag.Name)" -ForegroundColor Yellow
                }

                $appGroups += [PSCustomObject]@{
                    Name         = $ag.Name
                    FriendlyName = $ag.FriendlyName
                    ResourceGroup = $rg
                    Type         = $ag.ApplicationGroupType
                    HostPool     = $hostPoolName
                    Workspace    = $workspaceName
                    AccessGroups = $accessGroups -join "; "
                }
            }
        }
        catch {
            Write-Host "  Warning: Could not get application groups from $rg : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    return $appGroups
}

function Get-AVDSessionHostInfo {
    param ([PSCustomObject[]]$HostPools)

    $sessionHosts = @()
    $networkData = @()

    foreach ($hp in $HostPools) {
        Write-Host "  Collecting session hosts for: $($hp.Name)" -ForegroundColor Gray
        try {
            $shList = Get-AzWvdSessionHost -ResourceGroupName $hp.ResourceGroup -HostPoolName $hp.Name -ErrorAction SilentlyContinue
            foreach ($sh in $shList) {
                $fullName = $sh.Name.Split('/')[-1]
                $hostName = $fullName.Split('.')[0]

                # VM details
                $vmSize = "N/A"
                $osDiskSize = "N/A"
                $osType = "N/A"
                $privateIp = "N/A"
                $subnetName = "N/A"
                $vnetName = "N/A"
                $publicIp = "N/A"

                try {
                    $vm = Get-AzVM -ResourceGroupName $hp.ResourceGroup -Name $hostName -ErrorAction SilentlyContinue
                    if (-not $vm) {
                        # Session host VM might be in a different RG — try via resource ID
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

                        # Network interface
                        $nicId = $vm.NetworkProfile.NetworkInterfaces[0].Id
                        if ($nicId) {
                            $nic = Get-AzNetworkInterface -ResourceId $nicId -ErrorAction Stop
                            if ($nic) {
                                $ipConfig = $nic.IpConfigurations[0]
                                $privateIp = $ipConfig.PrivateIpAddress

                                # Parse subnet and VNet from subnet ID
                                if ($ipConfig.Subnet.Id) {
                                    $subnetParts = $ipConfig.Subnet.Id -split '/'
                                    $vnetName = $subnetParts[8]
                                    $subnetName = $subnetParts[10]

                                    $networkData += [PSCustomObject]@{
                                        VNetName   = $vnetName
                                        SubnetName = $subnetName
                                        SubnetId   = $ipConfig.Subnet.Id
                                        VNetRg     = $subnetParts[4]
                                    }
                                }

                                # Public IP
                                if ($ipConfig.PublicIpAddress) {
                                    $pip = Get-AzPublicIpAddress -ResourceId $ipConfig.PublicIpAddress.Id -ErrorAction SilentlyContinue
                                    if ($pip) { $publicIp = $pip.IpAddress }
                                }
                            }
                            else {
                                Write-Host "    Warning: NIC returned null for $hostName (NIC ID: $nicId)" -ForegroundColor Yellow
                            }
                        }
                    }
                    else {
                        Write-Host "    Warning: Could not find VM '$hostName' — check Reader role on the VM resource group" -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "    Warning: Could not get VM/NIC details for $hostName : $($_.Exception.Message)" -ForegroundColor Yellow
                    # If NIC lookup failed but we have the VM, try to parse network info from the VM's NIC resource ID
                    if ($vm -and $vm.NetworkProfile.NetworkInterfaces[0].Id) {
                        Write-Host "    Info: NIC read failed (likely missing Network Reader role). Falling back to resource ID parsing." -ForegroundColor DarkYellow
                        try {
                            $nicParts = $vm.NetworkProfile.NetworkInterfaces[0].Id -split '/'
                            # NIC resource ID doesn't contain subnet info, but we can still try Get-AzVM with -Status
                            # to at least get the NIC RG for the warning message
                            Write-Host "    Info: Ensure Reader role on the NIC resource group: $($nicParts[4])" -ForegroundColor DarkYellow
                        }
                        catch {}
                    }
                }

                # Assigned user (populated for Personal host pools)
                $assignedUser = if ($sh.AssignedUser) { $sh.AssignedUser } else { "" }

                $sessionHosts += [PSCustomObject]@{
                    HostPool      = $hp.Name
                    HostPoolType  = $hp.HostPoolType
                    HostName      = $hostName
                    Status        = $sh.Status
                    AssignedUser  = $assignedUser
                    Sessions      = $sh.Session
                    LastHeartBeat = $sh.LastHeartBeat
                    VMSize        = $vmSize
                    OSDiskSize    = $osDiskSize
                    OSType        = $osType
                    PrivateIP     = $privateIp
                    Subnet        = $subnetName
                    VNet          = $vnetName
                    PublicIP      = $publicIp
                }
            }
        }
        catch {
            Write-Host "  Warning: Could not get session hosts for $($hp.Name): $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    return @{ SessionHosts = $sessionHosts; NetworkData = $networkData }
}

function Get-AVDNetworkInfo {
    param ([PSCustomObject[]]$NetworkData)

    $vnets = @()

    if (-not $NetworkData -or $NetworkData.Count -eq 0) {
        Write-Host "    No network data collected from session hosts — NIC lookups may have failed" -ForegroundColor Yellow
        return $vnets
    }

    $uniqueVNets = $NetworkData | Select-Object VNetName, VNetRg -Unique

    foreach ($entry in $uniqueVNets) {
        Write-Host "    Querying VNet: $($entry.VNetName) in RG: $($entry.VNetRg)" -ForegroundColor Gray
        try {
            $vnet = Get-AzVirtualNetwork -ResourceGroupName $entry.VNetRg -Name $entry.VNetName -ErrorAction Stop
            $subnets = @()
            foreach ($subnet in $vnet.Subnets) {
                $nsgName = "None"
                if ($subnet.NetworkSecurityGroup) {
                    $nsgName = ($subnet.NetworkSecurityGroup.Id -split '/')[-1]
                }
                $subnets += [PSCustomObject]@{
                    Name          = $subnet.Name
                    AddressPrefix = $subnet.AddressPrefix -join ", "
                    NSG           = $nsgName
                }
            }

            $vnets += [PSCustomObject]@{
                Name          = $vnet.Name
                ResourceGroup = $entry.VNetRg
                Location      = $vnet.Location
                AddressSpace  = $vnet.AddressSpace.AddressPrefixes -join ", "
                Subnets       = $subnets
            }
        }
        catch {
            Write-Host "    Warning: Could not read VNet '$($entry.VNetName)' in RG '$($entry.VNetRg)': $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "    Info: Ensure Reader role on the VNet resource group '$($entry.VNetRg)' — VNets are often in a separate networking RG" -ForegroundColor DarkYellow

            # Fall back to the subnet names we already parsed from NIC data
            $knownSubnets = $NetworkData | Where-Object { $_.VNetName -eq $entry.VNetName } | Select-Object SubnetName -Unique
            $fallbackSubnets = $knownSubnets | ForEach-Object {
                [PSCustomObject]@{
                    Name          = $_.SubnetName
                    AddressPrefix = "(unavailable — no read access)"
                    NSG           = "(unavailable)"
                }
            }

            $vnets += [PSCustomObject]@{
                Name          = $entry.VNetName
                ResourceGroup = $entry.VNetRg
                Location      = "(unavailable)"
                AddressSpace  = "(unavailable — no read access to VNet)"
                Subnets       = $fallbackSubnets
            }
        }
    }
    return $vnets
}

function Get-AVDStorageInfo {
    param ([string[]]$ResourceGroupNames)

    $storageAccounts = @()
    foreach ($rg in $ResourceGroupNames) {
        try {
            $accounts = Get-AzStorageAccount -ResourceGroupName $rg -ErrorAction SilentlyContinue
            foreach ($sa in $accounts) {
                $shares = @()
                try {
                    $ctx = $sa.Context
                    $fileShares = Get-AzStorageShare -Context $ctx -ErrorAction SilentlyContinue
                    $shares = $fileShares | ForEach-Object { $_.Name }
                }
                catch {
                    Write-Host "    Warning: Could not list file shares for $($sa.StorageAccountName)" -ForegroundColor Yellow
                }

                $storageAccounts += [PSCustomObject]@{
                    Name          = $sa.StorageAccountName
                    ResourceGroup = $rg
                    Location      = $sa.Location
                    Kind          = $sa.Kind
                    Sku           = $sa.Sku.Name
                    FileShares    = $shares -join ", "
                }
            }
        }
        catch {
            Write-Host "  Warning: Could not get storage accounts from $rg : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    return $storageAccounts
}

# ============================================================================
# MARKDOWN GENERATION
# ============================================================================

function Build-AVDDocumentation {
    param (
        [PSCustomObject[]]$Workspaces,
        [PSCustomObject[]]$HostPools,
        [PSCustomObject[]]$AppGroups,
        [PSCustomObject[]]$SessionHosts,
        [PSCustomObject[]]$VNets,
        [PSCustomObject[]]$StorageAccounts,
        [string]$SubscriptionName,
        [string]$SubscriptionId
    )

    $md = [System.Text.StringBuilder]::new()

    # Header
    [void]$md.AppendLine("# Azure Virtual Desktop Documentation")
    [void]$md.AppendLine("")
    [void]$md.AppendLine("> Auto-generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm') | Subscription: $SubscriptionName ($SubscriptionId)")
    [void]$md.AppendLine("")

    # Table of Contents
    [void]$md.AppendLine("## Table of Contents")
    [void]$md.AppendLine("")
    [void]$md.AppendLine("- [Summary](#summary)")
    [void]$md.AppendLine("- [Workspaces](#workspaces)")
    [void]$md.AppendLine("- [Host Pools](#host-pools)")
    [void]$md.AppendLine("- [Application Groups](#application-groups)")
    [void]$md.AppendLine("- [Networking](#networking)")
    if ($StorageAccounts) {
        [void]$md.AppendLine("- [Storage Accounts](#storage-accounts)")
    }
    [void]$md.AppendLine("")

    # Summary
    [void]$md.AppendLine("## Summary")
    [void]$md.AppendLine("")
    [void]$md.AppendLine("| Metric | Count |")
    [void]$md.AppendLine("|--------|-------|")
    [void]$md.AppendLine("| Workspaces | $($Workspaces.Count) |")
    [void]$md.AppendLine("| Host Pools | $($HostPools.Count) |")
    [void]$md.AppendLine("| Application Groups | $($AppGroups.Count) |")
    [void]$md.AppendLine("| Session Hosts | $($SessionHosts.Count) |")
    [void]$md.AppendLine("| Virtual Networks | $($VNets.Count) |")
    if ($StorageAccounts) {
        [void]$md.AppendLine("| Storage Accounts | $($StorageAccounts.Count) |")
    }
    [void]$md.AppendLine("")

    # Workspaces
    [void]$md.AppendLine("## Workspaces")
    [void]$md.AppendLine("")
    if ($Workspaces.Count -eq 0) {
        [void]$md.AppendLine("No workspaces found.")
    }
    else {
        [void]$md.AppendLine("| Name | Friendly Name | Resource Group | Location | Application Groups |")
        [void]$md.AppendLine("|------|---------------|----------------|----------|--------------------|")
        foreach ($ws in $Workspaces) {
            [void]$md.AppendLine("| $($ws.Name) | $($ws.FriendlyName) | $($ws.ResourceGroup) | $($ws.Location) | $($ws.ApplicationGroups) |")
        }
    }
    [void]$md.AppendLine("")

    # Host Pools
    [void]$md.AppendLine("## Host Pools")
    [void]$md.AppendLine("")
    if ($HostPools.Count -eq 0) {
        [void]$md.AppendLine("No host pools found.")
    }
    else {
        foreach ($hp in $HostPools) {
            [void]$md.AppendLine("### $($hp.Name)")
            [void]$md.AppendLine("")
            if ($hp.FriendlyName) {
                [void]$md.AppendLine("**Friendly Name:** $($hp.FriendlyName)  ")
            }
            [void]$md.AppendLine("**Resource Group:** $($hp.ResourceGroup)  ")
            [void]$md.AppendLine("**Location:** $($hp.Location)  ")
            [void]$md.AppendLine("**Type:** $($hp.HostPoolType)  ")
            [void]$md.AppendLine("**Load Balancer:** $($hp.LoadBalancerType)  ")
            [void]$md.AppendLine("**Max Session Limit:** $($hp.MaxSessionLimit)  ")
            [void]$md.AppendLine("**Preferred App Group Type:** $($hp.PreferredAppGroupType)  ")
            [void]$md.AppendLine("**Start VM On Connect:** $($hp.StartVMOnConnect)  ")
            [void]$md.AppendLine("**Validation Environment:** $($hp.ValidationEnvironment)  ")
            [void]$md.AppendLine("")

            # Session hosts for this host pool
            $hpSessionHosts = $SessionHosts | Where-Object { $_.HostPool -eq $hp.Name }
            $isPersonal = $hp.HostPoolType -eq "Personal"
            if ($hpSessionHosts.Count -gt 0) {
                [void]$md.AppendLine("#### Session Hosts ($($hpSessionHosts.Count))")
                [void]$md.AppendLine("")
                if ($isPersonal) {
                    [void]$md.AppendLine("| Hostname | Status | Assigned User | VM Size | OS | Disk | Private IP | Subnet | VNet |")
                    [void]$md.AppendLine("|----------|--------|---------------|---------|----|----- |------------|--------|------|")
                }
                else {
                    [void]$md.AppendLine("| Hostname | Status | VM Size | OS | Disk | Private IP | Subnet | VNet |")
                    [void]$md.AppendLine("|----------|--------|---------|----|------|------------|--------|------|")
                }
                foreach ($sh in $hpSessionHosts) {
                    if ($isPersonal) {
                        [void]$md.AppendLine("| $($sh.HostName) | $($sh.Status) | $($sh.AssignedUser) | $($sh.VMSize) | $($sh.OSType) | $($sh.OSDiskSize) | $($sh.PrivateIP) | $($sh.Subnet) | $($sh.VNet) |")
                    }
                    else {
                        [void]$md.AppendLine("| $($sh.HostName) | $($sh.Status) | $($sh.VMSize) | $($sh.OSType) | $($sh.OSDiskSize) | $($sh.PrivateIP) | $($sh.Subnet) | $($sh.VNet) |")
                    }
                }
                [void]$md.AppendLine("")
            }
            else {
                [void]$md.AppendLine("*No session hosts found.*")
                [void]$md.AppendLine("")
            }
        }
    }

    # Application Groups
    [void]$md.AppendLine("## Application Groups")
    [void]$md.AppendLine("")
    if ($AppGroups.Count -eq 0) {
        [void]$md.AppendLine("No application groups found.")
    }
    else {
        [void]$md.AppendLine("| Name | Type | Host Pool | Workspace | Access Groups |")
        [void]$md.AppendLine("|------|------|-----------|-----------|---------------|")
        foreach ($ag in $AppGroups) {
            [void]$md.AppendLine("| $($ag.Name) | $($ag.Type) | $($ag.HostPool) | $($ag.Workspace) | $($ag.AccessGroups) |")
        }
    }
    [void]$md.AppendLine("")

    # Networking
    [void]$md.AppendLine("## Networking")
    [void]$md.AppendLine("")
    if ($VNets.Count -eq 0) {
        [void]$md.AppendLine("No virtual networks found.")
    }
    else {
        foreach ($vnet in $VNets) {
            [void]$md.AppendLine("### $($vnet.Name)")
            [void]$md.AppendLine("")
            [void]$md.AppendLine("**Resource Group:** $($vnet.ResourceGroup)  ")
            [void]$md.AppendLine("**Location:** $($vnet.Location)  ")
            [void]$md.AppendLine("**Address Space:** $($vnet.AddressSpace)  ")
            [void]$md.AppendLine("")

            if ($vnet.Subnets.Count -gt 0) {
                [void]$md.AppendLine("| Subnet | Address Prefix | NSG |")
                [void]$md.AppendLine("|--------|----------------|-----|")
                foreach ($subnet in $vnet.Subnets) {
                    [void]$md.AppendLine("| $($subnet.Name) | $($subnet.AddressPrefix) | $($subnet.NSG) |")
                }
                [void]$md.AppendLine("")
            }
        }
    }

    # Storage Accounts
    if ($StorageAccounts) {
        [void]$md.AppendLine("## Storage Accounts")
        [void]$md.AppendLine("")
        if ($StorageAccounts.Count -eq 0) {
            [void]$md.AppendLine("No storage accounts found.")
        }
        else {
            [void]$md.AppendLine("| Name | Resource Group | Location | Kind | SKU | File Shares |")
            [void]$md.AppendLine("|------|----------------|----------|------|-----|-------------|")
            foreach ($sa in $StorageAccounts) {
                [void]$md.AppendLine("| $($sa.Name) | $($sa.ResourceGroup) | $($sa.Location) | $($sa.Kind) | $($sa.Sku) | $($sa.FileShares) |")
            }
        }
        [void]$md.AppendLine("")
    }

    return $md.ToString()
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  AVD Documentation Generator" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# --- Authentication ---
Write-Host "`nAuthenticating..." -ForegroundColor Cyan
if ($UseManagedIdentity) {
    Connect-AzAccount -Identity | Out-Null
    Write-Host "  Connected via Managed Identity" -ForegroundColor Green
}
else {
    Connect-AzAccount | Out-Null
    Write-Host "  Connected interactively" -ForegroundColor Green
}

# --- Subscription ---
if ($SubscriptionId) {
    Set-AzContext -SubscriptionId $SubscriptionId | Out-Null
}
$context = Get-AzContext
$subName = $context.Subscription.Name
$subId = $context.Subscription.Id
Write-Host "  Subscription: $subName ($subId)" -ForegroundColor Green

# --- Data Collection ---
Write-Host "`nCollecting AVD data..." -ForegroundColor Cyan

Write-Host "  [1/6] Workspaces..." -ForegroundColor White
$workspaces = Get-AVDWorkspaceInfo -ResourceGroupNames $ResourceGroupNames
Write-Host "    Found $($workspaces.Count) workspace(s)" -ForegroundColor Gray

Write-Host "  [2/6] Host Pools..." -ForegroundColor White
$hostPools = Get-AVDHostPoolInfo -ResourceGroupNames $ResourceGroupNames
Write-Host "    Found $($hostPools.Count) host pool(s)" -ForegroundColor Gray

Write-Host "  [3/6] Application Groups..." -ForegroundColor White
$appGroups = Get-AVDAppGroupInfo -ResourceGroupNames $ResourceGroupNames
Write-Host "    Found $($appGroups.Count) application group(s)" -ForegroundColor Gray

Write-Host "  [4/6] Session Hosts & VM Details..." -ForegroundColor White
$shResult = Get-AVDSessionHostInfo -HostPools $hostPools
$sessionHosts = $shResult.SessionHosts
$networkData = $shResult.NetworkData
Write-Host "    Found $($sessionHosts.Count) session host(s)" -ForegroundColor Gray

Write-Host "  [5/6] Networking..." -ForegroundColor White
$vnets = Get-AVDNetworkInfo -NetworkData $networkData
Write-Host "    Found $($vnets.Count) VNet(s)" -ForegroundColor Gray

$storageAccounts = $null
if ($IncludeStorageAccounts) {
    Write-Host "  [6/6] Storage Accounts..." -ForegroundColor White
    $storageAccounts = Get-AVDStorageInfo -ResourceGroupNames $ResourceGroupNames
    Write-Host "    Found $($storageAccounts.Count) storage account(s)" -ForegroundColor Gray
}
else {
    Write-Host "  [6/6] Storage Accounts... skipped (-IncludeStorageAccounts not set)" -ForegroundColor DarkGray
}

# --- Generate Documentation ---
Write-Host "`nGenerating Markdown documentation..." -ForegroundColor Cyan

$markdown = Build-AVDDocumentation `
    -Workspaces $workspaces `
    -HostPools $hostPools `
    -AppGroups $appGroups `
    -SessionHosts $sessionHosts `
    -VNets $vnets `
    -StorageAccounts $storageAccounts `
    -SubscriptionName $subName `
    -SubscriptionId $subId

$markdown | Out-File -FilePath $OutputPath -Encoding utf8
Write-Host "  Written to: $OutputPath" -ForegroundColor Green

# --- Summary ---
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "  Documentation generated successfully!" -ForegroundColor Green
Write-Host "  Workspaces:       $($workspaces.Count)" -ForegroundColor White
Write-Host "  Host Pools:       $($hostPools.Count)" -ForegroundColor White
Write-Host "  App Groups:       $($appGroups.Count)" -ForegroundColor White
Write-Host "  Session Hosts:    $($sessionHosts.Count)" -ForegroundColor White
Write-Host "  VNets:            $($vnets.Count)" -ForegroundColor White
if ($storageAccounts) {
    Write-Host "  Storage Accounts: $($storageAccounts.Count)" -ForegroundColor White
}
Write-Host "==========================================" -ForegroundColor Cyan
