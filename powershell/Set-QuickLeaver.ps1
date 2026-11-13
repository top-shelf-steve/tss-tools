<#
.SYNOPSIS
    Quick leaver: cuts a departing user's access as fast as possible.
    Designed for a hybrid tenant (AD-synced users, pass-through auth, SSPR with writeback).

.DESCRIPTION
    Enter a UPN and the script runs these steps in order:

      1. BLOCK SIGN-IN    Adds the user to a cloud-only "leavers" group that a Conditional
                          Access policy blocks from all resources. Entra won't let you
                          disable a synced account, so this is the cloud-side equivalent.
      2. AUTH METHODS     Deletes every registered authentication method (Authenticator,
                          phone, FIDO2, Windows Hello, OATH, email, TAP, ...). Graph has no
                          separate "require re-register MFA" call; deleting every method is
                          how it is done. It also means the user can't use SSPR to set
                          themselves a new password.
      3. REVOKE SESSIONS  Invalidates all refresh tokens and browser session cookies. Runs
                          last, so it also catches anything issued while the script ran.

    Every step runs even if an earlier one failed. A summary prints at the end, and if any
    step failed the script throws, so the runbook job shows as Failed.

    WHY THERE IS NO PASSWORD RESET
    With pass-through authentication, the password that matters is the one in on-prem AD.
    Changing it only in the cloud doesn't lock anyone out. The only Graph API that writes a
    reset back to AD (authenticationMethod: resetPassword) supports delegated permissions
    only. A managed identity can't call it. For synced users, the summary therefore lists a
    manual password reset as a follow-up. An admin reset in the Entra admin center does write
    back to AD; one in the Microsoft 365 admin center does not.

.PARAMETER UserPrincipalName
    UPN of the departing user, e.g. jdoe@contoso.com.

.PARAMETER BlockSignInGroupId
    Object ID of the cloud-only group targeted by the "block all access" Conditional Access
    policy. If omitted, the value from the DEFAULTS section is used.

.PARAMETER UserAssignedIdentityClientId
    Client ID of a user-assigned managed identity. Omit to use the system-assigned identity.

.PARAMETER DryRun
    Look everything up and report what would happen, without changing anything.

.PARAMETER ForceInteractive
    Use interactive sign-in even if a managed identity endpoint is detected.

.EXAMPLE
    .\Set-QuickLeaver.ps1 -UserPrincipalName "jdoe@contoso.com" -DryRun
    # Local test: signs you in interactively, shows what would happen, changes nothing.

.EXAMPLE
    .\Set-QuickLeaver.ps1 -UserPrincipalName "test.leaver@contoso.com"
    # Local run against a TEST account (never a real person) to check the steps end to end.

.NOTES
    Required module:  Microsoft.Graph.Authentication (the only Graph module this script uses)

    ONE-TIME SETUP
      1. Create a cloud-only security group, e.g. "SG-Leavers-BlockSignIn". It must not be
         synced from AD (the cloud can't change a synced group's membership) and must not be
         role-assignable. Put its object ID in $DefaultBlockSignInGroupId below.
      2. Create a Conditional Access policy:
           Users:            include that group (exclude your break-glass accounts)
           Target resources: All resources
           Grant:            Block access
         Run it in Report-only first, test it with a test account, then switch it On.
      3. Grant the managed identity these Microsoft Graph APPLICATION permissions. The
         portal can't do this for managed identities; use
         New-MgServicePrincipalAppRoleAssignment.
           User.Read.All                           look up the user
           GroupMember.ReadWrite.All               add the user to the block group
           UserAuthenticationMethod.ReadWrite.All  list and delete authentication methods
           User.RevokeSessions.All                 revoke sessions
      4. Import Microsoft.Graph.Authentication into the Automation runtime environment.

    SECURITY
      UserAuthenticationMethod.ReadWrite.All (app-only) applies to every user, admins
      included, and also allows creating Temporary Access Passes. Anyone who can edit this
      runbook can act as the identity, so restrict who has Contributor rights on the
      Automation account.

    LOCAL RUNS
      Signs in interactively with delegated scopes, so your own Entra roles decide what
      succeeds (e.g. Authentication Administrator to delete another user's methods).
#>

#Requires -Modules Microsoft.Graph.Authentication

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$BlockSignInGroupId,

    [Parameter(Mandatory = $false)]
    [string]$UserAssignedIdentityClientId,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [switch]$ForceInteractive
)

$ErrorActionPreference = 'Stop'

# ========================
# DEFAULTS - edit these to match your environment.
# Anything passed via parameters overrides what's set here.
# ========================
$DefaultBlockSignInGroupId = ""   # object ID of the cloud-only group the block CA policy targets

if (-not $BlockSignInGroupId) { $BlockSignInGroupId = $DefaultBlockSignInGroupId }
$UserPrincipalName = $UserPrincipalName.Trim()

$GraphScopes = @(
    'User.Read.All',
    'GroupMember.ReadWrite.All',
    'UserAuthenticationMethod.ReadWrite.All',
    'User.RevokeSessions.All'
)

# Maps the @odata.type returned by /authentication/methods to the endpoint used to delete it.
# Password isn't listed; it can't be deleted.
$MethodEndpoints = @{
    '#microsoft.graph.phoneAuthenticationMethod'                   = 'phoneMethods'
    '#microsoft.graph.emailAuthenticationMethod'                   = 'emailMethods'
    '#microsoft.graph.fido2AuthenticationMethod'                   = 'fido2Methods'
    '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'  = 'microsoftAuthenticatorMethods'
    '#microsoft.graph.softwareOathAuthenticationMethod'            = 'softwareOathMethods'
    '#microsoft.graph.temporaryAccessPassAuthenticationMethod'     = 'temporaryAccessPassMethods'
    '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' = 'windowsHelloForBusinessMethods'
    '#microsoft.graph.platformCredentialAuthenticationMethod'      = 'platformCredentialMethods'
    '#microsoft.graph.externalAuthenticationMethod'                = 'externalAuthenticationMethods'
}

# ========================
# RUNTIME DETECTION
# ========================
$InAutomation = [bool]($PSPrivateMetadata.JobId -or $env:AUTOMATION_ASSET_ACCOUNTID)
$UseManagedIdentity = (-not $ForceInteractive) -and ($InAutomation -or $env:IDENTITY_ENDPOINT -or $env:MSI_ENDPOINT)

# ========================
# HELPERS
# ========================
$Results = [System.Collections.Generic.List[object]]::new()

function Write-Log {
    # In Automation, everything goes to the output stream so it shows on the job's Output
    # tab, in order. Locally, use colored host output.
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warn', 'Error', 'Header', 'Detail')]
        [string]$Level = 'Info'
    )
    if ($InAutomation) { Write-Output $Message; return }
    $color = switch ($Level) {
        'Success' { 'Green' }
        'Warn'    { 'Yellow' }
        'Error'   { 'Red' }
        'Header'  { 'Cyan' }
        'Detail'  { 'Gray' }
        default   { 'White' }
    }
    Write-Host $Message -ForegroundColor $color
}

function Add-Result {
    param([string]$Step, [ValidateSet('Success', 'Failed', 'Skipped', 'ActionRequired', 'DryRun')][string]$Status, [string]$Detail)
    $Results.Add([pscustomobject]@{ Step = $Step; Status = $Status; Detail = $Detail })
    $level = switch ($Status) {
        'Success'        { 'Success' }
        'Failed'         { 'Error' }
        'ActionRequired' { 'Warn' }
        default          { 'Detail' }
    }
    Write-Log "  [$Status] $Detail" $level
}

function Get-ErrorText {
    # Pulls the Graph error message out of an Invoke-MgGraphRequest failure when there is one.
    param($ErrorRecord)
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        try { return ($ErrorRecord.ErrorDetails.Message | ConvertFrom-Json).error.message }
        catch { return $ErrorRecord.ErrorDetails.Message }
    }
    return $ErrorRecord.Exception.Message
}

function Get-MethodLabel {
    # Short, human-readable label. Phone numbers and emails are left out so the job log
    # doesn't collect personal data.
    param($Method)
    $type = ($Method.'@odata.type' -replace '^#microsoft\.graph\.', '') -replace 'AuthenticationMethod$', ''
    $extra = @($Method.phoneType, $Method.displayName, $Method.model) | Where-Object { $_ }
    if ($extra) { return "$type ($($extra -join ', '))" }
    return $type
}

# ========================
# AUTHENTICATION
# ========================
Write-Log "==========================================" Header
Write-Log "  Quick Leaver" Header
if ($DryRun) { Write-Log "  DRY RUN - no changes will be made" Header }
Write-Log "==========================================" Header

Write-Log "`nConnecting to Microsoft Graph..." Header
if ($UseManagedIdentity) {
    $mgParams = @{ Identity = $true; NoWelcome = $true }
    if ($UserAssignedIdentityClientId) { $mgParams['ClientId'] = $UserAssignedIdentityClientId }
    Connect-MgGraph @mgParams
}
else {
    Connect-MgGraph -Scopes $GraphScopes -NoWelcome
}
$ctx = Get-MgContext
$who = if ($ctx.Account) { $ctx.Account } else { "managed identity (client $($ctx.ClientId))" }
Write-Log "  Connected as $who" Success

# ========================
# RESOLVE USER
# ========================
Write-Log "`nResolving user..." Header
$select = 'id,displayName,userPrincipalName,accountEnabled,jobTitle,department,' +
          'onPremisesSyncEnabled,onPremisesSamAccountName,onPremisesDomainName'
try {
    $user = Invoke-MgGraphRequest -Method GET -Uri "v1.0/users/$($UserPrincipalName.Replace('#', '%23'))?`$select=$select"
}
catch {
    throw "Could not find user '$UserPrincipalName': $(Get-ErrorText $_)"
}

if ($ctx.Account -and $ctx.Account -eq $user.userPrincipalName) {
    throw "Refusing to run against your own account ($($ctx.Account))."
}

$isSynced = [bool]$user.onPremisesSyncEnabled
Write-Log "  $($user.displayName) <$($user.userPrincipalName)>" Success
Write-Log "  Object ID:  $($user.id)" Detail
Write-Log "  Title/Dept: $($user.jobTitle) / $($user.department)" Detail
Write-Log "  Source:     $(if ($isSynced) { "Synced from AD ($($user.onPremisesDomainName)\$($user.onPremisesSamAccountName))" } else { 'Cloud-only' })" Detail
Write-Log "  Enabled:    $($user.accountEnabled)" Detail

# ========================
# 1. BLOCK SIGN-IN (Conditional Access group)
# ========================
$step = 'Block sign-in (CA group)'
Write-Log "`n[1/3] $step" Header
if (-not $BlockSignInGroupId) {
    Add-Result $step 'ActionRequired' 'No block group configured. Set $DefaultBlockSignInGroupId or pass -BlockSignInGroupId.'
}
else {
    try {
        $group = Invoke-MgGraphRequest -Method GET -Uri "v1.0/groups/$BlockSignInGroupId`?`$select=id,displayName,onPremisesSyncEnabled"
        if ($group.onPremisesSyncEnabled) {
            throw "Group '$($group.displayName)' is synced from AD, so its membership can't be changed from the cloud. Use a cloud-only group."
        }
        if ($DryRun) {
            Add-Result $step 'DryRun' "Would add to '$($group.displayName)'"
        }
        else {
            $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.id)" } | ConvertTo-Json
            try {
                Invoke-MgGraphRequest -Method POST -Uri "v1.0/groups/$($group.id)/members/`$ref" -Body $body -ContentType 'application/json'
                Add-Result $step 'Success' "Added to '$($group.displayName)'"
            }
            catch {
                if ((Get-ErrorText $_) -match 'already exist') {
                    Add-Result $step 'Success' "Already a member of '$($group.displayName)'"
                }
                else { throw }
            }
        }
    }
    catch {
        Add-Result $step 'Failed' (Get-ErrorText $_)
    }
}

# ========================
# 2. REMOVE AUTHENTICATION METHODS (= require re-register MFA, kills SSPR)
# ========================
$step = 'Remove auth methods'
Write-Log "`n[2/3] Remove all authentication methods" Header
try {
    $methods  = @((Invoke-MgGraphRequest -Method GET -Uri "v1.0/users/$($user.id)/authentication/methods").value)
    $toRemove = @($methods | Where-Object { $_.'@odata.type' -ne '#microsoft.graph.passwordAuthenticationMethod' })
    $unknown  = @($toRemove | Where-Object { -not $MethodEndpoints.ContainsKey($_.'@odata.type') })
    $pending  = @($toRemove | Where-Object { $MethodEndpoints.ContainsKey($_.'@odata.type') })

    foreach ($m in $toRemove) { Write-Log "  Found: $(Get-MethodLabel $m)" Detail }

    if ($toRemove.Count -eq 0) {
        Add-Result $step 'Success' 'No methods registered (password only).'
    }
    elseif ($DryRun) {
        Add-Result $step 'DryRun' "Would remove $($pending.Count) method(s)"
    }
    else {
        # Graph won't delete the user's default MFA method while other methods still exist,
        # so retry anything that failed once the others are gone.
        $removed = 0
        $lastErrors = @{}
        for ($pass = 1; $pass -le 3 -and $pending.Count -gt 0; $pass++) {
            $failed = @()
            foreach ($m in $pending) {
                $segment = $MethodEndpoints[$m.'@odata.type']
                try {
                    Invoke-MgGraphRequest -Method DELETE -Uri "v1.0/users/$($user.id)/authentication/$segment/$($m.id)"
                    Write-Log "  Removed: $(Get-MethodLabel $m)" Success
                    $removed++
                }
                catch {
                    $failed += $m
                    $lastErrors[$m.id] = Get-ErrorText $_
                }
            }
            $pending = $failed
        }

        $problems = @()
        foreach ($m in $pending) { $problems += "$(Get-MethodLabel $m): $($lastErrors[$m.id])" }
        foreach ($m in $unknown) { $problems += "$(Get-MethodLabel $m): type not handled by this script, remove it in the portal" }

        if ($problems.Count -eq 0) {
            Add-Result $step 'Success' "Removed $removed method(s)"
        }
        else {
            Add-Result $step 'Failed' "Removed $removed; could not remove: $($problems -join ' | ')"
        }
    }
}
catch {
    Add-Result $step 'Failed' (Get-ErrorText $_)
}

# ========================
# 3. REVOKE SESSIONS (last, so it covers anything issued during the run)
# ========================
$step = 'Revoke sessions'
Write-Log "`n[3/3] Revoke sign-in sessions" Header
try {
    if ($DryRun) {
        Add-Result $step 'DryRun' 'Would revoke all refresh tokens and session cookies'
    }
    else {
        $null = Invoke-MgGraphRequest -Method POST -Uri "v1.0/users/$($user.id)/revokeSignInSessions"
        Add-Result $step 'Success' 'All refresh tokens and session cookies revoked'
    }
}
catch {
    Add-Result $step 'Failed' (Get-ErrorText $_)
}

# ========================
# FOLLOW-UP REMINDER (synced users only; this script makes no password or AD changes)
# ========================
if ($isSynced) {
    Write-Log "`nManual follow-up" Header
    Add-Result 'Reset password (manual)' 'ActionRequired' ("Reset the password in the Entra admin center; with writeback on it reaches AD. " +
        "Don't use the Microsoft 365 admin center, whose resets don't write back. " +
        "AD account: $($user.onPremisesDomainName)\$($user.onPremisesSamAccountName)")
}

# ========================
# SUMMARY
# ========================
Write-Log "`n==========================================" Header
Write-Log "  Summary: $($user.displayName) <$($user.userPrincipalName)>" Header
Write-Log "==========================================" Header
Write-Log ($Results | Format-Table Step, Status, Detail -AutoSize -Wrap | Out-String -Width 220).TrimEnd()

$failedSteps = @($Results | Where-Object Status -eq 'Failed')
$todoSteps   = @($Results | Where-Object Status -eq 'ActionRequired')

if ($todoSteps.Count -gt 0) {
    Write-Log "`nStill needs doing by hand:" Header
    foreach ($t in $todoSteps) { Write-Log "  - $($t.Step): $($t.Detail)" Warn }
}
if (-not $DryRun) {
    Write-Log "`nNote: tokens are cut off within minutes in apps that support Continuous Access Evaluation" Detail
    Write-Log "(Exchange, SharePoint, Teams). In other apps an existing access token can last up to about an hour." Detail
}

if ($failedSteps.Count -gt 0) {
    throw "Quick leaver finished with $($failedSteps.Count) failed step(s): $(($failedSteps.Step) -join ', ')"
}
Write-Log "`nDone." Success
