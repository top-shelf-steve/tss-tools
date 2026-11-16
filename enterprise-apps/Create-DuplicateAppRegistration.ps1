<#
.SYNOPSIS
    Creates a new app registration with the same settings as an existing one. No secrets,
    certificates or consent are copied.

.DESCRIPTION
    Reads an existing app registration and registers a new app, under the name you give,
    with the same configuration.

    COPIED
      - Supported account types
      - Authentication: Web, SPA and mobile/desktop redirect URIs, front-channel logout URL,
        home page URL, implicit grant (ID and access token) settings, "Allow public client
        flows"
      - API permissions, as configured permissions only (see NOT GRANTED below)
      - App roles, with new IDs
      - Expose an API: scopes (with new IDs), access token version, known client
        applications, accept mapped claims
      - Application ID URI: api://<source appId> becomes api://<new appId>. Any URI that
        contains the source app ID is carried over the same way.
      - Token configuration: optional claims and the groups claim
      - Branding & properties: description, notes, service management reference, and the
        marketing, privacy, support and terms of service URLs
      - Tags, parental control settings, SAML metadata URL, device-only auth, OAuth POST
        response, app instance property lock, request signature verification and
        authentication behaviors
      - If the source app requests permissions on its own API, the copy requests the same
        permissions on the new app's API.

    NOT GRANTED
      API permissions are added to the new app but show as "Not granted". The script never
      calls the consent endpoints (oauth2PermissionGrants, appRoleAssignments).

    NOT COPIED, ON PURPOSE
      - Client secrets, certificates and federated credentials. The new app has none.
      - Authorized client applications (Expose an API > Authorized client applications).
        These let client apps skip consent, so copying them would grant consent. The script
        lists them so you can add them by hand if you want them.
      - Owners, unless you use -CopyOwners.

    NOT COPIED, BECAUSE THEY CAN'T BE
      - Application ID URIs that don't contain the source app ID. They must be unique in
        Entra ID. If none can be carried over, the new app gets api://<new appId>.
      - Logo, extension properties, add-ins, and the token encryption key (it points at a
        certificate, which isn't copied).
      - Enterprise application settings such as user and group assignments, SSO and
        provisioning. Only "Assignment required?" and "Visible to users?" are copied.

    The script also creates the enterprise application (service principal) for the new app,
    as the portal does when you register an app. Creating it grants nothing. Use
    -SkipServicePrincipal to leave it out.

    Before creating anything, the script prints what it will copy and what it won't. It
    stops if an app registration with the new name already exists. If a step fails after
    the app is created, the script carries on with the remaining steps, lists what failed,
    and throws at the end. It doesn't delete the new app: fix it by hand or delete it.

.PARAMETER SourceApp
    The app registration to copy: its Application (client) ID, Object ID, or exact display
    name. If more than one app has that display name, use an ID instead.

.PARAMETER NewDisplayName
    The display name for the new app registration. Must not already be in use.

.PARAMETER CopyOwners
    Also add the source app's owners as owners of the new app registration. Off by default
    because owners can add secrets to the app.

.PARAMETER SkipServicePrincipal
    Don't create the enterprise application (service principal) for the new app.

.PARAMETER DryRun
    Read the source app and show what would be created, without creating anything. Signs
    in with read-only scope.

.EXAMPLE
    .\Create-DuplicateAppRegistration.ps1 -SourceApp "HR Portal" -NewDisplayName "HR Portal - Test" -DryRun
    # Shows what would be copied. Creates nothing.

.EXAMPLE
    .\Create-DuplicateAppRegistration.ps1 -SourceApp "11111111-1111-1111-1111-111111111111" -NewDisplayName "HR Portal - Test"
    # Creates the copy, without owners.

.EXAMPLE
    .\Create-DuplicateAppRegistration.ps1 -SourceApp "HR Portal" -NewDisplayName "HR Portal - Test" -CopyOwners
    # Creates the copy and gives it the same owners as the source app.

.NOTES
    Required module: Microsoft.Graph.Authentication

    Permissions (delegated): Application.Read.All for -DryRun. Application.ReadWrite.All
    otherwise, plus Directory.Read.All with -CopyOwners (Graph needs it to add owners).

    Your account needs a role that can create app registrations and enterprise applications,
    such as Cloud Application Administrator or Application Administrator.

    App role and scope IDs only need to be unique within one app, but the copy gets new
    IDs so the two apps stay independent. Clients use the scope and role values (for
    example api://<appId>/access_as_user), which are copied unchanged.
#>

#Requires -Modules Microsoft.Graph.Authentication

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SourceApp,

    [Parameter(Mandatory = $true)]
    [ValidateLength(1, 256)]
    [string]$NewDisplayName,

    [Parameter(Mandatory = $false)]
    [switch]$CopyOwners,

    [Parameter(Mandatory = $false)]
    [switch]$SkipServicePrincipal,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$SourceApp      = $SourceApp.Trim()
$NewDisplayName = $NewDisplayName.Trim()
if (-not $NewDisplayName) { throw 'NewDisplayName is blank.' }

$GuidPattern = '^[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}$'

# Errors Entra returns while an object the script has just created is still replicating.
$NotReadyPattern = 'does not exist|does not reference a valid application object|not found'

# App settings copied unchanged when the source app has a value for them.
$CopyAsIs = @(
    'signInAudience'
    'description'
    'notes'
    'tags'
    'groupMembershipClaims'
    'optionalClaims'
    'isFallbackPublicClient'
    'isDeviceOnlyAuthSupported'
    'oauth2RequiredPostResponse'
    'parentalControlSettings'
    'serviceManagementReference'
    'samlMetadataUrl'
    'servicePrincipalLockConfiguration'
    'requestSignatureVerification'
)

# The portal puts this tag on the service principals it creates. It's what makes the app
# show in the default Enterprise applications list.
$IntegratedAppTag = 'WindowsAzureActiveDirectoryIntegratedApp'

# ========================
# HELPERS
# ========================
$Problems  = [System.Collections.Generic.List[string]]::new()
$NotCopied = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warn', 'Error', 'Header', 'Detail')]
        [string]$Level = 'Info'
    )
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

function Get-ErrorText {
    # Pulls the Graph error message out of an Invoke-MgGraphRequest failure when there is one.
    param($ErrorRecord)
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        try { return ($ErrorRecord.ErrorDetails.Message | ConvertFrom-Json).error.message }
        catch { return $ErrorRecord.ErrorDetails.Message }
    }
    return $ErrorRecord.Exception.Message
}

function Add-Problem {
    param([string]$Message)
    $Problems.Add($Message)
    Write-Log "  FAILED: $Message" Error
}

function Get-Count {
    # Number of non-null items. Safe on $null and on single objects, unlike @($x).Count,
    # which counts $null as one item.
    param($Items)
    @($Items | Where-Object { $null -ne $_ }).Count
}

function Copy-Clean {
    # Deep-copies a value returned by Invoke-MgGraphRequest, dropping @odata.* annotations,
    # so it can be sent back in a request body. Arrays are returned wrapped so PowerShell
    # doesn't unroll them: assign the result directly, don't wrap the call in @().
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Collections.IDictionary]) {
        $copy = [ordered]@{}
        foreach ($key in $Value.Keys) {
            if ("$key" -like '@odata.*') { continue }
            $copy[$key] = Copy-Clean $Value[$key]
        }
        return $copy
    }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $list = @(foreach ($item in $Value) { Copy-Clean $item })
        return , $list
    }
    return $Value
}

function Invoke-Graph {
    # Invoke-MgGraphRequest with a JSON body. With -Retry, waits and retries a few times
    # when Entra hasn't finished replicating an object the script has just created.
    param(
        [string]$Method = 'GET',
        [string]$Uri,
        $Body,
        [switch]$Retry
    )
    $request = @{ Method = $Method; Uri = $Uri }
    if ($null -ne $Body) {
        $request.Body        = ConvertTo-Json -InputObject $Body -Depth 20 -Compress
        $request.ContentType = 'application/json'
    }
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return Invoke-MgGraphRequest @request
        }
        catch {
            $text = Get-ErrorText $_
            if (-not $Retry -or $attempt -ge 5 -or $text -notmatch $NotReadyPattern) { throw }
            $wait = $attempt * 5
            Write-Log "    Not ready yet, retrying in ${wait}s ($text)" Detail
            Start-Sleep -Seconds $wait
        }
    }
}

function Get-AllPages {
    # Returns every item from a paged Graph list. PowerShell unrolls the result, so wrap
    # calls in @().
    param([string]$Uri, [switch]$Retry)
    while ($Uri) {
        $page = Invoke-Graph -Uri $Uri -Retry:$Retry
        foreach ($item in $page.value) { $item }
        $Uri = $page.'@odata.nextLink'
    }
}

function Get-NameFilter {
    # $filter on displayName, with the name quoted and URL-encoded.
    param([string]$Name)
    '$filter=' + [uri]::EscapeDataString("displayName eq '$($Name.Replace("'", "''"))'")
}

function Get-OwnerLabel {
    param($Owner)
    if ($Owner.userPrincipalName) { return "$($Owner.displayName) ($($Owner.userPrincipalName))" }
    if ($Owner.displayName) { return $Owner.displayName }
    return $Owner.id
}

function Find-SourceApp {
    param([string]$Identity)
    if ($Identity -match $GuidPattern) {
        $lastError = $null
        foreach ($uri in "v1.0/applications(appId='$Identity')", "v1.0/applications/$Identity") {
            try { return Invoke-Graph -Uri $uri }
            catch { $lastError = Get-ErrorText $_ }
        }
        throw "No app registration found with Application (client) ID or Object ID '$Identity'. $lastError"
    }

    $found = @(Get-AllPages "v1.0/applications?$(Get-NameFilter $Identity)&`$select=id,appId,displayName")
    if ($found.Count -eq 0) {
        throw "No app registration found named '$Identity'."
    }
    if ($found.Count -gt 1) {
        $list = ($found | ForEach-Object { "  $($_.displayName)  appId $($_.appId)" }) -join "`n"
        throw "$($found.Count) app registrations are named '$Identity'. Use the Application (client) ID instead:`n$list"
    }
    return Invoke-Graph -Uri "v1.0/applications/$($found[0].id)"
}

function Get-NewIdentifierUris {
    # Application ID URIs for the copy. URIs that contain the source app ID get the new app
    # ID instead. Others must be unique in Entra, so they can't be copied; if none can be
    # carried over, the copy gets the standard api://<appId>.
    param([string]$NewAppId)
    $escaped = [regex]::Escape($srcAppId)
    $carried = @(foreach ($uri in $src.identifierUris) {
        if ($uri -match $escaped) { $uri -replace $escaped, $NewAppId }
    })
    if ($carried.Count -eq 0 -and (Get-Count $src.identifierUris) -gt 0) {
        $carried = @("api://$NewAppId")
    }
    $carried
}

# ========================
# CONNECT
# ========================
Write-Log "==========================================" Header
Write-Log "  Duplicate app registration" Header
Write-Log "  Source : $SourceApp" Header
Write-Log "  New    : $NewDisplayName" Header
if ($DryRun) { Write-Log "  DRY RUN - nothing will be created" Header }
Write-Log "==========================================" Header

$scopes = @(if ($DryRun) { 'Application.Read.All' } else { 'Application.ReadWrite.All' })
if ($CopyOwners -and -not $DryRun) { $scopes += 'Directory.Read.All' }
Write-Log "`nConnecting to Microsoft Graph ($($scopes -join ', '))..." Header
Connect-MgGraph -Scopes $scopes -NoWelcome
$ctx = Get-MgContext
Write-Log "  Connected as $($ctx.Account)" Success

# ========================
# READ THE SOURCE APP
# ========================
Write-Log "`nReading the source app registration..." Header
$src      = Find-SourceApp $SourceApp
$srcAppId = $src.appId
Write-Log "  Found $($src.displayName) (appId $srcAppId)" Success

# authenticationBehaviors is only returned when asked for by name.
$authBehaviors = $null
try {
    $selected = Invoke-Graph -Uri "v1.0/applications/$($src.id)?`$select=authenticationBehaviors"
    if ($selected.authenticationBehaviors) {
        $authBehaviors = [ordered]@{}
        foreach ($key in $selected.authenticationBehaviors.Keys) {
            $value = $selected.authenticationBehaviors[$key]
            if ("$key" -notlike '@odata.*' -and $null -ne $value) { $authBehaviors[$key] = $value }
        }
        if ($authBehaviors.Count -eq 0) { $authBehaviors = $null }
    }
}
catch {
    Write-Log "  Couldn't read authentication behaviors, so they won't be copied: $(Get-ErrorText $_)" Warn
    $NotCopied.Add("Authentication behaviors (couldn't be read)")
}

$srcSp = $null
try {
    $srcSp = Invoke-Graph -Uri "v1.0/servicePrincipals(appId='$srcAppId')?`$select=id,appRoleAssignmentRequired,tags,preferredSingleSignOnMode"
    Write-Log "  Found its enterprise application" Success
}
catch {
    Write-Log "  No enterprise application found for the source app ($(Get-ErrorText $_))" Detail
}

$srcOwners = @()
try {
    $srcOwners = @(Get-AllPages "v1.0/applications/$($src.id)/owners?`$select=id,displayName,userPrincipalName")
}
catch {
    if ($CopyOwners) { throw "Couldn't read the source app's owners: $(Get-ErrorText $_)" }
    Write-Log "  Couldn't read the source app's owners: $(Get-ErrorText $_)" Warn
}

$srcFederated = @()
try { $srcFederated = @(Get-AllPages "v1.0/applications/$($src.id)/federatedIdentityCredentials") }
catch { Write-Log "  Couldn't read federated credentials: $(Get-ErrorText $_)" Warn }

$srcExtensions = @()
try { $srcExtensions = @(Get-AllPages "v1.0/applications/$($src.id)/extensionProperties") }
catch { Write-Log "  Couldn't read extension properties: $(Get-ErrorText $_)" Warn }

# ========================
# CHECK THE NEW NAME
# ========================
Write-Log "`nChecking the new name is free..." Header
$sameName = @(Get-AllPages "v1.0/applications?$(Get-NameFilter $NewDisplayName)&`$select=id,appId,displayName")
if ($sameName.Count -gt 0) {
    throw "An app registration named '$NewDisplayName' already exists (appId $($sameName[0].appId)). Choose another name."
}
Write-Log "  '$NewDisplayName' is not in use" Success

# ========================
# BUILD THE NEW APP
# ========================
$body = [ordered]@{ displayName = $NewDisplayName }

foreach ($name in $CopyAsIs) {
    if ($null -ne $src[$name]) { $body[$name] = Copy-Clean $src[$name] }
}
if ($authBehaviors) { $body.authenticationBehaviors = $authBehaviors }
if ($src.nativeAuthenticationApisEnabled -and $src.nativeAuthenticationApisEnabled -ne 'none') {
    $body.nativeAuthenticationApisEnabled = $src.nativeAuthenticationApisEnabled
}

# Branding URLs. logoUrl is read-only; the logo itself is uploaded separately.
if ($src.info) {
    $info = [ordered]@{}
    foreach ($key in 'marketingUrl', 'privacyStatementUrl', 'supportUrl', 'termsOfServiceUrl') {
        if ($src.info[$key]) { $info[$key] = $src.info[$key] }
    }
    if ($info.Count -gt 0) { $body.info = $info }
}

# Authentication platforms. Only the settable parts are sent (Graph also returns
# read-only extras such as redirectUriSettings).
if ($src.web) {
    $web = [ordered]@{ redirectUris = @($src.web.redirectUris | Where-Object { $_ }) }
    if ($src.web.homePageUrl) { $web.homePageUrl = $src.web.homePageUrl }
    if ($src.web.logoutUrl)   { $web.logoutUrl   = $src.web.logoutUrl }
    if ($src.web.implicitGrantSettings) { $web.implicitGrantSettings = Copy-Clean $src.web.implicitGrantSettings }
    $body.web = $web
}
if ($src.spa) {
    $body.spa = [ordered]@{ redirectUris = @($src.spa.redirectUris | Where-Object { $_ }) }
}
if ($src.publicClient) {
    $body.publicClient = [ordered]@{ redirectUris = @($src.publicClient.redirectUris | Where-Object { $_ }) }
}

# App roles and scopes get new IDs. Separate maps, because a role and a scope may share an ID.
$roleIdMap = @{}
$body.appRoles = @(foreach ($role in $src.appRoles) {
    $newId = [guid]::NewGuid().Guid
    $roleIdMap[$role.id] = $newId
    [ordered]@{
        id                 = $newId
        allowedMemberTypes = @($role.allowedMemberTypes)
        description        = $role.description
        displayName        = $role.displayName
        isEnabled          = $role.isEnabled
        value              = $role.value
    }
})

$scopeIdMap   = @{}
$scopeValues  = @{}
$preAuthorized = @()
if ($src.api) {
    $newScopes = @(foreach ($scope in $src.api.oauth2PermissionScopes) {
        $newId = [guid]::NewGuid().Guid
        $scopeIdMap[$scope.id]  = $newId
        $scopeValues[$scope.id] = $scope.value
        [ordered]@{
            id                      = $newId
            adminConsentDescription = $scope.adminConsentDescription
            adminConsentDisplayName = $scope.adminConsentDisplayName
            isEnabled               = $scope.isEnabled
            type                    = $scope.type
            userConsentDescription  = $scope.userConsentDescription
            userConsentDisplayName  = $scope.userConsentDisplayName
            value                   = $scope.value
        }
    })

    # preAuthorizedApplications is left out on purpose: it lets client apps skip consent.
    $api = [ordered]@{ oauth2PermissionScopes = $newScopes }
    if ($null -ne $src.api.requestedAccessTokenVersion) { $api.requestedAccessTokenVersion = $src.api.requestedAccessTokenVersion }
    if ($null -ne $src.api.acceptMappedClaims)          { $api.acceptMappedClaims          = $src.api.acceptMappedClaims }
    $api.knownClientApplications = @($src.api.knownClientApplications | Where-Object { $_ -and $_ -ne $srcAppId })
    $body.api = $api

    if (@($src.api.knownClientApplications) -contains $srcAppId) {
        $NotCopied.Add('Known client applications: the entry pointing at the source app itself')
    }
    $preAuthorized = @($src.api.preAuthorizedApplications | Where-Object { $_ })
}

# API permissions. Ones the source app requests on its own API are added after the new
# app exists, pointing at the new app and its new role and scope IDs.
$rraOther   = [System.Collections.Generic.List[object]]::new()
$selfAccess = @()
foreach ($rra in $src.requiredResourceAccess) {
    if ($rra.resourceAppId -ne $srcAppId) {
        $rraOther.Add((Copy-Clean $rra))
        continue
    }
    $selfAccess += @(foreach ($access in $rra.resourceAccess) {
        $map = if ($access.type -eq 'Role') { $roleIdMap } else { $scopeIdMap }
        if ($map.ContainsKey($access.id)) {
            [ordered]@{ id = $map[$access.id]; type = $access.type }
        }
        else {
            $NotCopied.Add("API permission $($access.id) on its own API (no longer defined on the source app)")
        }
    })
}
$body.requiredResourceAccess = @($rraOther)

$permissionCount = 0
foreach ($rra in $rraOther) { $permissionCount += Get-Count $rra.resourceAccess }
$permissionCount += $selfAccess.Count
$apiCount = $rraOther.Count + [int]($selfAccess.Count -gt 0)

# Enterprise application. "Assignment required?" and "Visible to users?" (the HideApp
# tag) come from the source app's enterprise application when it has one.
$spBody = $null
if (-not $SkipServicePrincipal) {
    $spTags = @($IntegratedAppTag)
    if ($srcSp -and @($srcSp.tags) -contains 'HideApp') { $spTags += 'HideApp' }
    $spBody = [ordered]@{ tags = $spTags }
    if ($srcSp) { $spBody.appRoleAssignmentRequired = [bool]$srcSp.appRoleAssignmentRequired }
}

# What won't be copied, for the report.
$secretCount = Get-Count $src.passwordCredentials
$certCount   = Get-Count $src.keyCredentials
if ($secretCount -gt 0)          { $NotCopied.Add("$secretCount client secret(s) (on purpose)") }
if ($certCount -gt 0)            { $NotCopied.Add("$certCount certificate(s) (on purpose)") }
if ($srcFederated.Count -gt 0)   { $NotCopied.Add("$($srcFederated.Count) federated credential(s) (on purpose): $(($srcFederated | ForEach-Object { $_.name }) -join ', ')") }
foreach ($pre in $preAuthorized) {
    $values = @($pre.delegatedPermissionIds | ForEach-Object { if ($scopeValues[$_]) { $scopeValues[$_] } else { $_ } })
    $NotCopied.Add("Authorized client application $($pre.appId) for $($values -join ', ') (on purpose: it skips consent)")
}
foreach ($uri in $src.identifierUris) {
    if ($uri -notmatch [regex]::Escape($srcAppId)) { $NotCopied.Add("Application ID URI $uri (must be unique)") }
}
if ($src.info -and $src.info.logoUrl)  { $NotCopied.Add('Logo (upload it under Branding & properties)') }
if ($srcExtensions.Count -gt 0)        { $NotCopied.Add("Extension properties ($($srcExtensions.Count)): $(($srcExtensions | ForEach-Object { $_.name }) -join ', ')") }
if ((Get-Count $src.addIns) -gt 0)     { $NotCopied.Add("$(Get-Count $src.addIns) add-in(s)") }
if ($src.tokenEncryptionKeyId)         { $NotCopied.Add('Token encryption key (it points at a certificate)') }
if ($srcOwners.Count -gt 0 -and -not $CopyOwners) {
    $NotCopied.Add("$($srcOwners.Count) owner(s) (use -CopyOwners to copy them)")
}

# ========================
# SHOW THE PLAN
# ========================
$plannedUris = @(Get-NewIdentifierUris '<new appId>')
$assignment  = if ($srcSp) { if ($spBody.appRoleAssignmentRequired) { 'Yes' } else { 'No' } } else { 'No (default)' }
$visible     = if ($spBody -and $spBody.tags -contains 'HideApp') { 'No' } else { 'Yes' }

Write-Log "`nWill copy:" Header
Write-Log ("  {0,-26} {1}" -f 'Supported account types', $body.signInAudience)
Write-Log ("  {0,-26} Web {1}, SPA {2}, mobile/desktop {3}" -f 'Redirect URIs', (Get-Count $body.web.redirectUris), (Get-Count $body.spa.redirectUris), (Get-Count $body.publicClient.redirectUris))
Write-Log ("  {0,-26} {1} across {2} API(s) - configured, NOT granted" -f 'API permissions', $permissionCount, $apiCount)
Write-Log ("  {0,-26} {1}" -f 'App roles', $body.appRoles.Count)
Write-Log ("  {0,-26} {1}" -f 'Exposed scopes', (Get-Count $body.api.oauth2PermissionScopes))
Write-Log ("  {0,-26} {1}" -f 'Application ID URI', $(if ($plannedUris.Count) { $plannedUris -join ', ' } else { '(none)' }))
Write-Log ("  {0,-26} {1}" -f 'Optional claims', $(if ($body.optionalClaims) { 'Yes' } else { 'None' }))
Write-Log ("  {0,-26} {1}" -f 'Groups claim', $(if ($body.groupMembershipClaims) { $body.groupMembershipClaims } else { 'None' }))
if ($SkipServicePrincipal) {
    Write-Log ("  {0,-26} {1}" -f 'Enterprise application', 'Not created (-SkipServicePrincipal)')
}
else {
    Write-Log ("  {0,-26} Created. Assignment required: {1}. Visible to users: {2}" -f 'Enterprise application', $assignment, $visible)
}
if ($CopyOwners) {
    Write-Log ("  {0,-26} {1}" -f 'Owners', $(if ($srcOwners.Count) { ($srcOwners | ForEach-Object { Get-OwnerLabel $_ }) -join '; ' } else { 'Source app has none' }))
}

Write-Log "`nNot copied:" Header
if ($NotCopied.Count -eq 0) { Write-Log "  Nothing" Detail }
foreach ($line in $NotCopied) { Write-Log "  - $line" Warn }

if ($srcSp -and $srcSp.preferredSingleSignOnMode -in 'saml', 'password') {
    $ssoName = if ($srcSp.preferredSingleSignOnMode -eq 'saml') { 'SAML' } else { 'password-based' }
    Write-Log "`nThe source app uses $ssoName single sign-on. That's set up on the enterprise application and isn't copied." Warn
}

$result = [pscustomobject]@{
    Status                = 'DryRun'
    SourceDisplayName     = $src.displayName
    SourceAppId           = $srcAppId
    NewDisplayName        = $NewDisplayName
    NewAppId              = $null
    NewObjectId           = $null
    NewServicePrincipalId = $null
    NotCopied             = ($NotCopied -join '; ')
    Problems              = $null
}

if ($DryRun) {
    Write-Log "`nRequest body that would be sent to create the app:" Header
    Write-Log (ConvertTo-Json -InputObject $body -Depth 20) Detail
    if ($selfAccess.Count -gt 0) {
        Write-Log "`nThen, once the app exists, it would also request $($selfAccess.Count) permission(s) on its own API." Detail
    }
    Write-Log "`nDry run finished. Nothing was created." Header
    return $result
}

# ========================
# CREATE
# ========================
Write-Log "`nCreating the app registration..." Header
$new = Invoke-Graph -Method POST -Uri 'v1.0/applications' -Body $body
$result.Status      = 'Created'
$result.NewAppId    = $new.appId
$result.NewObjectId = $new.id
Write-Log "  Created $($new.displayName)" Success
Write-Log "    Application (client) ID : $($new.appId)" Detail
Write-Log "    Object ID               : $($new.id)" Detail

$newUris = @(Get-NewIdentifierUris $new.appId)
if ($newUris.Count -gt 0) {
    try {
        $null = Invoke-Graph -Method PATCH -Uri "v1.0/applications/$($new.id)" -Body @{ identifierUris = $newUris } -Retry
        Write-Log "  Set Application ID URI: $($newUris -join ', ')" Success
    }
    catch { Add-Problem "Application ID URI $($newUris -join ', '): $(Get-ErrorText $_)" }
}

if ($selfAccess.Count -gt 0) {
    $fullAccess = @($rraOther) + @([ordered]@{ resourceAppId = $new.appId; resourceAccess = $selfAccess })
    try {
        $null = Invoke-Graph -Method PATCH -Uri "v1.0/applications/$($new.id)" -Body @{ requiredResourceAccess = $fullAccess } -Retry
        Write-Log "  Added $($selfAccess.Count) permission(s) on its own API" Success
    }
    catch { Add-Problem "API permissions on its own API: $(Get-ErrorText $_)" }
}

if ($spBody) {
    Write-Log "`nCreating the enterprise application..." Header
    $spBody.appId = $new.appId
    try {
        $newSp = Invoke-Graph -Method POST -Uri 'v1.0/servicePrincipals' -Body $spBody -Retry
        $result.NewServicePrincipalId = $newSp.id
        Write-Log "  Created (Object ID $($newSp.id))" Success
    }
    catch { Add-Problem "Enterprise application: $(Get-ErrorText $_)" }
}

if ($CopyOwners -and $srcOwners.Count -gt 0) {
    Write-Log "`nCopying owners..." Header
    try {
        # Some roles (such as Application Developer) are made owner of apps they create.
        $haveIds = @(Get-AllPages "v1.0/applications/$($new.id)/owners?`$select=id" -Retry | ForEach-Object { $_['id'] })
        foreach ($owner in $srcOwners) {
            if ($haveIds -contains $owner.id) {
                Write-Log "  Already an owner: $(Get-OwnerLabel $owner)" Detail
                continue
            }
            try {
                $ref = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($owner.id)" }
                $null = Invoke-Graph -Method POST -Uri "v1.0/applications/$($new.id)/owners/`$ref" -Body $ref -Retry
                Write-Log "  Added: $(Get-OwnerLabel $owner)" Success
            }
            catch { Add-Problem "Owner $(Get-OwnerLabel $owner): $(Get-ErrorText $_)" }
        }
    }
    catch { Add-Problem "Owners: $(Get-ErrorText $_)" }
}

# ========================
# CHECK THE RESULT
# ========================
Write-Log "`nChecking the new app..." Header
try {
    $check = Invoke-Graph -Uri "v1.0/applications/$($new.id)" -Retry

    $credentials = (Get-Count $check.passwordCredentials) + (Get-Count $check.keyCredentials)
    if ($credentials -eq 0) { Write-Log "  No client secrets or certificates" Success }
    else { Add-Problem "The new app has $credentials secret(s) or certificate(s). It should have none." }

    $actualPermissions = 0
    foreach ($rra in $check.requiredResourceAccess) { $actualPermissions += Get-Count $rra.resourceAccess }

    $comparisons = [ordered]@{
        'Web redirect URIs'            = @((Get-Count $body.web.redirectUris), (Get-Count $check.web.redirectUris))
        'SPA redirect URIs'            = @((Get-Count $body.spa.redirectUris), (Get-Count $check.spa.redirectUris))
        'Mobile/desktop redirect URIs' = @((Get-Count $body.publicClient.redirectUris), (Get-Count $check.publicClient.redirectUris))
        'API permissions'              = @($permissionCount, $actualPermissions)
        'App roles'                    = @($body.appRoles.Count, (Get-Count $check.appRoles))
        'Exposed scopes'               = @((Get-Count $body.api.oauth2PermissionScopes), (Get-Count $check.api.oauth2PermissionScopes))
        'Application ID URIs'          = @($newUris.Count, (Get-Count $check.identifierUris))
    }
    foreach ($item in $comparisons.GetEnumerator()) {
        $expected, $actual = $item.Value
        if ($expected -eq $actual) { Write-Log ("  {0,-30} {1}" -f $item.Key, $actual) Success }
        else { Add-Problem "$($item.Key): expected $expected, found $actual" }
    }

    $newOwners = @(Get-AllPages "v1.0/applications/$($new.id)/owners?`$select=id,displayName,userPrincipalName" -Retry)
    $ownerText = if ($newOwners.Count) { ($newOwners | ForEach-Object { Get-OwnerLabel $_ }) -join '; ' } else { 'None' }
    Write-Log ("  {0,-30} {1}" -f 'Owners', $ownerText) Detail
}
catch { Add-Problem "Checking the new app: $(Get-ErrorText $_)" }

# ========================
# SUMMARY
# ========================
Write-Log "`n==========================================" Header
Write-Log "  Done" Header
Write-Log "==========================================" Header
Write-Log "  New app                 : $($new.displayName)"
Write-Log "  Application (client) ID : $($new.appId)"
Write-Log "  Object ID               : $($new.id)"
if ($result.NewServicePrincipalId) { Write-Log "  Enterprise app ID       : $($result.NewServicePrincipalId)" }

Write-Log "`nNext steps:" Header
Write-Log "  - API permissions are configured but NOT granted. If the app needs consent, grant it in"
Write-Log "    App registrations > $($new.displayName) > API permissions."
Write-Log "  - Add a client secret, certificate or federated credential if the app needs one."
if ($NotCopied.Count -gt 0) { Write-Log "  - Review the 'Not copied' list above." }

if ($Problems.Count -gt 0) {
    $result.Status   = 'CreatedWithProblems'
    $result.Problems = $Problems -join '; '
    Write-Log "`nThe app was created, but these steps failed. Fix them by hand or delete the app:" Error
    foreach ($problem in $Problems) { Write-Log "  - $problem" Error }
}

$result

if ($Problems.Count -gt 0) {
    throw "$($Problems.Count) step(s) failed after the app was created. See the list above."
}
