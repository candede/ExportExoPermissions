# ==============================================================================
# Script  : Export-ExoPermissions-HybridReady.ps1
# Version : 2.8
# Purpose : Export ALL mailbox-level and folder-level Exchange Online permissions
#           into a restorable JSON schema for cross-tenant migration.
#           Supports Cloud-Only and Hybrid (AD-Synced) tenants.
#
# Captures:
#   - Mailbox X500 aliases (all historical X500 proxy addresses)
#   - Full Access (MailboxPermission)
#   - Send As    (RecipientPermission)
#   - Send on Behalf (GrantSendOnBehalfTo)
#   - ALL Folder-Level Permissions (dynamically enumerated)
#   - Unresolvable trustee logging (critical for hybrid envs)
#
# Output  : JSON file per run, path auto-resolved cross-platform.
# ==============================================================================

#Requires -Modules ExchangeOnlineManagement

[CmdletBinding()]
param (
    # Optional: path override for the output JSON. Defaults to script dir.
    [string]$OutputPath = "",

    # Use tenant/app credentials from the .env file instead of interactive sign-in.
    [switch]$UseEnvCredentials,

    # Optional path to the .env file. Defaults to .env in the script directory.
    [string]$EnvFilePath = "",

    # Optional: restrict to specific mailbox types (default: all user mailboxes)
    [string[]]$RecipientTypeDetails = @(
        "UserMailbox",
        "SharedMailbox",
        "RoomMailbox",
        "EquipmentMailbox"
    ),

    # Skip folder permission collection entirely (faster, useful for mailbox-only audits)
    [switch]$SkipFolderPermissions,

    # Optional explicit folder selection. When provided, only matching folders are scanned.
    # When left empty, all folders are scanned.
    # Supports folder type names (Calendar, Inbox, Root), folder names, or full folder paths
    # such as /Calendar or /Inbox/Subfolder.
    [string[]]$IncludedFolders = @(),

    # How many ms to sleep between mailboxes to avoid throttling
    [int]$ThrottleDelayMs = 150,

    # Max retry attempts for throttled/failed EXO calls
    [int]$MaxRetries = 3,

    # How often to incrementally save the JSON output to disk (in number of mailboxes)
    [int]$SaveInterval = 10,

    # Automatically find the most recent output file in the script directory and resume it.
    [switch]$Resume
)

# ==============================================================================
# SECTION 0 — Setup
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"
$ScriptVersion = "2.8"

# Cross-platform output path
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

if ([string]::IsNullOrWhiteSpace($EnvFilePath)) {
    $EnvFilePath = Join-Path $scriptDir ".env"
}

if ($Resume -and [string]::IsNullOrWhiteSpace($OutputPath)) {
    $latestFile = Get-ChildItem -Path $scriptDir -Filter "ExoPermissions_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latestFile) {
        $OutputPath = $latestFile.FullName
        Write-Host "Resume switch active. Found latest save: $OutputPath" -ForegroundColor Yellow
    }
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $scriptDir "ExoPermissions_$(Get-Date -Format 'yyyyMMdd_HHmm').json"
}

Write-Host "`n=== EXO Permission Exporter v$ScriptVersion ===" -ForegroundColor Cyan
Write-Host "Output : $OutputPath" -ForegroundColor DarkGray
Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" -ForegroundColor DarkGray

# Global identity cache: raw trustee string -> resolved SMTP or $null
$global:IdentityCache = [System.Collections.Generic.Dictionary[string, object]]::new(
    [System.StringComparer]::OrdinalIgnoreCase)
$global:ResolveErrors = 0
$global:FolderErrors = 0
$global:sw = [System.Diagnostics.Stopwatch]::StartNew()

function Get-EnvVariables {
    param([string]$Path)

    if (-not (Test-Path -Path $Path)) {
        throw ".env file not found at: $Path"
    }

    $envVars = @{}
    Get-Content -Path $Path | ForEach-Object {
        if ($_ -match '^\s*([^#][^=]+)\s*=\s*(.+)\s*$') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $value = $value -replace '^["'']|["'']$', ''
            $envVars[$key] = $value
        }
    }

    return $envVars
}

function Connect-ExchangeOnlineSession {
    param(
        [switch]$UseEnvCredentials,
        [string]$EnvFilePath
    )

    if (-not $UseEnvCredentials) {
        Write-Host "Connecting to Exchange Online using interactive sign-in..." -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowBanner:$false
        return
    }

    Write-Host "Loading Exchange Online app credentials from .env..." -ForegroundColor Cyan
    $envVars = Get-EnvVariables -Path $EnvFilePath

    $tenantId = $envVars["TENANT_ID"]
    $clientId = $envVars["CLIENT_ID"]
    $clientSecret = $envVars["CLIENT_SECRET"]
    $organizationName = $envVars["ORGANIZATION"]

    if (-not $tenantId -or -not $clientId -or -not $clientSecret) {
        throw "Missing required environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET"
    }

    if (-not $organizationName) {
        Write-Warning "ORGANIZATION not set in .env file. Attempting connection without it."
    }

    $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $clientId
        client_secret = $clientSecret
        scope         = "https://outlook.office365.com/.default"
        grant_type    = "client_credentials"
    }

    Write-Host "Acquiring Exchange Online access token..." -ForegroundColor DarkGray
    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body -ContentType "application/x-www-form-urlencoded"
    $accessToken = $tokenResponse.access_token

    $connectParams = @{
        AccessToken = $accessToken
        ShowBanner  = $false
    }

    if ($organizationName) {
        $connectParams.Organization = $organizationName
    }

    Write-Host "Connecting to Exchange Online using app-only authentication..." -ForegroundColor Cyan
    Connect-ExchangeOnline @connectParams
}

# ==============================================================================
# SECTION 1 — Connect
# ==============================================================================

try {
    Connect-ExchangeOnlineSession -UseEnvCredentials:$UseEnvCredentials -EnvFilePath $EnvFilePath
}
catch {
    Write-Error "Failed to connect to Exchange Online: $_"
    exit 1
}

# Capture tenant context for schema metadata
$tenantInfo = $null
try {
    $tenantInfo = Get-OrganizationConfig -ErrorAction Stop
}
catch {
    Write-Warning "Could not retrieve OrganizationConfig. Tenant metadata will be partial."
}

function Get-SafeObjectPropertyValue {
    param(
        [object]$InputObject,
        [string[]]$PropertyNames
    )

    if ($null -eq $InputObject) {
        return $null
    }

    foreach ($propertyName in $PropertyNames) {
        $property = $InputObject.PSObject.Properties[$propertyName]
        if ($property -and $null -ne $property.Value -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            return [string]$property.Value
        }
    }

    return $null
}

function Get-OptionalObjectPropertyValue {
    param(
        [object]$InputObject,
        [string]$PropertyName,
        [object]$DefaultValue = $null
    )

    if ($null -eq $InputObject -or [string]::IsNullOrWhiteSpace($PropertyName)) {
        return $DefaultValue
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($property) {
        return $property.Value
    }

    return $DefaultValue
}

function Normalize-X500Address {
    param([string]$Address)

    if ([string]::IsNullOrWhiteSpace($Address)) {
        return $null
    }

    $trimmed = $Address.Trim()
    if ($trimmed -match '^(?i)x500:(.+)$') {
        return "X500:$($Matches[1])"
    }

    return "X500:$trimmed"
}

function Get-X500ProxyAddresses {
    param([object]$Mailbox)

    $x500Addresses = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)

    $emailAddressesProperty = $Mailbox.PSObject.Properties['EmailAddresses']
    if ($emailAddressesProperty -and $null -ne $emailAddressesProperty.Value) {
        foreach ($emailAddress in @($emailAddressesProperty.Value)) {
            $addressString = [string]$emailAddress
            if ($addressString -match '^(?i)x500:') {
                $normalizedAddress = Normalize-X500Address -Address $addressString
                if (-not [string]::IsNullOrWhiteSpace($normalizedAddress)) {
                    $x500Addresses.Add($normalizedAddress) | Out-Null
                }
            }
        }
    }

    $legacyExchangeDn = Get-SafeObjectPropertyValue -InputObject $Mailbox -PropertyNames @('LegacyExchangeDN')
    if (-not [string]::IsNullOrWhiteSpace($legacyExchangeDn)) {
        $normalizedLegacyAddress = Normalize-X500Address -Address $legacyExchangeDn
        if (-not [string]::IsNullOrWhiteSpace($normalizedLegacyAddress)) {
            $x500Addresses.Add($normalizedLegacyAddress) | Out-Null
        }
    }

    return @($x500Addresses | Sort-Object)
}

function Normalize-FolderSelector {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $normalized = $Value.Trim().Replace('\\', '/').TrimEnd('/').ToLowerInvariant()

    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return '/'
    }

    return $normalized
}

function Test-FolderSelection {
    param(
        [object]$Folder,
        [string[]]$Selectors
    )

    if ($null -eq $Selectors -or $Selectors.Count -eq 0) {
        return $true
    }

    $folderPath = [string]$Folder.FolderPath
    $folderType = [string]$Folder.FolderType

    $normalizedPath = Normalize-FolderSelector -Value $folderPath
    $normalizedPathNoSlash = if ($normalizedPath -eq '/') { 'root' } else { $normalizedPath.TrimStart('/') }
    $leafName = if ($normalizedPath -eq '/') {
        'root'
    }
    else {
        ($normalizedPathNoSlash -split '/')[(-1)]
    }

    $candidates = @(
        $normalizedPath,
        $normalizedPathNoSlash,
        $leafName,
        (Normalize-FolderSelector -Value $folderType)
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    foreach ($selector in $Selectors) {
        $normalizedSelector = Normalize-FolderSelector -Value $selector
        if ([string]::IsNullOrWhiteSpace($normalizedSelector)) {
            continue
        }

        if ($normalizedSelector -in @('root', '/root')) {
            $normalizedSelector = '/'
        }

        if ($normalizedSelector -in $candidates) {
            return $true
        }

        if ($normalizedSelector.StartsWith('/') -and $normalizedPath -eq $normalizedSelector) {
            return $true
        }
    }

    return $false
}

$usingExplicitFolderSelection = $IncludedFolders.Count -gt 0

$sourceTenantDomain = Get-SafeObjectPropertyValue -InputObject $tenantInfo -PropertyNames @(
    "DefaultDomainName",
    "Name"
)

$sourceTenantId = Get-SafeObjectPropertyValue -InputObject $tenantInfo -PropertyNames @(
    "ExchangeObjectId",
    "Guid",
    "OrganizationId"
)

if ([string]::IsNullOrWhiteSpace($sourceTenantDomain)) {
    try {
        $defaultAcceptedDomain = Get-AcceptedDomain -ErrorAction Stop |
        Where-Object { $_.Default -eq $true } |
        Select-Object -First 1

        $sourceTenantDomain = Get-SafeObjectPropertyValue -InputObject $defaultAcceptedDomain -PropertyNames @(
            "DomainName",
            "Name"
        )
    }
    catch {
        Write-Warning "Could not retrieve default accepted domain. Tenant domain metadata will be 'Unknown'."
    }
}

if ([string]::IsNullOrWhiteSpace($sourceTenantDomain)) {
    $sourceTenantDomain = "Unknown"
}

if ([string]::IsNullOrWhiteSpace($sourceTenantId)) {
    $sourceTenantId = "Unknown"
}

# ==============================================================================
# SECTION 2 — Helper: EXO call with retry/backoff
# ==============================================================================

function Invoke-EXOWithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [string]$Description = "EXO call",
        [int]$MaxAttempts = $MaxRetries
    )

    $attempt = 0
    do {
        $attempt++
        try {
            return (& $ScriptBlock)
        }
        catch {
            $msg = $_.Exception.Message
            if ($msg -match "throttl|429|ServiceUnavailable|TooManyRequests" -and $attempt -lt $MaxAttempts) {
                $wait = [math]::Pow(2, $attempt) * 1000  # exponential backoff: 2s, 4s, 8s
                Write-Warning "Throttled on '$Description' — retrying in $($wait/1000)s (attempt $attempt/$MaxAttempts)"
                Start-Sleep -Milliseconds $wait
            }
            else {
                throw
            }
        }
    } while ($attempt -lt $MaxAttempts)
}

# ==============================================================================
# SECTION 3 — Helper: Resolve trustee identity to SMTP + metadata
# ==============================================================================

function Get-ResolvedIdentity {
    param([string]$RawValue, [string]$PermissionType, [string]$FolderPath = "")

    $result = [ordered]@{
        DelegateUPN                       = $null
        DelegateDisplayName               = $null
        DelegateExternalDirectoryObjectId = $null
        DelegateRecipientType             = $null
        PermissionStatus                  = "Unknown"
        RawTrusteeValue                   = $RawValue
        UnresolvableReason                = $null
    }

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        $result.PermissionStatus = "SkippedEmpty"
        return $result
    }

    # System/special principals — skip silently
    if ($RawValue -match "^NT AUTHORITY\\|^BUILTIN\\") {
        $result.PermissionStatus = "SkippedSystem"
        return $result
    }

    # Orphaned on-premises SIDs (object deleted from AD but ACE remains)
    if ($RawValue -match "^S-1-5-") {
        $result.PermissionStatus = "Unresolvable"
        $result.UnresolvableReason = "OrphanedSID"
        return $result
    }

    # Well-known folder permission principals — keep as-is
    if ($RawValue -match "^Default$|^Anonymous$") {
        $result.DelegateUPN = $RawValue
        $result.PermissionStatus = "WellKnown"
        return $result
    }

    # Cache hit
    if ($global:IdentityCache.ContainsKey($RawValue)) {
        $cached = $global:IdentityCache[$RawValue]
        if ($null -eq $cached) {
            $result.PermissionStatus = "Unresolvable"
            $result.UnresolvableReason = "CachedNotFound"
        }
        else {
            $result.DelegateUPN = $cached.PrimarySmtpAddress
            $result.DelegateDisplayName = $cached.DisplayName
            $result.DelegateExternalDirectoryObjectId = $cached.ExternalDirectoryObjectId
            $result.DelegateRecipientType = $cached.RecipientTypeDetails
            $result.PermissionStatus = "Resolved"
        }
        return $result
    }

    # Live resolution via EXO
    try {
        $recipient = Invoke-EXOWithRetry -Description "Resolve $RawValue" -ScriptBlock {
            Get-EXORecipient -Identity $RawValue -Properties DisplayName, ExternalDirectoryObjectId, RecipientTypeDetails -ErrorAction Stop
        }

        $obj = [ordered]@{
            PrimarySmtpAddress        = $recipient.PrimarySmtpAddress
            DisplayName               = $recipient.DisplayName
            ExternalDirectoryObjectId = $recipient.ExternalDirectoryObjectId
            RecipientTypeDetails      = $recipient.RecipientTypeDetails
        }

        $global:IdentityCache[$RawValue] = $obj

        $result.DelegateUPN = $obj.PrimarySmtpAddress
        $result.DelegateDisplayName = $obj.DisplayName
        $result.DelegateExternalDirectoryObjectId = $obj.ExternalDirectoryObjectId
        $result.DelegateRecipientType = $obj.RecipientTypeDetails
        $result.PermissionStatus = "Resolved"
    }
    catch {
        $global:IdentityCache[$RawValue] = $null
        $global:ResolveErrors++
        $result.PermissionStatus = "Unresolvable"
        $result.UnresolvableReason = "RecipientNotFound"
        Write-Verbose "Could not resolve '$RawValue': $_"
    }

    return $result
}

# ==============================================================================
# SECTION 4 — Helper: Build delegate record (no PII duplication)
# ==============================================================================

function New-DelegateRecord {
    param(
        [hashtable]$Resolved,
        [hashtable]$ExtraFields = @{}
    )

    $rec = [ordered]@{
        DelegateUPN                       = $Resolved.DelegateUPN
        DelegateDisplayName               = $Resolved.DelegateDisplayName
        DelegateExternalDirectoryObjectId = $Resolved.DelegateExternalDirectoryObjectId
        DelegateRecipientType             = $Resolved.DelegateRecipientType
        PermissionStatus                  = $Resolved.PermissionStatus
    }

    foreach ($key in $ExtraFields.Keys) {
        $rec[$key] = $ExtraFields[$key]
    }

    return $rec
}

# ==============================================================================
# SECTION 5 — Discover mailboxes
# ==============================================================================

Write-Host "Discovering mailboxes..." -ForegroundColor Cyan

$mailboxes = Invoke-EXOWithRetry -Description "Get-EXOMailbox" -ScriptBlock {
    Get-EXOMailbox -ResultSize Unlimited `
        -RecipientTypeDetails $RecipientTypeDetails `
        -Properties UserPrincipalName, PrimarySmtpAddress, DisplayName,
    RecipientTypeDetails, ExternalDirectoryObjectId,
    LegacyExchangeDN, EmailAddresses, GrantSendOnBehalfTo -ErrorAction Stop
}

$total = $mailboxes.Count
$count = 0
$exportList = [System.Collections.Generic.List[object]]::new()
$processedUPNs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

if (Test-Path $OutputPath) {
    Write-Host "`nFound existing output file '$OutputPath'." -ForegroundColor Yellow
    Write-Host "Attempting to recover and resume..." -ForegroundColor DarkGray
    try {
        # Need to parse json with appropriate depth
        $existingOutput = Get-Content $OutputPath -Raw | ConvertFrom-Json -Depth 20
        if ($existingOutput.Mailboxes) {
            foreach ($mbx in $existingOutput.Mailboxes) {
                # Add to exportList
                $exportList.Add($mbx)
                if (-not [string]::IsNullOrWhiteSpace($mbx.SourceUser.UserPrincipalName)) {
                    $processedUPNs.Add($mbx.SourceUser.UserPrincipalName) | Out-Null
                }
            }
            Write-Host "Successfully loaded $($processedUPNs.Count) already-processed mailboxes." -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "Could not parse existing output file. Starting fresh."
    }
}

Write-Host "Found $total mailboxes." -ForegroundColor Green

# ==============================================================================
# SECTION 5.5 — Pre-populate identity cache (bulk fetch all recipients)
# ==============================================================================
# Instead of resolving each trustee one-by-one (hundreds of API calls per mailbox),
# we fetch ALL recipients in one bulk call and cache them by multiple keys.
# This typically reduces runtime from hours to minutes for large tenants.

Write-Host "Pre-loading recipient directory for fast trustee resolution..." -ForegroundColor Cyan
$preCacheTimer = [System.Diagnostics.Stopwatch]::StartNew()

try {
    $allRecipients = Invoke-EXOWithRetry -Description "Get-EXORecipient (bulk pre-cache)" -ScriptBlock {
        Get-EXORecipient -ResultSize Unlimited `
            -Properties DisplayName, ExternalDirectoryObjectId, RecipientTypeDetails, PrimarySmtpAddress `
            -ErrorAction Stop
    }

    $recipientCount = 0
    foreach ($r in $allRecipients) {
        $obj = [ordered]@{
            PrimarySmtpAddress        = $r.PrimarySmtpAddress
            DisplayName               = $r.DisplayName
            ExternalDirectoryObjectId = $r.ExternalDirectoryObjectId
            RecipientTypeDetails      = $r.RecipientTypeDetails
        }

        # Cache by PrimarySmtpAddress (most common lookup key for mailbox-level perms)
        if (-not [string]::IsNullOrWhiteSpace($r.PrimarySmtpAddress)) {
            $global:IdentityCache[$r.PrimarySmtpAddress] = $obj
        }

        # Cache by DisplayName (common lookup key for folder-level perms)
        if (-not [string]::IsNullOrWhiteSpace($r.DisplayName) -and
            -not $global:IdentityCache.ContainsKey($r.DisplayName)) {
            $global:IdentityCache[$r.DisplayName] = $obj
        }

        # Cache by Alias (fallback lookup key)
        if (-not [string]::IsNullOrWhiteSpace($r.Alias) -and
            -not $global:IdentityCache.ContainsKey($r.Alias)) {
            $global:IdentityCache[$r.Alias] = $obj
        }

        # Cache by Name property (used in some permission entries)
        if (-not [string]::IsNullOrWhiteSpace($r.Name) -and
            -not $global:IdentityCache.ContainsKey($r.Name)) {
            $global:IdentityCache[$r.Name] = $obj
        }

        $recipientCount++
    }

    $preCacheTimer.Stop()
    Write-Host "Cached $($global:IdentityCache.Count) identity keys from $recipientCount recipients in $($preCacheTimer.Elapsed.ToString('mm\:ss'))." -ForegroundColor Green
}
catch {
    $preCacheTimer.Stop()
    Write-Warning "Could not pre-load recipients ($($preCacheTimer.Elapsed.ToString('mm\:ss'))). Falling back to per-trustee resolution (slower). Error: $_"
}

Write-Host "Starting extraction...`n" -ForegroundColor Cyan

# ==============================================================================
# SECTION 6 — Main loop
# ==============================================================================

$processedThisRun = 0

foreach ($mbx in $mailboxes) {
    $count++
    $pct = [int](($count / $total) * 100)
    $x500Addresses = @(Get-X500ProxyAddresses -Mailbox $mbx)
    
    # Check if mailbox was already processed in a previous run
    if ($processedUPNs.Contains($mbx.UserPrincipalName)) {
        Write-Progress -Id 0 -Activity "Extracting EXO Permissions" `
            -Status "$($mbx.UserPrincipalName) ($count of $total) | Skipped (Already Processed)" `
            -PercentComplete $pct
        continue
    }

    $processedThisRun++
    $mbxSw = [System.Diagnostics.Stopwatch]::StartNew()

    # Calculate ETA based on average processing speed so far
    $elapsed = $global:sw.Elapsed
    $etaStr = if ($processedThisRun -gt 1) {
        $avgMs = $elapsed.TotalMilliseconds / ($processedThisRun - 1)
        $remainingMs = $avgMs * ($total - $count + 1)
        $eta = [TimeSpan]::FromMilliseconds($remainingMs)
        "{0:hh\:mm\:ss}" -f $eta
    }
    else { "calculating..." }

    Write-Progress -Id 0 -Activity "Extracting EXO Permissions" `
        -Status "$($mbx.UserPrincipalName) ($count of $total) | ETA: $etaStr" `
        -PercentComplete $pct

    $mbxErrors = [System.Collections.Generic.List[string]]::new()
    $unresolvable = [System.Collections.Generic.List[object]]::new()

    $mbxData = [ordered]@{
        SourceUser              = [ordered]@{
            UserPrincipalName         = $mbx.UserPrincipalName
            PrimarySmtpAddress        = $mbx.PrimarySmtpAddress
            DisplayName               = $mbx.DisplayName
            ExternalDirectoryObjectId = $mbx.ExternalDirectoryObjectId
            RecipientTypeDetails      = $mbx.RecipientTypeDetails
            LegacyExchangeDN          = $mbx.LegacyExchangeDN
            X500Addresses             = $x500Addresses
        }
        MailboxDelegations      = [ordered]@{
            FullAccess   = [System.Collections.Generic.List[object]]::new()
            SendAs       = [System.Collections.Generic.List[object]]::new()
            SendOnBehalf = [System.Collections.Generic.List[object]]::new()
        }
        FolderPermissions       = [System.Collections.Generic.List[object]]::new()
        UnresolvablePermissions = $unresolvable
        ExportErrors            = $mbxErrors
    }

    # ─── A. FULL ACCESS ───────────────────────────────────────────────────────
    try {
        $faPerms = Invoke-EXOWithRetry -Description "FullAccess:$($mbx.UserPrincipalName)" -ScriptBlock {
            Get-EXOMailboxPermission -Identity $mbx.UserPrincipalName -ErrorAction Stop |
            Where-Object { $_.IsInherited -eq $false -and $_.User -ne $mbx.UserPrincipalName -and
                $_.AccessRights -contains "FullAccess" }
        }

        foreach ($fa in $faPerms) {
            $resolved = Get-ResolvedIdentity -RawValue $fa.User -PermissionType "FullAccess"

            if ($resolved.PermissionStatus -eq "Resolved" -or $resolved.PermissionStatus -eq "WellKnown") {
                $rec = New-DelegateRecord -Resolved $resolved -ExtraFields @{
                    AutoMapping = Get-OptionalObjectPropertyValue -InputObject $fa -PropertyName 'AutoMapping'
                }
                $mbxData.MailboxDelegations.FullAccess.Add($rec)
            }
            elseif ($resolved.PermissionStatus -eq "Unresolvable") {
                $unresolvable.Add([ordered]@{
                        PermissionType  = "FullAccess"
                        RawTrusteeValue = $resolved.RawTrusteeValue
                        Reason          = $resolved.UnresolvableReason
                    })
            }
        }
    }
    catch {
        $mbxErrors.Add("FullAccess: $_")
        Write-Warning "[$($mbx.UserPrincipalName)] Full Access failed: $_"
    }

    # ─── B. SEND AS ───────────────────────────────────────────────────────────
    try {
        $saPerms = Invoke-EXOWithRetry -Description "SendAs:$($mbx.UserPrincipalName)" -ScriptBlock {
            Get-EXORecipientPermission -Identity $mbx.UserPrincipalName -ErrorAction Stop |
            Where-Object { $_.IsInherited -eq $false -and $_.Trustee -ne $mbx.UserPrincipalName -and
                $_.AccessRights -contains "SendAs" }
        }

        foreach ($sa in $saPerms) {
            $resolved = Get-ResolvedIdentity -RawValue $sa.Trustee -PermissionType "SendAs"

            if ($resolved.PermissionStatus -eq "Resolved" -or $resolved.PermissionStatus -eq "WellKnown") {
                $rec = New-DelegateRecord -Resolved $resolved
                $mbxData.MailboxDelegations.SendAs.Add($rec)
            }
            elseif ($resolved.PermissionStatus -eq "Unresolvable") {
                $unresolvable.Add([ordered]@{
                        PermissionType  = "SendAs"
                        RawTrusteeValue = $resolved.RawTrusteeValue
                        Reason          = $resolved.UnresolvableReason
                    })
            }
        }
    }
    catch {
        $mbxErrors.Add("SendAs: $_")
        Write-Warning "[$($mbx.UserPrincipalName)] Send As failed: $_"
    }

    # ─── C. SEND ON BEHALF ────────────────────────────────────────────────────
    if ($mbx.GrantSendOnBehalfTo) {
        foreach ($sobRaw in $mbx.GrantSendOnBehalfTo) {
            $resolved = Get-ResolvedIdentity -RawValue $sobRaw -PermissionType "SendOnBehalf"

            if ($resolved.PermissionStatus -eq "Resolved" -or $resolved.PermissionStatus -eq "WellKnown") {
                $rec = New-DelegateRecord -Resolved $resolved
                $mbxData.MailboxDelegations.SendOnBehalf.Add($rec)
            }
            elseif ($resolved.PermissionStatus -eq "Unresolvable") {
                $unresolvable.Add([ordered]@{
                        PermissionType  = "SendOnBehalf"
                        RawTrusteeValue = $resolved.RawTrusteeValue
                        Reason          = $resolved.UnresolvableReason
                    })
            }
        }
    }

    # ─── D. FOLDER PERMISSIONS (full scan or explicit folder selection) ─────
    if (-not $SkipFolderPermissions) {
        try {
            $folderStats = Invoke-EXOWithRetry -Description "FolderStats:$($mbx.UserPrincipalName)" -ScriptBlock {
                Get-EXOMailboxFolderStatistics -Identity $mbx.UserPrincipalName `
                    -FolderScope All -ErrorAction Stop
            }

            if ($usingExplicitFolderSelection) {
                # Explicit selection mode: only scan folders the admin requested.
                $eligibleFolders = @($folderStats | Where-Object {
                        Test-FolderSelection -Folder $_ -Selectors $IncludedFolders
                    })
            }
            else {
                # Default mode: scan every folder when no explicit selection was provided.
                $eligibleFolders = @($folderStats)
            }

            $folderTotal = $eligibleFolders.Count
            $folderCount = 0

            foreach ($folder in $eligibleFolders) {
                $folderCount++
                $folderPct = [int](($folderCount / [math]::Max($folderTotal, 1)) * 100)
                Write-Progress -Id 1 -ParentId 0 `
                    -Activity "  Scanning folders for $($mbx.UserPrincipalName)" `
                    -Status "$($folder.FolderPath) ($folderCount of $folderTotal)" `
                    -PercentComplete $folderPct

                $rawFolderPath = $folder.FolderPath
                $folderIdentity = "$($mbx.PrimarySmtpAddress):$rawFolderPath"

                try {
                    $folderPerms = Invoke-EXOWithRetry -Description "FolderPerms:$folderIdentity" -ScriptBlock {
                        Get-EXOMailboxFolderPermission -Identity $folderIdentity -ErrorAction Stop
                    }

                    # EARLY EXIT: If the folder only has Default and/or Anonymous entries,
                    # there is nothing interesting to record — skip immediately.
                    $nonDefaultPerms = @($folderPerms | Where-Object {
                            $_.User -notin @("Default", "Anonymous") -and
                            $_.AccessRights -notcontains "None"
                        })

                    if ($nonDefaultPerms.Count -eq 0) { continue }

                    $permRecords = [System.Collections.Generic.List[object]]::new()

                    foreach ($fp in $folderPerms) {
                        if ($fp.AccessRights -contains "None") { continue }

                        $resolved = Get-ResolvedIdentity -RawValue $fp.User `
                            -PermissionType "FolderPermission" -FolderPath $rawFolderPath

                        if ($resolved.PermissionStatus -eq "Resolved") {
                            if ($resolved.DelegateUPN -eq $mbx.PrimarySmtpAddress) { continue }

                            $rec = New-DelegateRecord -Resolved $resolved -ExtraFields @{
                                AccessRights           = @($fp.AccessRights)
                                SharingPermissionFlags = $fp.SharingPermissionFlags
                            }
                            $permRecords.Add($rec)
                        }
                        elseif ($resolved.PermissionStatus -eq "WellKnown") {
                            $rec = New-DelegateRecord -Resolved $resolved -ExtraFields @{
                                AccessRights           = @($fp.AccessRights)
                                SharingPermissionFlags = $fp.SharingPermissionFlags
                            }
                            $permRecords.Add($rec)
                        }
                        elseif ($resolved.PermissionStatus -eq "Unresolvable") {
                            $unresolvable.Add([ordered]@{
                                    PermissionType  = "FolderPermission"
                                    FolderPath      = $rawFolderPath
                                    RawTrusteeValue = $resolved.RawTrusteeValue
                                    Reason          = $resolved.UnresolvableReason
                                })
                        }
                    }

                    if ($permRecords.Count -gt 0) {
                        $mbxData.FolderPermissions.Add([ordered]@{
                                FolderPath  = $rawFolderPath
                                FolderType  = $folder.FolderType
                                Permissions = $permRecords
                            })
                    }
                }
                catch {
                    $global:FolderErrors++
                    Write-Verbose "[$($mbx.UserPrincipalName)] Folder perm error on '$rawFolderPath': $_"
                }
            }

            Write-Progress -Id 1 -ParentId 0 -Activity "  Scanning folders" -Completed
        }
        catch {
            $mbxErrors.Add("FolderStatistics: $_")
            Write-Warning "[$($mbx.UserPrincipalName)] Folder stats failed: $_"
        }
    }

    $exportList.Add($mbxData)

    # Per-mailbox timing & stats
    $mbxSw.Stop()
    $mbxElapsed = $mbxSw.Elapsed.ToString("mm\:ss\.f")
    $delegationCount = $mbxData.MailboxDelegations.FullAccess.Count +
    $mbxData.MailboxDelegations.SendAs.Count +
    $mbxData.MailboxDelegations.SendOnBehalf.Count
    $folderPermCount = 0
    foreach ($fpEntry in $mbxData.FolderPermissions) { $folderPermCount += $fpEntry.Permissions.Count }

    Write-Host "  [$count/$total] $($mbx.UserPrincipalName) — ${mbxElapsed} | $delegationCount delegations, $folderPermCount folder perms, $($unresolvable.Count) unresolvable" -ForegroundColor DarkGray

    # Incremental save
    if ($processedThisRun -gt 0 -and (($processedThisRun % $SaveInterval) -eq 0 -or $count -eq $total)) {
        Write-Host "  ---> Saving progress ($($exportList.Count)/$total mailboxes)..." -ForegroundColor Yellow
        try {
            $tempOutput = "$OutputPath.tmp"
            $currentOutput = [ordered]@{
                ExportMetadata = [ordered]@{
                    ExportedAt                 = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                    ScriptVersion              = $ScriptVersion
                    SourceTenantDomain         = $sourceTenantDomain
                    SourceTenantId             = $sourceTenantId
                    TotalMailboxes             = $total
                    TotalMailboxesExported     = $exportList.Count
                    TotalIdentityCacheSize     = $global:IdentityCache.Count
                    TotalResolveErrors         = $global:ResolveErrors
                    TotalFolderErrors          = $global:FolderErrors
                    RecipientTypesIncluded     = $RecipientTypeDetails
                    IncludedFolders            = $IncludedFolders
                    FolderPermissionsCollected = (-not $SkipFolderPermissions)
                    IsPartial                  = ($exportList.Count -lt $total)
                }
                Mailboxes      = $exportList
            }
            ConvertTo-Json -InputObject $currentOutput -Depth 20 | Out-File -FilePath $tempOutput -Encoding utf8 -Force
            Move-Item -Path $tempOutput -Destination $OutputPath -Force
        }
        catch {
            Write-Warning "Failed to incrementally save to $OutputPath. Error: $($_.Exception.Message) - $($_.InvocationInfo.PositionMessage)"
        }
    }

    # Light throttle between mailboxes
    if ($ThrottleDelayMs -gt 0) {
        Start-Sleep -Milliseconds $ThrottleDelayMs
    }
}

Write-Progress -Id 0 -Activity "Extracting EXO Permissions" -Completed
$global:sw.Stop()

# ==============================================================================
# SECTION 7 — Build final output document
# ==============================================================================

$output = [ordered]@{
    ExportMetadata = [ordered]@{
        ExportedAt                 = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        ScriptVersion              = $ScriptVersion
        SourceTenantDomain         = $sourceTenantDomain
        SourceTenantId             = $sourceTenantId
        TotalMailboxes             = $total
        TotalMailboxesExported     = $exportList.Count
        TotalIdentityCacheSize     = $global:IdentityCache.Count
        TotalResolveErrors         = $global:ResolveErrors
        TotalFolderErrors          = $global:FolderErrors
        RecipientTypesIncluded     = $RecipientTypeDetails
        IncludedFolders            = $IncludedFolders
        FolderPermissionsCollected = (-not $SkipFolderPermissions)
    }
    Mailboxes      = $exportList
}

# ==============================================================================
# SECTION 8 — Write JSON
# ==============================================================================

Write-Host "`nSerializing to JSON..." -ForegroundColor Cyan

try {
    $output | ConvertTo-Json -Depth 20 | Out-File -FilePath $OutputPath -Encoding utf8 -Force
    Write-Host "Export complete: $OutputPath" -ForegroundColor Green
}
catch {
    Write-Error "Failed to write output file: $_"
}

# ==============================================================================
# SECTION 9 — Summary
# ==============================================================================

Write-Host "`n========== EXPORT SUMMARY ==========" -ForegroundColor Cyan
Write-Host "Mailboxes processed  : $total"
Write-Host "Identity cache size  : $($global:IdentityCache.Count) unique identity keys cached"
Write-Host "Resolve errors       : $global:ResolveErrors (logged as Unresolvable)"
Write-Host "Folder errors        : $global:FolderErrors (suppressed)"
Write-Host "Total elapsed time   : $($global:sw.Elapsed.ToString('hh\:mm\:ss'))"
Write-Host "Avg per mailbox      : $([math]::Round($global:sw.Elapsed.TotalSeconds / [math]::Max($total, 1), 1))s"
Write-Host "Output file          : $OutputPath"
Write-Host "====================================`n" -ForegroundColor Cyan

Disconnect-ExchangeOnline -Confirm:$false