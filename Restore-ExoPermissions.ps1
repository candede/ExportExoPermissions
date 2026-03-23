# ==============================================================================
# Script  : Restore-ExoPermissions.ps1
# Version : 1.2
# Purpose : Restore mailbox-level and folder-level Exchange Online permissions
#           on a TARGET tenant using the JSON export produced by
#           Export-ExoPermissions-HybridReady.ps1 and a UPN mapping CSV.
#
# Permission types restored:
#   - X500 aliases       (Set-Mailbox -EmailAddresses @{Add=...})
#   - Full Access        (Add-MailboxPermission)
#   - Send As            (Add-RecipientPermission)
#   - Send on Behalf     (Set-Mailbox -GrantSendOnBehalfTo)
#   - Folder Permissions (Add-MailboxFolderPermission / Set-MailboxFolderPermission)
#
# Mapping CSV must contain at minimum:
#   SourceUPN, TargetUPN
#   (optional extra columns are ignored)
#
# Run modes:
#   -WhatIf      : Simulate all actions, no changes made
#   -SkipExisting: Skip permissions that already exist (default: true)
#   -PermTypes   : Limit to specific permission types
#   -UseEnvCredentials: Use .env app-only authentication instead of interactive sign-in
# ==============================================================================

#Requires -Modules ExchangeOnlineManagement

[CmdletBinding(SupportsShouldProcess)]
param (
    # Path to the JSON file exported by Export-ExoPermissions-HybridReady.ps1
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ExportJsonPath,

    # Path to CSV with at least SourceUPN and TargetUPN columns
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$MappingCsvPath,

    # Limit which permission types to restore
    [ValidateSet("X500Addresses", "FullAccess", "SendAs", "SendOnBehalf", "FolderPermissions", "All")]
    [string[]]$PermTypes = @("All"),

    # Skip adding a permission if it already exists on the target mailbox
    [bool]$SkipExisting = $true,

    # How long to sleep between mailboxes to avoid throttling (ms)
    [int]$ThrottleDelayMs = 200,

    # Max retries for throttled calls
    [int]$MaxRetries = 3,

    # Where to write the restore log (CSV). Defaults to same dir as this script.
    [string]$LogPath = "",

    # Optional: only restore for the specified mailbox email addresses.
    # Matches source UPN, source primary SMTP, or mapped target UPN.
    [string[]]$TestMailboxes = @(),

    # Use tenant/app credentials from the .env file instead of interactive sign-in.
    [switch]$UseEnvCredentials,

    # Optional path to the .env file. Defaults to .env in the script directory.
    [string]$EnvFilePath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# ==============================================================================
# SECTION 0 — Setup & Logging
# ==============================================================================

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

if ([string]::IsNullOrWhiteSpace($EnvFilePath)) {
    $EnvFilePath = Join-Path $scriptDir ".env"
}

# Convenience behavior: if the caller explicitly passed EnvFilePath, prefer app-only auth.
if (-not $UseEnvCredentials -and $PSBoundParameters.ContainsKey('EnvFilePath')) {
    $UseEnvCredentials = $true
    Write-Host "EnvFilePath was provided explicitly. Enabling app-only authentication." -ForegroundColor Yellow
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path $scriptDir "RestoreLog_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
}

# Restore log — every action taken or skipped is written here
$restoreLog = [System.Collections.Generic.List[object]]::new()

function Write-LogEntry {
    param(
        [string]$TargetMailbox,
        [string]$PermType,
        [string]$Delegate,
        [string]$Detail,
        [ValidateSet("Applied", "Skipped", "WhatIf", "Error", "Warning", "Info")]
        [string]$Status,
        [string]$Message = ""
    )
    $entry = [PSCustomObject][ordered]@{
        Timestamp     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        TargetMailbox = $TargetMailbox
        PermType      = $PermType
        Delegate      = $Delegate
        Detail        = $Detail
        Status        = $Status
        Message       = $Message
    }
    $restoreLog.Add($entry)

    $color = switch ($Status) {
        "Applied" { "Green" }
        "Skipped" { "DarkGray" }
        "WhatIf" { "Cyan" }
        "Error" { "Red" }
        "Warning" { "Yellow" }
        default { "White" }
    }
    Write-Host "[$Status] $PermType | $TargetMailbox | $Delegate | $Detail $(if($Message){ "| $Message" })" -ForegroundColor $color
}

# ==============================================================================
# SECTION 1 — Helpers
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
                $wait = [math]::Pow(2, $attempt) * 1000
                Write-Warning "Throttled on '$Description' — retrying in $($wait/1000)s (attempt $attempt/$MaxAttempts)"
                Start-Sleep -Milliseconds $wait
            }
            else { throw }
        }
    } while ($attempt -lt $MaxAttempts)
}

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

# Resolve a SourceUPN to a TargetUPN via the mapping table
function Resolve-TargetUPN {
    param([string]$SourceUPN, [hashtable]$Mapping)
    if ([string]::IsNullOrWhiteSpace($SourceUPN)) { return $null }
    $key = $SourceUPN.ToLower()
    if ($Mapping.ContainsKey($key)) { return $Mapping[$key] }
    return $null  # not in scope of migration
}

# Check if a ShouldProcess permission type is requested
function Test-PermTypeRequested {
    param([string]$Type)
    return ($PermTypes -contains "All" -or $PermTypes -contains $Type)
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

function Get-ExportX500Addresses {
    param([object]$MailboxRecord)

    $x500Addresses = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)

    $sourceUser = $MailboxRecord.SourceUser
    if ($null -eq $sourceUser) {
        return @()
    }

    $x500Property = $sourceUser.PSObject.Properties['X500Addresses']
    if ($x500Property -and $null -ne $x500Property.Value) {
        foreach ($x500Address in @($x500Property.Value)) {
            $normalizedAddress = Normalize-X500Address -Address ([string]$x500Address)
            if (-not [string]::IsNullOrWhiteSpace($normalizedAddress)) {
                $x500Addresses.Add($normalizedAddress) | Out-Null
            }
        }
    }

    $legacyExchangeDnProperty = $sourceUser.PSObject.Properties['LegacyExchangeDN']
    if ($legacyExchangeDnProperty -and -not [string]::IsNullOrWhiteSpace([string]$legacyExchangeDnProperty.Value)) {
        $normalizedLegacyAddress = Normalize-X500Address -Address ([string]$legacyExchangeDnProperty.Value)
        if (-not [string]::IsNullOrWhiteSpace($normalizedLegacyAddress)) {
            $x500Addresses.Add($normalizedLegacyAddress) | Out-Null
        }
    }

    return @($x500Addresses | Sort-Object)
}

# ==============================================================================
# SECTION 2 — Load inputs
# ==============================================================================

Write-Host "`n=== EXO Permission Restore v1.2 ===" -ForegroundColor Cyan
Write-Host "Export JSON : $ExportJsonPath"
Write-Host "Mapping CSV : $MappingCsvPath"
Write-Host "Log Output  : $LogPath"
Write-Host "WhatIf Mode : $WhatIfPreference"
Write-Host "PermTypes   : $($PermTypes -join ', ')"
Write-Host "Test scope  : $(if ($TestMailboxes.Count -gt 0) { $TestMailboxes -join ', ' } else { 'Disabled (all exported mailboxes)' })"
Write-Host "Auth Mode   : $(if ($UseEnvCredentials) { '.env app-only' } else { 'Interactive' })"
Write-Host "Started     : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"

# --- Load UPN mapping ---
Write-Host "Loading UPN mapping..." -ForegroundColor Cyan
$mappingCsv = Import-Csv -Path $MappingCsvPath

if (-not ($mappingCsv | Get-Member -Name "SourceUPN") -or
    -not ($mappingCsv | Get-Member -Name "TargetUPN")) {
    Write-Error "Mapping CSV must contain 'SourceUPN' and 'TargetUPN' columns."
    exit 1
}

# Build a hashtable: sourceupn (lowercase) -> targetUPN
$upnMap = @{}
foreach ($row in $mappingCsv) {
    $src = $row.SourceUPN.Trim().ToLower()
    $tgt = $row.TargetUPN.Trim()
    if (-not [string]::IsNullOrWhiteSpace($src) -and -not [string]::IsNullOrWhiteSpace($tgt)) {
        $upnMap[$src] = $tgt
    }
}
Write-Host "Loaded $($upnMap.Count) UPN mappings." -ForegroundColor Green

# --- Load export JSON ---
Write-Host "Loading export JSON..." -ForegroundColor Cyan
try {
    $exportData = Get-Content -Raw -Path $ExportJsonPath | ConvertFrom-Json
}
catch {
    Write-Error "Failed to parse export JSON: $_"
    exit 1
}

$mailboxes = $exportData.Mailboxes

if ($TestMailboxes.Count -gt 0) {
    $testMailboxSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($testMailbox in $TestMailboxes) {
        if (-not [string]::IsNullOrWhiteSpace($testMailbox)) {
            $testMailboxSet.Add($testMailbox.Trim()) | Out-Null
        }
    }

    $mailboxes = @($mailboxes | Where-Object {
            $sourceUser = $_.SourceUser
            $sourceUPN = Get-OptionalObjectPropertyValue -InputObject $sourceUser -PropertyName 'UserPrincipalName'
            $sourcePrimarySmtp = Get-OptionalObjectPropertyValue -InputObject $sourceUser -PropertyName 'PrimarySmtpAddress'
            $targetUPN = Resolve-TargetUPN -SourceUPN $sourceUPN -Mapping $upnMap

            $testMailboxSet.Contains([string]$sourceUPN) -or
            $testMailboxSet.Contains([string]$sourcePrimarySmtp) -or
            $testMailboxSet.Contains([string]$targetUPN)
        })

    Write-Host "TestMailboxes filter active. Restoring $($mailboxes.Count) mailbox(es)." -ForegroundColor Yellow
}

$totalMbx = $mailboxes.Count
$mbxCount = 0

Write-Host "Found $totalMbx mailboxes in export.`n" -ForegroundColor Green

# --- Connect to TARGET tenant ---
Write-Host "Connecting to Exchange Online (TARGET tenant)..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnlineSession -UseEnvCredentials:$UseEnvCredentials -EnvFilePath $EnvFilePath
}
catch {
    Write-Error "Failed to connect to Exchange Online: $_"
    exit 1
}

# ==============================================================================
# SECTION 3 — Cache of existing target mailbox permissions
#             Loaded on demand per mailbox to support -SkipExisting
# ==============================================================================

# Cache: targetUPN -> @{ X500Addresses=[]; FullAccess=[]; SendAs=[]; SendOnBehalf=[] }
$targetPermCache = @{}

function Get-ExistingMailboxPerms {
    param([string]$TargetUPN)

    if ($targetPermCache.ContainsKey($TargetUPN)) {
        return $targetPermCache[$TargetUPN]
    }

    $perms = @{
        X500Addresses = @()
        FullAccess    = @()
        SendAs        = @()
        SendOnBehalf  = @()
    }

    try {
        $perms.FullAccess = @(
            Invoke-EXOWithRetry -Description "ExistingFA:$TargetUPN" -ScriptBlock {
                Get-EXOMailboxPermission -Identity $TargetUPN -ErrorAction Stop |
                Where-Object { $_.IsInherited -eq $false } |
                Select-Object -ExpandProperty User
            }
        )
    }
    catch { Write-Verbose "Could not pre-cache FullAccess for $TargetUPN" }

    try {
        $perms.SendAs = @(
            Invoke-EXOWithRetry -Description "ExistingSA:$TargetUPN" -ScriptBlock {
                Get-EXORecipientPermission -Identity $TargetUPN -ErrorAction Stop |
                Where-Object { $_.IsInherited -eq $false } |
                Select-Object -ExpandProperty Trustee
            }
        )
    }
    catch { Write-Verbose "Could not pre-cache SendAs for $TargetUPN" }

    try {
        $mbxObj = Invoke-EXOWithRetry -Description "ExistingSOB:$TargetUPN" -ScriptBlock {
            Get-EXOMailbox -Identity $TargetUPN -Properties GrantSendOnBehalfTo, EmailAddresses -ErrorAction Stop
        }
        $perms.SendOnBehalf = @($mbxObj.GrantSendOnBehalfTo)

        $existingX500 = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase)
        foreach ($emailAddress in @($mbxObj.EmailAddresses)) {
            $addressString = [string]$emailAddress
            if ($addressString -match '^(?i)x500:') {
                $normalizedAddress = Normalize-X500Address -Address $addressString
                if (-not [string]::IsNullOrWhiteSpace($normalizedAddress)) {
                    $existingX500.Add($normalizedAddress) | Out-Null
                }
            }
        }
        $perms.X500Addresses = @($existingX500 | Sort-Object)
    }
    catch { Write-Verbose "Could not pre-cache SendOnBehalf for $TargetUPN" }

    $targetPermCache[$TargetUPN] = $perms
    return $perms
}

# ==============================================================================
# SECTION 4 — Main restore loop
# ==============================================================================

foreach ($mbxRecord in $mailboxes) {
    $mbxCount++
    $srcUPN = $mbxRecord.SourceUser.UserPrincipalName
    $tgtUPN = Resolve-TargetUPN -SourceUPN $srcUPN -Mapping $upnMap

    $pct = if ($totalMbx -gt 0) { [int](($mbxCount / $totalMbx) * 100) } else { 0 }
    Write-Progress -Activity "Restoring EXO Permissions" `
        -Status "[$mbxCount/$totalMbx] $srcUPN -> $(if ($tgtUPN) { $tgtUPN } else { 'NOT IN MAPPING' })" `
        -PercentComplete $pct

    # Skip mailboxes not in mapping (out of scope)
    if ($null -eq $tgtUPN) {
        Write-LogEntry -TargetMailbox $srcUPN -PermType "N/A" -Delegate "N/A" -Detail "N/A" `
            -Status "Skipped" -Message "SourceUPN not found in mapping CSV"
        continue
    }

    # Pre-load existing permissions on target if SkipExisting is on
    $existing = if ($SkipExisting) { Get-ExistingMailboxPerms -TargetUPN $tgtUPN } else { @{} }

    # ─── 0. X500 ADDRESSES ───────────────────────────────────────────────────
    if (Test-PermTypeRequested "X500Addresses") {
        $sourceX500Addresses = @(Get-ExportX500Addresses -MailboxRecord $mbxRecord)

        if ($sourceX500Addresses.Count -gt 0) {
            $existingX500Set = [System.Collections.Generic.HashSet[string]]::new(
                [System.StringComparer]::OrdinalIgnoreCase)

            foreach ($existingX500Address in @($existing.X500Addresses)) {
                $normalizedExistingAddress = Normalize-X500Address -Address ([string]$existingX500Address)
                if (-not [string]::IsNullOrWhiteSpace($normalizedExistingAddress)) {
                    $existingX500Set.Add($normalizedExistingAddress) | Out-Null
                }
            }

            $x500ToAdd = [System.Collections.Generic.List[string]]::new()
            foreach ($sourceX500Address in $sourceX500Addresses) {
                if (-not ($SkipExisting -and $existingX500Set.Contains($sourceX500Address))) {
                    $x500ToAdd.Add($sourceX500Address)
                }
            }

            if ($x500ToAdd.Count -eq 0) {
                Write-LogEntry -TargetMailbox $tgtUPN -PermType "X500Addresses" `
                    -Delegate "(mailbox)" -Detail "Count=$($sourceX500Addresses.Count)" `
                    -Status "Skipped" -Message "All exported X500 aliases already exist on target"
            }
            elseif ($PSCmdlet.ShouldProcess($tgtUPN, "Add $($x500ToAdd.Count) X500 alias(es)")) {
                try {
                    Invoke-EXOWithRetry -Description "AddX500:$tgtUPN" -ScriptBlock {
                        Set-Mailbox -Identity $tgtUPN `
                            -EmailAddresses @{ Add = $x500ToAdd.ToArray() } `
                            -ErrorAction Stop
                    }
                    Write-LogEntry -TargetMailbox $tgtUPN -PermType "X500Addresses" `
                        -Delegate "(mailbox)" -Detail "Count=$($x500ToAdd.Count)" `
                        -Status "Applied" -Message ($x500ToAdd -join '; ')
                }
                catch {
                    Write-LogEntry -TargetMailbox $tgtUPN -PermType "X500Addresses" `
                        -Delegate "(mailbox)" -Detail "Count=$($x500ToAdd.Count)" `
                        -Status "Error" -Message "$_"
                }
            }
            else {
                Write-LogEntry -TargetMailbox $tgtUPN -PermType "X500Addresses" `
                    -Delegate "(mailbox)" -Detail "Count=$($x500ToAdd.Count)" `
                    -Status "WhatIf" -Message ($x500ToAdd -join '; ')
            }
        }
    }

    # ─── A. FULL ACCESS ───────────────────────────────────────────────────────
    if (Test-PermTypeRequested "FullAccess") {
        foreach ($fa in $mbxRecord.MailboxDelegations.FullAccess) {
            $autoMapping = Get-OptionalObjectPropertyValue -InputObject $fa -PropertyName 'AutoMapping'

            if ($fa.PermissionStatus -ne "Resolved" -and $fa.PermissionStatus -ne "WellKnown") {
                Write-LogEntry -TargetMailbox $tgtUPN -PermType "FullAccess" `
                    -Delegate $fa.DelegateUPN -Detail "AutoMapping=$autoMapping" `
                    -Status "Skipped" -Message "PermissionStatus=$($fa.PermissionStatus) — delegate not resolvable at export time"
                continue
            }

            $delTgtUPN = Resolve-TargetUPN -SourceUPN $fa.DelegateUPN -Mapping $upnMap
            if ($null -eq $delTgtUPN) {
                Write-LogEntry -TargetMailbox $tgtUPN -PermType "FullAccess" `
                    -Delegate $fa.DelegateUPN -Detail "AutoMapping=$autoMapping" `
                    -Status "Skipped" -Message "Delegate not in mapping CSV — skipped"
                continue
            }

            if ($SkipExisting -and ($existing.FullAccess -contains $delTgtUPN)) {
                Write-LogEntry -TargetMailbox $tgtUPN -PermType "FullAccess" `
                    -Delegate $delTgtUPN -Detail "AutoMapping=$autoMapping" `
                    -Status "Skipped" -Message "Already exists on target"
                continue
            }

            if ($PSCmdlet.ShouldProcess($tgtUPN, "Add FullAccess for $delTgtUPN")) {
                try {
                    Invoke-EXOWithRetry -Description "AddFA:$tgtUPN->$delTgtUPN" -ScriptBlock {
                        $addPermissionParams = @{
                            Identity     = $tgtUPN
                            User         = $delTgtUPN
                            AccessRights = 'FullAccess'
                            ErrorAction  = 'Stop'
                        }

                        if ($null -ne $autoMapping) {
                            $addPermissionParams.AutoMapping = [bool]$autoMapping
                        }

                        Add-MailboxPermission @addPermissionParams | Out-Null
                    }
                    Write-LogEntry -TargetMailbox $tgtUPN -PermType "FullAccess" `
                        -Delegate $delTgtUPN -Detail "AutoMapping=$autoMapping" -Status "Applied"
                }
                catch {
                    Write-LogEntry -TargetMailbox $tgtUPN -PermType "FullAccess" `
                        -Delegate $delTgtUPN -Detail "AutoMapping=$autoMapping" `
                        -Status "Error" -Message "$_"
                }
            }
            else {
                Write-LogEntry -TargetMailbox $tgtUPN -PermType "FullAccess" `
                    -Delegate $delTgtUPN -Detail "AutoMapping=$autoMapping" -Status "WhatIf"
            }
        }
    }

    # ─── B. SEND AS ───────────────────────────────────────────────────────────
    if (Test-PermTypeRequested "SendAs") {
        foreach ($sa in $mbxRecord.MailboxDelegations.SendAs) {
            if ($sa.PermissionStatus -ne "Resolved" -and $sa.PermissionStatus -ne "WellKnown") {
                Write-LogEntry -TargetMailbox $tgtUPN -PermType "SendAs" `
                    -Delegate $sa.DelegateUPN -Detail "-" `
                    -Status "Skipped" -Message "PermissionStatus=$($sa.PermissionStatus)"
                continue
            }

            $delTgtUPN = Resolve-TargetUPN -SourceUPN $sa.DelegateUPN -Mapping $upnMap
            if ($null -eq $delTgtUPN) {
                Write-LogEntry -TargetMailbox $tgtUPN -PermType "SendAs" `
                    -Delegate $sa.DelegateUPN -Detail "-" `
                    -Status "Skipped" -Message "Delegate not in mapping CSV"
                continue
            }

            if ($SkipExisting -and ($existing.SendAs -contains $delTgtUPN)) {
                Write-LogEntry -TargetMailbox $tgtUPN -PermType "SendAs" `
                    -Delegate $delTgtUPN -Detail "-" -Status "Skipped" -Message "Already exists on target"
                continue
            }

            if ($PSCmdlet.ShouldProcess($tgtUPN, "Add SendAs for $delTgtUPN")) {
                try {
                    Invoke-EXOWithRetry -Description "AddSA:$tgtUPN->$delTgtUPN" -ScriptBlock {
                        Add-RecipientPermission -Identity $tgtUPN `
                            -Trustee $delTgtUPN `
                            -AccessRights SendAs `
                            -Confirm:$false `
                            -ErrorAction Stop | Out-Null
                    }
                    Write-LogEntry -TargetMailbox $tgtUPN -PermType "SendAs" -Delegate $delTgtUPN -Detail "-" -Status "Applied"
                }
                catch {
                    Write-LogEntry -TargetMailbox $tgtUPN -PermType "SendAs" `
                        -Delegate $delTgtUPN -Detail "-" -Status "Error" -Message "$_"
                }
            }
            else {
                Write-LogEntry -TargetMailbox $tgtUPN -PermType "SendAs" -Delegate $delTgtUPN -Detail "-" -Status "WhatIf"
            }
        }
    }

    # ─── C. SEND ON BEHALF ────────────────────────────────────────────────────
    if (Test-PermTypeRequested "SendOnBehalf") {
        $sobTargetList = [System.Collections.Generic.List[string]]::new()
        $existingSOBSet = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase)

        # Collect already-existing Send on Behalf trustees on target (to merge, not overwrite)
        if ($SkipExisting) {
            foreach ($existingSOB in $existing.SendOnBehalf) {
                if (-not [string]::IsNullOrWhiteSpace($existingSOB)) {
                    $sobTargetList.Add($existingSOB)
                    $existingSOBSet.Add($existingSOB) | Out-Null
                }
            }
        }

        $sobNewDelegates = [System.Collections.Generic.List[string]]::new()

        foreach ($sob in $mbxRecord.MailboxDelegations.SendOnBehalf) {
            if ($sob.PermissionStatus -ne "Resolved" -and $sob.PermissionStatus -ne "WellKnown") {
                Write-LogEntry -TargetMailbox $tgtUPN -PermType "SendOnBehalf" `
                    -Delegate $sob.DelegateUPN -Detail "-" `
                    -Status "Skipped" -Message "PermissionStatus=$($sob.PermissionStatus)"
                continue
            }

            $delTgtUPN = Resolve-TargetUPN -SourceUPN $sob.DelegateUPN -Mapping $upnMap
            if ($null -eq $delTgtUPN) {
                Write-LogEntry -TargetMailbox $tgtUPN -PermType "SendOnBehalf" `
                    -Delegate $sob.DelegateUPN -Detail "-" `
                    -Status "Skipped" -Message "Delegate not in mapping CSV"
                continue
            }

            if ($SkipExisting -and $existingSOBSet.Contains($delTgtUPN)) {
                Write-LogEntry -TargetMailbox $tgtUPN -PermType "SendOnBehalf" `
                    -Delegate $delTgtUPN -Detail "-" -Status "Skipped" -Message "Already exists on target"
                continue
            }

            $sobTargetList.Add($delTgtUPN)
            $sobNewDelegates.Add($delTgtUPN)
        }

        # Set-Mailbox takes the FULL list — apply once after collecting all trustees
        if ($sobNewDelegates.Count -gt 0) {
            if ($PSCmdlet.ShouldProcess($tgtUPN, "Set SendOnBehalf for $($sobNewDelegates -join ', ')")) {
                try {
                    Invoke-EXOWithRetry -Description "SetSOB:$tgtUPN" -ScriptBlock {
                        Set-Mailbox -Identity $tgtUPN `
                            -GrantSendOnBehalfTo $sobTargetList.ToArray() `
                            -ErrorAction Stop
                    }
                    foreach ($del in $sobNewDelegates) {
                        Write-LogEntry -TargetMailbox $tgtUPN -PermType "SendOnBehalf" `
                            -Delegate $del -Detail "-" -Status "Applied"
                    }
                }
                catch {
                    Write-LogEntry -TargetMailbox $tgtUPN -PermType "SendOnBehalf" `
                        -Delegate "(batch)" -Detail "List=$($sobNewDelegates -join ',')" `
                        -Status "Error" -Message "$_"
                }
            }
            else {
                foreach ($del in $sobNewDelegates) {
                    Write-LogEntry -TargetMailbox $tgtUPN -PermType "SendOnBehalf" `
                        -Delegate $del -Detail "-" -Status "WhatIf"
                }
            }
        }
    }

    # ─── D. FOLDER PERMISSIONS ────────────────────────────────────────────────
    if (Test-PermTypeRequested "FolderPermissions") {
        foreach ($folderRecord in $mbxRecord.FolderPermissions) {
            $rawFolderPath = $folderRecord.FolderPath   # e.g. \Inbox
            $folderIdentity = "$($tgtUPN):$rawFolderPath"

            # Pre-load existing folder permissions for this specific folder
            $existingFolderPerms = @()
            if ($SkipExisting) {
                try {
                    $existingFolderPerms = @(
                        Invoke-EXOWithRetry -Description "ExistingFolderPerms:$folderIdentity" -ScriptBlock {
                            Get-EXOMailboxFolderPermission -Identity $folderIdentity -ErrorAction Stop
                        }
                    )
                }
                catch {
                    Write-Verbose "Could not pre-load folder perms for '$folderIdentity' — will attempt to set anyway"
                }
            }

            foreach ($fp in $folderRecord.Permissions) {
                # Skip Default/Anonymous unless they carry meaningful rights
                $isWellKnown = $fp.PermissionStatus -eq "WellKnown"

                if ($fp.PermissionStatus -ne "Resolved" -and -not $isWellKnown) {
                    Write-LogEntry -TargetMailbox $tgtUPN -PermType "FolderPermission" `
                        -Delegate $fp.DelegateUPN -Detail $rawFolderPath `
                        -Status "Skipped" -Message "PermissionStatus=$($fp.PermissionStatus)"
                    continue
                }

                # Map folder delegate to target
                $delTgtUPN = $fp.DelegateUPN  # Default / Anonymous kept as-is

                if (-not $isWellKnown) {
                    $delTgtUPN = Resolve-TargetUPN -SourceUPN $fp.DelegateUPN -Mapping $upnMap
                    if ($null -eq $delTgtUPN) {
                        Write-LogEntry -TargetMailbox $tgtUPN -PermType "FolderPermission" `
                            -Delegate $fp.DelegateUPN -Detail $rawFolderPath `
                            -Status "Skipped" -Message "Delegate not in mapping CSV"
                        continue
                    }
                }

                # Normalize AccessRights — comes from JSON as array or string
                $accessRights = if ($fp.AccessRights -is [array]) { $fp.AccessRights } else { @($fp.AccessRights) }

                # Check if this delegate already has a permission on this folder
                $existingEntry = $existingFolderPerms | Where-Object {
                    ([string]$_.User) -ieq $delTgtUPN
                }

                if ($PSCmdlet.ShouldProcess($folderIdentity, "Set FolderPermission for $delTgtUPN ($($accessRights -join ','))")) {
                    try {
                        if ($existingEntry) {
                            if ($SkipExisting) {
                                Write-LogEntry -TargetMailbox $tgtUPN -PermType "FolderPermission" `
                                    -Delegate $delTgtUPN -Detail "$rawFolderPath [$($accessRights -join ',')]" `
                                    -Status "Skipped" -Message "Permission already exists on target — use Set-MailboxFolderPermission to update"
                                continue
                            }
                            # Update existing permission
                            Invoke-EXOWithRetry -Description "SetFP:$folderIdentity->$delTgtUPN" -ScriptBlock {
                                Set-MailboxFolderPermission -Identity $folderIdentity `
                                    -User $delTgtUPN `
                                    -AccessRights $accessRights `
                                    -ErrorAction Stop | Out-Null
                            }
                        }
                        else {
                            # Add new permission
                            Invoke-EXOWithRetry -Description "AddFP:$folderIdentity->$delTgtUPN" -ScriptBlock {
                                Add-MailboxFolderPermission -Identity $folderIdentity `
                                    -User $delTgtUPN `
                                    -AccessRights $accessRights `
                                    -ErrorAction Stop | Out-Null
                            }
                        }
                        Write-LogEntry -TargetMailbox $tgtUPN -PermType "FolderPermission" `
                            -Delegate $delTgtUPN -Detail "$rawFolderPath [$($accessRights -join ',')]" -Status "Applied"
                    }
                    catch {
                        Write-LogEntry -TargetMailbox $tgtUPN -PermType "FolderPermission" `
                            -Delegate $delTgtUPN -Detail "$rawFolderPath [$($accessRights -join ',')]" `
                            -Status "Error" -Message "$_"
                    }
                }
                else {
                    Write-LogEntry -TargetMailbox $tgtUPN -PermType "FolderPermission" `
                        -Delegate $delTgtUPN -Detail "$rawFolderPath [$($accessRights -join ',')]" -Status "WhatIf"
                }
            }
        }
    }

    # Light throttle between mailboxes
    if ($ThrottleDelayMs -gt 0) {
        Start-Sleep -Milliseconds $ThrottleDelayMs
    }
}

Write-Progress -Activity "Restoring EXO Permissions" -Completed

# ==============================================================================
# SECTION 5 — Write restore log
# ==============================================================================

Write-Host "`nWriting restore log to: $LogPath" -ForegroundColor Cyan

$restoreLog | Export-Csv -Path $LogPath -NoTypeInformation -Encoding utf8 -Force

# ==============================================================================
# SECTION 6 — Summary
# ==============================================================================

$applied = ($restoreLog | Where-Object { $_.Status -eq "Applied" }).Count
$skipped = ($restoreLog | Where-Object { $_.Status -eq "Skipped" }).Count
$errors = ($restoreLog | Where-Object { $_.Status -eq "Error" }).Count
$whatifs = ($restoreLog | Where-Object { $_.Status -eq "WhatIf" }).Count
$warningCount = ($restoreLog | Where-Object { $_.Status -eq "Warning" }).Count

Write-Host "`n========== RESTORE SUMMARY ==========" -ForegroundColor Cyan
Write-Host "Mailboxes in scope   : $totalMbx"
Write-Host "Permissions applied  : $applied"  -ForegroundColor $(if ($applied -gt 0) { "Green" } else { "White" })
Write-Host "Permissions skipped  : $skipped"
Write-Host "WhatIf (not applied) : $whatifs"  -ForegroundColor Cyan
Write-Host "Warnings             : $warningCount" -ForegroundColor $(if ($warningCount -gt 0) { "Yellow" } else { "White" })
Write-Host "Errors               : $errors"   -ForegroundColor $(if ($errors -gt 0) { "Red" } else { "White" })
Write-Host "Restore log          : $LogPath"
Write-Host "====================================`n" -ForegroundColor Cyan

if ($errors -gt 0) {
    Write-Warning "$errors permissions could not be applied. Review the log for details."
    Write-Host "`nError summary:" -ForegroundColor Red
    $restoreLog | Where-Object { $_.Status -eq "Error" } |
    Format-Table TargetMailbox, PermType, Delegate, Detail, Message -AutoSize -Wrap
}

Disconnect-ExchangeOnline -Confirm:$false
