# Export EXO Permissions

PowerShell scripts to export and restore Exchange Online mailbox and folder permissions for tenant-to-tenant migration, remediation, or audit scenarios.

For more information, visit www.candede.com.

This repository contains:

- `Export-ExoPermissions-HybridReady.ps1`: exports mailbox-level and folder-level permissions from a source tenant into a restorable JSON document.
- `Restore-ExoPermissions.ps1`: restores supported permissions to a target tenant by using the export JSON and a source-to-target UPN mapping CSV.
- `.env.example`: sample app-only authentication file for unattended Exchange Online connections.
- `MailboxMap.example.csv`: sample UPN mapping file for restore runs.

The export script is designed for both cloud-only and hybrid environments. It records unresolvable trustees instead of failing the entire run, which is useful when source permissions still reference deleted or on-premises-only identities.

## Features

- Exports mailbox-level and folder-level Exchange Online permissions into a restorable JSON format.
- Supports both interactive sign-in and app-only `.env` authentication for export and restore.
- Handles hybrid edge cases by logging unresolvable trustees instead of failing the full run.
- Supports scoped exports by mailbox type or selected folders.
- Includes restore dry-run support with `-WhatIf` and action logging to CSV.
- Supports incremental export saves and resume behavior for long-running collections.

## Table of contents

- [Quick start](#quick-start)
- [What the scripts capture and restore](#what-the-scripts-capture-and-restore)
- [Prerequisites](#prerequisites)
- [Authentication methods](#authentication-methods)
- [App permissions required for `.env` authentication](#app-permissions-required-for-env-authentication)
- [Repository files](#repository-files)
- [Input file formats](#input-file-formats)
- [Logging and output files](#logging-and-output-files)
- [How unresolvable permissions are handled](#how-unresolvable-permissions-are-handled)
- [Throttling and retry behavior](#throttling-and-retry-behavior)
- [Recommended usage workflow](#recommended-usage-workflow)
- [Known limitations and implementation notes](#known-limitations-and-implementation-notes)
- [Publishing checklist](#publishing-checklist)
- [License](#license)

## Quick start

### 1. Install the required module

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

### 2. Prepare authentication

Choose one of the following:

- Interactive sign-in: no extra setup beyond having sufficient Exchange Online rights.
- App-only `.env` authentication: copy `.env.example` to `.env` and fill in `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, and `ORGANIZATION`.

### 3. Export permissions from the source tenant

Interactive:

```powershell
./Export-ExoPermissions-HybridReady.ps1
```

App-only:

```powershell
./Export-ExoPermissions-HybridReady.ps1 -UseEnvCredentials
```

Folder permission collection can take a long time in larger tenants because the script has to enumerate folders and read permissions mailbox by mailbox.

If you only need calendar permissions, limit the export to the Calendar folder:

```powershell
./Export-ExoPermissions-HybridReady.ps1 -IncludedFolders Calendar
```

If you want to use app-only authentication and only collect Calendar permissions:

```powershell
./Export-ExoPermissions-HybridReady.ps1 -UseEnvCredentials -IncludedFolders Calendar
```

If you do not need any folder permissions, skip them entirely to speed up the export:

```powershell
./Export-ExoPermissions-HybridReady.ps1 -SkipFolderPermissions
```

You can also limit the export to a small set of folders to reduce runtime, for example:

```powershell
./Export-ExoPermissions-HybridReady.ps1 -IncludedFolders Calendar,Inbox
```

### 4. Prepare the mailbox mapping CSV

Copy `MailboxMap.example.csv` and replace the sample values with your source and target UPN pairs.

Minimum format:

```csv
SourceUPN,TargetUPN
alice@source.com,alice@target.com
bob@source.com,bob@target.com
```

### 5. Dry-run the restore

Interactive:

```powershell
./Restore-ExoPermissions.ps1 \
  -ExportJsonPath ./ExoPermissions_20260309_1015.json \
  -MappingCsvPath ./MailboxMap.csv \
  -WhatIf
```

App-only:

```powershell
./Restore-ExoPermissions.ps1 \
  -ExportJsonPath ./ExoPermissions_20260309_1015.json \
  -MappingCsvPath ./MailboxMap.csv \
  -UseEnvCredentials \
  -WhatIf
```

### 6. Run the restore for real

```powershell
./Restore-ExoPermissions.ps1 \
  -ExportJsonPath ./ExoPermissions_20260309_1015.json \
  -MappingCsvPath ./MailboxMap.csv
```

Review the generated `RestoreLog_*.csv` after each run.

## What the scripts capture and restore

### Exported data

The export script captures the following for each mailbox:

- Source mailbox identity data:
  - `UserPrincipalName`
  - `PrimarySmtpAddress`
  - `DisplayName`
  - `ExternalDirectoryObjectId`
  - `RecipientTypeDetails`
  - `LegacyExchangeDN`
  - historical and current `X500` addresses
- Mailbox-level delegations:
  - `FullAccess`
  - `SendAs`
  - `SendOnBehalf`
- Folder-level permissions:
  - all folders by default, or only explicitly selected folders
- Unresolvable trustee entries:
  - orphaned SIDs
  - deleted recipients
  - other principals that cannot be resolved to an Exchange Online recipient
- Export metadata:
  - source tenant metadata
  - counts and error totals
  - included mailbox types
  - included folders
  - whether folder permissions were collected

### Restored data

The restore script can apply the following permission types:

- `X500Addresses`
- `FullAccess`
- `SendAs`
- `SendOnBehalf`
- `FolderPermissions`

## Prerequisites

### PowerShell

- PowerShell 7 is recommended for cross-platform execution.
- Windows PowerShell 5.1 may also work, but these scripts were written with modern Exchange Online cmdlets in mind.

### Required module

Both scripts declare the following requirement:

```powershell
#Requires -Modules ExchangeOnlineManagement
```

Install the module before running either script:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

If you already have it installed, it is still worth updating to the latest version:

```powershell
Update-Module ExchangeOnlineManagement
```

### Exchange Online connectivity

You must be able to connect to Exchange Online from the machine running the scripts.

Typical prerequisites:

- outbound HTTPS access to Microsoft 365 endpoints
- modern authentication enabled
- an account or app registration with sufficient Exchange Online rights

### Recommended operational prerequisites

Before running these scripts in production:

- test first in a pilot scope
- validate a small mapping CSV before a full restore
- run restore with `-WhatIf` before making changes
- keep the exported JSON and restore log files in a secure location because they contain mailbox and delegate metadata

## Authentication methods

## Interactive authentication

Both scripts support interactive Exchange Online sign-in. This is the default behavior for:

- `Export-ExoPermissions-HybridReady.ps1` when `-UseEnvCredentials` is not supplied
- `Restore-ExoPermissions.ps1` when `-UseEnvCredentials` is not supplied

The basic interactive workflow is shown in the Quick Start section above.

## `.env` app-only authentication

Both scripts support non-interactive app-only authentication when you pass `-UseEnvCredentials`. Use this for unattended runs, scheduled jobs, or environments where interactive sign-in is not practical.

In both scripts, this mode is implemented by:

1. Reading values from a `.env` file.
2. Requesting a client credentials token from Microsoft Entra ID.
3. Requesting the token for the Exchange Online resource scope:
   - `https://outlook.office365.com/.default`
4. Passing the resulting access token into `Connect-ExchangeOnline`.

## `.env` file format

By default, both scripts look for a file named `.env` in the same directory as the script. You can override that with `-EnvFilePath`.

Expected variables:

```dotenv
TENANT_ID=aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee
CLIENT_ID=11111111-2222-3333-4444-555555555555
CLIENT_SECRET=your-client-secret
ORGANIZATION=contoso.onmicrosoft.com
```

### Variable details

- `TENANT_ID`
  - Required.
  - The Microsoft Entra tenant ID where the Exchange Online organization exists.
- `CLIENT_ID`
  - Required.
  - The application ID of the Microsoft Entra app registration.
- `CLIENT_SECRET`
  - Required.
  - The client secret created for that app registration.
- `ORGANIZATION`
  - Optional in the script, but recommended.
  - Usually your Exchange Online organization domain, such as `contoso.onmicrosoft.com`.

If `ORGANIZATION` is omitted, the script warns and attempts the connection without it.

### How the scripts read `.env`

The parser used by both scripts:

- ignores commented lines that start with `#`
- expects `KEY=VALUE` format
- trims surrounding quotes from values
- does not require you to pre-load environment variables into the current shell session

### Security notes for `.env`

- Do not commit `.env` to GitHub.
- Add `.env` to `.gitignore` before publishing.
- Treat the client secret like a password.
- Prefer storing secrets in a secure vault or secret management system if you later automate this in CI/CD.

Suggested `.gitignore` entry:

```gitignore
.env
ExoPermissions_*.json
RestoreLog_*.csv
```

You may or may not want to ignore export JSON files in Git depending on your workflow, but they should generally not be committed because they contain tenant metadata and permission assignments.

## App permissions required for `.env` authentication

For app-only authentication to work with Exchange Online, the Entra application must be configured correctly in both Microsoft Entra ID and Exchange Online.

### Microsoft Entra application permission

The app registration needs the Exchange Online application permission:

- `Office 365 Exchange Online` -> `Exchange.ManageAsApp`

Admin consent must be granted for that permission.

Without this permission, the token request may succeed, but Exchange Online connection or cmdlet execution will fail.

### Exchange Online RBAC / service principal authorization

The app also needs Exchange-side authorization to run the cmdlets used by the script. In practice, that means the service principal must be granted appropriate Exchange Online role assignments.

At a minimum, the app must be allowed to run the read cmdlets used by the export script, including operations such as:

- `Get-OrganizationConfig`
- `Get-AcceptedDomain`
- `Get-EXOMailbox`
- `Get-EXORecipient`
- `Get-EXOMailboxPermission`
- `Get-EXORecipientPermission`
- `Get-EXOMailboxFolderStatistics`
- `Get-EXOMailboxFolderPermission`

For restore runs, the app also needs write access for cmdlets such as:

- `Set-Mailbox`
- `Add-MailboxPermission`
- `Add-RecipientPermission`
- `Add-MailboxFolderPermission`
- `Set-MailboxFolderPermission`

### Practical permission guidance

Many admins test this initially with a broad Exchange role assignment to confirm connectivity, then reduce scope afterward.

For least privilege, create a dedicated Exchange role group or service principal assignment that only includes the read or write operations you actually need.

Because Exchange RBAC design varies between tenants, validate the exact cmdlet coverage in your environment before relying on unattended execution.

## Repository files

### `Export-ExoPermissions-HybridReady.ps1`

Exports Exchange Online permissions to JSON.

#### Parameters

##### `-OutputPath`

- Type: `string`
- Required: No
- Default: auto-generated JSON file in the script directory
- Purpose: sets the output JSON file path

If omitted, the script creates a file like:

```text
ExoPermissions_yyyyMMdd_HHmm.json
```

##### `-UseEnvCredentials`

- Type: `switch`
- Required: No
- Default: `False`
- Purpose: tells the script to authenticate using values from the `.env` file instead of interactive sign-in

##### `-EnvFilePath`

- Type: `string`
- Required: No
- Default: `.env` in the script directory
- Purpose: custom path to the `.env` file used when `-UseEnvCredentials` is supplied

##### `-RecipientTypeDetails`

- Type: `string[]`
- Required: No
- Default:

```powershell
@(
  "UserMailbox",
  "SharedMailbox",
  "RoomMailbox",
  "EquipmentMailbox"
)
```

- Purpose: limits the mailbox types returned by `Get-EXOMailbox`

Use this if you want to target only certain mailbox classes.

##### `-SkipFolderPermissions`

- Type: `switch`
- Required: No
- Default: `False`
- Purpose: skips mailbox folder permission discovery entirely

This reduces runtime significantly if you only need mailbox-level delegations.

##### `-IncludedFolders`

- Type: `string[]`
- Required: No
- Default: empty array
- Purpose: limits folder scanning to specific folders

Supported selector styles:

- folder type name, for example `Calendar` or `Inbox`
- leaf folder name
- full folder path, for example `/Calendar` or `/Inbox/Subfolder`
- root, using `Root` or `/`

If omitted, the script scans all folders.

##### `-ThrottleDelayMs`

- Type: `int`
- Required: No
- Default: `150`
- Purpose: delay between mailboxes to reduce Exchange Online throttling pressure

##### `-MaxRetries`

- Type: `int`
- Required: No
- Default: `3`
- Purpose: maximum retry attempts for throttled or transient Exchange Online calls

The script uses exponential backoff.

##### `-SaveInterval`

- Type: `int`
- Required: No
- Default: `10`
- Purpose: controls how often incremental progress is written to disk, measured in processed mailboxes

##### `-Resume`

- Type: `switch`
- Required: No
- Default: `False`
- Purpose: resumes from the most recent `ExoPermissions_*.json` file in the script directory when `-OutputPath` is not explicitly supplied

This is useful for long-running exports.

#### Export examples

Interactive export:

```powershell
./Export-ExoPermissions-HybridReady.ps1
```

App-only export using `.env`:

```powershell
./Export-ExoPermissions-HybridReady.ps1 -UseEnvCredentials
```

Use a custom `.env` path and output file:

```powershell
./Export-ExoPermissions-HybridReady.ps1 \
  -UseEnvCredentials \
  -EnvFilePath ./secrets/source-tenant.env \
  -OutputPath ./exports/source-tenant-permissions.json
```

Mailbox-level permissions only:

```powershell
./Export-ExoPermissions-HybridReady.ps1 -SkipFolderPermissions
```

Only export user and shared mailboxes:

```powershell
./Export-ExoPermissions-HybridReady.ps1 -RecipientTypeDetails UserMailbox,SharedMailbox
```

Only scan selected folders:

```powershell
./Export-ExoPermissions-HybridReady.ps1 -IncludedFolders Calendar,Inbox,/Inbox/Projects
```

Resume a previous export:

```powershell
./Export-ExoPermissions-HybridReady.ps1 -Resume
```

### `Restore-ExoPermissions.ps1`

Restores supported permissions to the target tenant from a prior export.

#### Parameters

##### `-ExportJsonPath`

- Type: `string`
- Required: Yes
- Purpose: path to the JSON file produced by the export script

##### `-MappingCsvPath`

- Type: `string`
- Required: Yes
- Purpose: path to the source-to-target UPN mapping CSV

The CSV must contain these columns:

- `SourceUPN`
- `TargetUPN`

Additional columns are ignored.

##### `-PermTypes`

- Type: `string[]`
- Required: No
- Default: `All`
- Allowed values:
  - `X500Addresses`
  - `FullAccess`
  - `SendAs`
  - `SendOnBehalf`
  - `FolderPermissions`
  - `All`
- Purpose: limits which permission categories are restored

##### `-SkipExisting`

- Type: `bool`
- Required: No
- Default: `True`
- Purpose: skips assignments that are already present on the target mailbox

Important behavior:

- existing `FullAccess` and `SendAs` entries are skipped
- existing `SendOnBehalf` delegates are merged into the final list rather than overwritten when `SkipExisting` is enabled
- existing folder permissions are skipped rather than updated when `SkipExisting` is enabled

##### `-ThrottleDelayMs`

- Type: `int`
- Required: No
- Default: `200`
- Purpose: delay between target mailboxes to reduce throttling risk

##### `-MaxRetries`

- Type: `int`
- Required: No
- Default: `3`
- Purpose: retry count for throttled or transient Exchange Online operations

##### `-LogPath`

- Type: `string`
- Required: No
- Default: auto-generated CSV log file in the script directory
- Purpose: output path for the restore log CSV

If omitted, the script creates a file like:

```text
RestoreLog_yyyyMMdd_HHmm.csv
```

##### `-UseEnvCredentials`

- Type: `switch`
- Required: No
- Default: `False`
- Purpose: tells the script to authenticate by using values from the `.env` file instead of interactive sign-in

##### `-EnvFilePath`

- Type: `string`
- Required: No
- Default: `.env` in the script directory
- Purpose: custom path to the `.env` file used when `-UseEnvCredentials` is supplied

#### Restore examples

The Quick Start section covers the default dry-run workflow. The examples below focus on common variations.

Restore only X500 and mailbox delegations:

```powershell
./Restore-ExoPermissions.ps1 \
  -ExportJsonPath ./ExoPermissions_20260309_1015.json \
  -MappingCsvPath ./MailboxMap.csv \
  -PermTypes X500Addresses,FullAccess,SendAs,SendOnBehalf
```

Restore everything and write a custom log:

```powershell
./Restore-ExoPermissions.ps1 \
  -ExportJsonPath ./ExoPermissions_20260309_1015.json \
  -MappingCsvPath ./MailboxMap.csv \
  -LogPath ./logs/restore-run-01.csv
```

Force updates instead of skipping existing folder permissions:

```powershell
./Restore-ExoPermissions.ps1 \
  -ExportJsonPath ./ExoPermissions_20260309_1015.json \
  -MappingCsvPath ./MailboxMap.csv \
  -SkipExisting $false
```

## Input file formats

### Export JSON

The export JSON contains two top-level sections:

- `ExportMetadata`
- `Mailboxes`

High-level shape:

```json
{
  "ExportMetadata": {
    "ExportedAt": "2026-03-09T18:20:00Z",
    "ScriptVersion": "2.8",
    "SourceTenantDomain": "contoso.onmicrosoft.com",
    "SourceTenantId": "...",
    "TotalMailboxes": 100,
    "TotalMailboxesExported": 100,
    "TotalIdentityCacheSize": 500,
    "TotalResolveErrors": 3,
    "TotalFolderErrors": 1,
    "RecipientTypesIncluded": ["UserMailbox", "SharedMailbox"],
    "IncludedFolders": [],
    "FolderPermissionsCollected": true
  },
  "Mailboxes": [
    {
      "SourceUser": {},
      "MailboxDelegations": {
        "FullAccess": [],
        "SendAs": [],
        "SendOnBehalf": []
      },
      "FolderPermissions": [],
      "UnresolvablePermissions": [],
      "ExportErrors": []
    }
  ]
}
```

During incremental saves, the export file may also include:

- `IsPartial`

That flag indicates the JSON was saved before the run completed.

### Mapping CSV

Minimum required format:

```csv
SourceUPN,TargetUPN
alice@source.com,alice@target.com
bob@source.com,bob@target.com
```

Behavior:

- rows with blank source or target values are ignored
- mailbox records whose `SourceUPN` is not found in the mapping CSV are skipped during restore
- delegates whose exported `DelegateUPN` is not found in the mapping CSV are also skipped during restore

## Logging and output files

### Export output

The export script writes a JSON file containing:

- export metadata
- mailbox-level delegations
- folder-level permissions
- unresolvable trustees
- per-mailbox export errors

It also performs incremental saves every `SaveInterval` mailboxes.

### Restore log

The restore script writes a CSV log with these columns:

- `Timestamp`
- `TargetMailbox`
- `PermType`
- `Delegate`
- `Detail`
- `Status`
- `Message`

Possible `Status` values:

- `Applied`
- `Skipped`
- `WhatIf`
- `Error`
- `Warning`
- `Info`

## How unresolvable permissions are handled

The export script intentionally does not fail the full run when a trustee cannot be resolved.

Instead, it records those entries under `UnresolvablePermissions` with a reason such as:

- `OrphanedSID`
- `RecipientNotFound`
- `CachedNotFound`

This is especially useful in hybrid or long-lived tenants where stale ACL entries are common.

During restore, unresolved delegates are skipped and written to the restore log.

## Throttling and retry behavior

Both scripts use retry logic for Exchange Online operations.

Behavior:

- retries on errors matching throttling or transient patterns such as `429`, `TooManyRequests`, `ServiceUnavailable`, or `throttl`
- uses exponential backoff
- applies a configurable delay between mailbox operations

For larger tenants, increase `ThrottleDelayMs` if you see sustained throttling.

## Recommended usage workflow

1. Install or update `ExchangeOnlineManagement`.
2. Prepare authentication for the source tenant:
   - interactive sign-in, or
   - `.env` app-only authentication
3. Run the export script against the source tenant.
4. Review the JSON output for scope, unresolvable entries, and export errors.
5. Build a `SourceUPN` to `TargetUPN` mapping CSV.
6. Prepare authentication for the target tenant:
   - interactive sign-in, or
   - `.env` app-only authentication
7. Run the restore script with `-WhatIf` against the target tenant.
8. Review the restore log.
9. Run the restore again without `-WhatIf`.

## Known limitations and implementation notes

- Both scripts support interactive Exchange Online sign-in and `.env` app-only authentication.
- The scripts rely on Exchange Online cmdlet behavior and RBAC in your tenant; some cmdlets may behave differently depending on service changes or role assignments.
- Folder permissions are restored by using the folder path recorded in the export. If the corresponding folder structure does not exist in the target mailbox, those entries can fail.
- Delegates not present in the mapping CSV are skipped by design.
- Exported and restored data covers Exchange Online permissions only. It does not migrate mailbox content.

## Publishing checklist

Before publishing this repository to GitHub:

- verify `.gitignore` excludes `.env`, export JSON files, and restore logs
- confirm no tenant-specific sample files remain in the repository
- test the examples in this README against the current script versions
- review `.env.example` and `MailboxMap.example.csv` before publishing so the placeholders match your preferred naming conventions

## License

Add the license of your choice before publishing if you want the repository to be reusable by others.
