# M365-SPO-OD-Teams Interactive Backup

Backs up files from SharePoint Online document libraries, Teams channel documents, and OneDrive for Business to a local directory. Supports incremental backup, automatic throttling, and dry-run mode.

## Prerequisites

### PowerShell Modules

Install the required Microsoft Graph PowerShell SDK modules:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Files        -Scope CurrentUser
Install-Module Microsoft.Graph.Sites        -Scope CurrentUser
```

### Permissions

The script authenticates interactively via `Connect-MgGraph` and requests these scopes:

| Scope | Purpose |
|---|---|
| `Files.Read.All` | Read files from any drive (SPO libraries, OneDrive) |
| `Sites.Read.All` | Look up SharePoint sites and their document libraries |
| `User.Read` | Read the signed-in user's profile (required for OneDrive `me/drive` access) |

An Azure AD admin may need to grant consent for these scopes in your tenant.

### Environment

- PowerShell 5.1 or PowerShell 7+
- Network access to Microsoft Graph (`graph.microsoft.com`)

## Quick Start

1. Place `config.json` in the same directory as the script (see Configuration below).
2. Run the script:

```powershell
# Preview what would be downloaded (no files written)
.\Backup-M365-Interactive.ps1 -DryRun

# Run the actual backup (default: RenameNew mode — originals are preserved)
.\Backup-M365-Interactive.ps1

# Run backup with overwrite mode (replaces changed files in place)
.\Backup-M365-Interactive.ps1 -UpdateAction Overwrite
```

3. A browser window will open for authentication on the first run.

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-DryRun` | off | Preview what would be downloaded without writing any files. |
| `-UpdateAction` | `RenameNew` | Global update mode. `RenameNew` renames the existing local file with a `_prev_XXXXX` suffix and downloads the new version to the original filename (preserving the old copy while keeping incremental comparison working). `Overwrite` replaces existing files in place. Can be overridden per-task via the `UpdateAction` field in `config.json`. |

## Configuration

Create a `config.json` file in the script directory. It is a JSON array of task objects.

### SharePoint / Teams Task

```json
[
  {
    "Type": "SharePoint",
    "SiteName": "Marketing",
    "LibraryName": "Shared Documents",
    "SourcePath": "/",
    "LocalDownloadPath": "C:\\Backups\\Marketing",
    "UpdateAction": "Overwrite"
  }
]

> **Note:** The `UpdateAction` field in `config.json` is optional and overrides the global `-UpdateAction` parameter for that specific task. If omitted, the task inherits the global value (default: `RenameNew`).
```

| Field | Required | Description |
|---|---|---|
| `Type` | Yes | `"SharePoint"` |
| `SiteName` | Yes* | Display name of the site. The script searches for a matching site and prefers an exact DisplayName match. |
| `SiteUrl` | Yes* | Site identifier. Accepts full URLs (`https://contoso.sharepoint.com/sites/Marketing`), short format (`contoso.sharepoint.com:/sites/Marketing`), or without protocol (`contoso.sharepoint.com/sites/Marketing`). All formats are auto-converted to the Graph API format. Use this instead of `SiteName` to avoid ambiguous search results. |
| `LibraryName` | Yes | Name of the document library. Common values: `"Shared Documents"`, `"Documents"`. The script handles the alias between these two and also tries URL-decoded matching. |
| `SourcePath` | No | Subfolder path within the library to back up. Use `"/"` or omit for the library root. Do **not** include the library name in this path. |
| `LocalDownloadPath` | Yes | Local directory where files will be saved. Created automatically if it does not exist. |
| `UpdateAction` | No | Per-task override for the global `-UpdateAction` parameter. `"Overwrite"` replaces existing files. `"RenameNew"` renames the existing file with a `_prev_XXXXX` suffix and downloads the new version to the original filename. If omitted, inherits the global value (default: `RenameNew`). |

\* Provide either `SiteName` or `SiteUrl`. `SiteUrl` is recommended for reliability.

### Backing Up Teams Channel Documents

Teams stores channel files in a SharePoint site. Use the `"SharePoint"` type and point it at the Team's site:

```json
{
  "Type": "SharePoint",
  "SiteUrl": "contoso.sharepoint.com:/sites/SalesTeam",
  "LibraryName": "Shared Documents",
  "SourcePath": "/General",
  "LocalDownloadPath": "C:\\Backups\\SalesTeam-General"
}
```

- Each Teams channel has a folder inside `Shared Documents` named after the channel (e.g. `/General`, `/Project Alpha`).
- To back up all channels at once, set `SourcePath` to `"/"`.
- To find your Team's SharePoint URL: open the channel in Teams, click **Open in SharePoint** from the Files tab, and note the site URL.

### OneDrive Task

```json
{
  "Type": "OneDrive",
  "TargetUser": "jane.doe@contoso.com",
  "SourcePath": "/",
  "LocalDownloadPath": "C:\\Backups\\JaneDoe-OneDrive"
}
```

| Field | Required | Description |
|---|---|---|
| `Type` | Yes | `"OneDrive"` |
| `TargetUser` | No | UPN of the user whose OneDrive to back up (e.g. `jane.doe@contoso.com`). If omitted, the script uses the signed-in user's OneDrive. Required for app-only or admin-on-behalf scenarios. |
| `SourcePath` | No | Subfolder path within OneDrive. `"/"` or omit for the root. |
| `LocalDownloadPath` | Yes | Local directory where files will be saved. |
| `UpdateAction` | No | Same as SharePoint tasks. |

### Multi-Task Example

```json
[
  {
    "Type": "SharePoint",
    "SiteUrl": "contoso.sharepoint.com:/sites/Engineering",
    "LibraryName": "Shared Documents",
    "SourcePath": "/",
    "LocalDownloadPath": "C:\\Backups\\Engineering"
  },
  {
    "Type": "SharePoint",
    "SiteName": "HR Portal",
    "LibraryName": "Policies",
    "SourcePath": "/2024",
    "LocalDownloadPath": "C:\\Backups\\HR-Policies-2024"
  },
  {
    "Type": "OneDrive",
    "TargetUser": "admin@contoso.com",
    "SourcePath": "/Projects",
    "LocalDownloadPath": "C:\\Backups\\Admin-Projects"
  }
]
```

## Incremental Backup Logic

The script skips files that have not changed since the last backup:

1. **SHA256 / SHA1 hash** — Used when the remote file provides a SHA hash (typical for OneDrive personal). The local file hash is computed and compared.
2. **Size + last-modified date** — Used when only `quickXorHash` is available (typical for SharePoint and Teams files). Compares file size and `lastModifiedDateTime` with a 2-second tolerance.
3. **Size only** — Last resort when no hash or date is available from the API.

After downloading a file, the script sets the local file's last-modified timestamp to match the remote value so that date-based comparison works correctly on subsequent runs.

## Output Files

All output files are created in the script directory:

| File | Description |
|---|---|
| `script_log_YYYYMMDD-HHmmss.txt` | Full log for each run. A new file is created per execution. |
| `renamed_files_manifest.csv` | Tracks files that were renamed due to illegal characters in the original filename. Maps original name to safe name with the item ID for traceability. |

## Dry Run

Use `-DryRun` to preview the backup without writing any files:

```powershell
.\Backup-M365-Interactive.ps1 -DryRun
.\Backup-M365-Interactive.ps1 -DryRun -UpdateAction Overwrite
```

This logs what would be downloaded, skipped, or updated without making any changes to disk.

## GUI

Launch `Backup-GUI.ps1` for a graphical interface. The toolbar includes an **Update Mode** dropdown (`RenameNew` / `Overwrite`) that sets the global `-UpdateAction` parameter for the backup process. The default selection is `RenameNew`.

## Throttling and Error Handling

- **429 / 503 / 504 responses** are automatically retried with exponential backoff (up to 10 retries). The `Retry-After` header is respected when present.
- **Other errors** (404, 401, etc.) are not retried. The task logs the error and continues to the next file or task.
- **OneNote notebooks** and other package-type items in document libraries are detected and skipped with a warning (they cannot be downloaded as regular files via the Graph API).

## Troubleshooting

### "Library not found" on a Teams site

The log will show all available drives on the site. Check the exact `Name` value and use it as `LibraryName` in your config. For most Teams sites, the library is named `"Shared Documents"` or `"Documents"`.

### Wrong site matched

If `SiteName` is ambiguous (e.g. "Sales" matching "Sales", "Pre-Sales", "Sales Reports"), the log will warn and show which site was selected. Switch to `SiteUrl` for exact matching:

```json
"SiteUrl": "contoso.sharepoint.com:/sites/Sales"
```

### Config validation errors

Missing or invalid fields are reported at the start of each task with a clear message. The task is skipped and the script continues with the remaining tasks.

### Finding the SiteUrl value

1. Open the SharePoint site or Teams channel **Files** tab in a browser.
2. Copy the URL from the address bar — you can use it directly:

```json
"SiteUrl": "https://contoso.sharepoint.com/sites/MySiteName"
```

The script automatically converts any of these formats to the Graph API format:

| Input format | Example |
|---|---|
| Full URL | `https://contoso.sharepoint.com/sites/Sales` |
| Without protocol | `contoso.sharepoint.com/sites/Sales` |
| Graph format (colon) | `contoso.sharepoint.com:/sites/Sales` |

All three are equivalent. The script normalizes them internally to `contoso.sharepoint.com:/sites/Sales` (the format required by the Microsoft Graph API).
