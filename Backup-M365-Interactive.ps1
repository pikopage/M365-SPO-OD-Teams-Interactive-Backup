<#
.SYNOPSIS
    M365-SPO-OD-Teams Interactive Backup - backs up SharePoint Online, OneDrive, and Teams content.

.DESCRIPTION
    This script connects to Microsoft Graph to recursively download files from specified
    SharePoint Online libraries and OneDrive folders.
    It features:
    - Incremental backup (skips files with matching Hash, or Size+Date fallback).
    - Robust error handling and throttling management (429/503 retries).
    - Detailed logging to console and file.
    - Configuration via an external 'config.json' file.
    - Configurable update mode: RenameNew (default) preserves originals, Overwrite replaces in-place.

.PARAMETER DryRun
    Preview what would be downloaded without writing any files to disk.

.PARAMETER UpdateAction
    Global update mode applied to all tasks unless overridden per-task in config.json.
    - RenameNew  (default) — When a newer version is detected, the existing local file
                 is renamed with a _prev_XXXXX suffix, and the new version is downloaded
                 to the original filename. This preserves the old copy while keeping
                 the canonical path up to date for incremental comparison.
    - Overwrite  — Replaces the existing local file with the newer version.

.EXAMPLE
    .\Backup-M365-Interactive.ps1
    Runs backup with the default RenameNew mode.

.EXAMPLE
    .\Backup-M365-Interactive.ps1 -UpdateAction Overwrite
    Runs backup replacing changed files in place.

.EXAMPLE
    .\Backup-M365-Interactive.ps1 -DryRun
    Preview-only run (no files written) using the default RenameNew mode.

.NOTES
    Requires 'Microsoft.Graph.Files' and 'Microsoft.Graph.Sites' modules.
    Ensure 'config.json' is present in the script directory.
#>

param(
    [switch]$DryRun,

    [ValidateSet("RenameNew", "Overwrite")]
    [string]$UpdateAction = "RenameNew"
)

# --- 1. Setup Logging & Configuration ---
$LogFile = Join-Path $PSScriptRoot ("script_log_{0}.txt" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
$ConfigPath = Join-Path $PSScriptRoot "config.json"
$ManifestPath = Join-Path $PSScriptRoot "renamed_files_manifest.csv"

# Initialize Manifest if missing
if (-not (Test-Path $ManifestPath)) {
    '"Timestamp","OriginalName","NewName","ItemId","DriveId"' | Out-File -FilePath $ManifestPath -Encoding UTF8
}

# Helper Function: Write-Log
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [ConsoleColor]$Color = "White"
    )

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogLine = "[$Timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogLine -Force -Encoding UTF8
    Write-Host "[$Timestamp] [$Level] $Message" -ForegroundColor $Color
}

# --- 1.1 Config Validation (FIX #4: fail-fast on missing fields) ---
function Test-TaskConfig {
    param ([object]$Task, [int]$Index)

    $Errors = @()

    if (-not $Task.Type) { $Errors += "Missing 'Type'" }
    if (-not $Task.LocalDownloadPath) { $Errors += "Missing 'LocalDownloadPath'" }

    if ($Task.Type -eq "SharePoint") {
        if (-not $Task.SiteName -and -not $Task.SiteUrl) { $Errors += "Missing 'SiteName' or 'SiteUrl'" }
        if (-not $Task.LibraryName) { $Errors += "Missing 'LibraryName'" }
    }
    elseif ($Task.Type -eq "OneDrive") {
        # Valid with just Type + LocalDownloadPath (TargetUser is optional)
    }
    elseif ($Task.Type) {
        $Errors += "Unknown Type '$($Task.Type)'. Valid types: SharePoint, OneDrive"
    }

    if ($Errors.Count -gt 0) {
        $ErrorList = $Errors -join "; "
        Write-Log "Task #$Index config validation failed: $ErrorList" "ERROR" "Red"
        return $false
    }
    return $true
}

# --- 1.2 Normalize site URL to Graph API format ---
function ConvertTo-GraphSiteId {
    param ([string]$Url)
    $Url = $Url.Trim().TrimEnd('/')
    # Already in Graph format (hostname:/sites/path) - return as-is
    if ($Url -match '^[^/:]+\.(sharepoint\.com|sharepoint\.us):/.+') { return $Url }
    # Full URL: https://hostname/sites/path or https://hostname/teams/path
    if ($Url -match '^https?://([^/]+)(/(?:sites|teams)/.+)$') {
        return "$($Matches[1]):$($Matches[2])"
    }
    # hostname/sites/path (no protocol, no colon)
    if ($Url -match '^([^/:]+\.(sharepoint\.com|sharepoint\.us))(/(?:sites|teams)/.+)$') {
        return "$($Matches[1]):$($Matches[3])"
    }
    # Unrecognized format - return as-is and let Graph API report the error
    return $Url
}

# --- 2. Throttling Helper Function ---
# This function wraps any command and retries if it hits a 429 (Throttling) or 503/504 (Server Busy) error.
function Invoke-WithRetry {
    param (
        [ScriptBlock]$Command,
        [int]$MaxRetries = 10
    )

    $RetryCount = 0
    $Completed = $false
    $Result = $null

    while (-not $Completed) {
        try {
            # Attempt to run the command
            $Result = & $Command
            $Completed = $true
        }
        catch {
            # Check if the error is related to Throttling (429) or Server issues (5xx)
            $ex = $_.Exception
            $StatusCode = $null

            # Try to pull status (varies by SDK version)
            if ($null -ne $ex.Data -and $ex.Data.Contains("StatusCode")) { $StatusCode = $ex.Data["StatusCode"] }
            elseif ($ex.PSObject.Properties.Match('ResponseStatusCode').Count) { $StatusCode = $ex.ResponseStatusCode }
            elseif ($ex.PSObject.Properties.Match('Response').Count -and $ex.Response -and $ex.Response.StatusCode) { $StatusCode = $ex.Response.StatusCode }

            # Some SDK errors wrap the status code in the message or inner exception
            if ($null -eq $StatusCode -and $ex.Message -match "429|TooManyRequests") { $StatusCode = 429 }

            if ($StatusCode -eq 429 -or $StatusCode -eq 503 -or $StatusCode -eq 504) {
                $RetryCount++
                if ($RetryCount -gt $MaxRetries) {
                    Write-Log "Max retries ($MaxRetries) reached for error: $($ex.Message)" "ERROR" "Red"
                    throw
                }

                # Default wait time (in case Retry-After header is missing)
                $WaitSeconds = 10 * $RetryCount

                # Try to read 'Retry-After' header if available
                $headers = $null
                if ($ex.ResponseHeaders) { $headers = $ex.ResponseHeaders }
                elseif ($ex.HttpResponse -and $ex.HttpResponse.Headers) { $headers = $ex.HttpResponse.Headers }
                elseif ($ex.Response -and $ex.Response.Headers) { $headers = $ex.Response.Headers }

                if ($headers) {
                    if ($headers.ContainsKey("Retry-After")) {
                        $WaitSeconds = [int]($headers["Retry-After"] | Select-Object -First 1)
                    } elseif ($headers.ContainsKey("retry-after")) {
                        $WaitSeconds = [int]($headers["retry-after"] | Select-Object -First 1)
                    }
                }

                Write-Log "Throttling detected (429/503). Waiting $WaitSeconds seconds before retry #$RetryCount..." "WARN" "Yellow"
                Start-Sleep -Seconds $WaitSeconds
            }
            else {
                # If it's a different error (e.g., 404 Not Found, 401 Access Denied), do not retry.
                throw $_
            }
        }
    }
    return $Result
}

# --- 3. Load Config & Connect ---
# Ensure console can display Czech/Unicode characters correctly
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Log "========================================" "INFO" "Magenta"
Write-Log "M365-SPO-OD-Teams Interactive Backup starting" "INFO" "Magenta"
Write-Log "Global Update Action: $UpdateAction" "INFO" "Cyan"
if ($DryRun) {
    Write-Log "!!! DRY RUN MODE ACTIVE - No files will be downloaded !!!" "WARN" "Yellow"
}

if (!(Test-Path $ConfigPath)) {
    Write-Log "Config file not found at: $ConfigPath" "ERROR" "Red"
    return
}

try {
    $TaskList = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
} catch {
    Write-Log "Error reading config.json: $_" "ERROR" "Red"
    return
}

try {
    Connect-MgGraph -Scopes "Files.Read.All", "Sites.Read.All", "User.Read" -NoWelcome
    $ctx = Get-MgContext
    if (-not $ctx -or [string]::IsNullOrEmpty($ctx.Account)) {
        Write-Log "[AUTH-FAILED] Connect-MgGraph returned no valid context — token may have been cancelled or consent denied." "ERROR" "Red"
        return
    }
    Write-Log "Connected to Microsoft Graph successfully." "INFO" "Green"
    Write-Log "[AUTH] $($ctx.Account)  |  tenant: $($ctx.TenantId)" "INFO" "Cyan"
}
catch {
    Write-Log "[AUTH-FAILED] Failed to connect to Microsoft Graph. Error: $_" "ERROR" "Red"
    return
}

# --- 3.1 Initialize Statistics (FIX #7: per-task + global stats) ---
$Script:TotalStats = @{ Downloaded = 0; Skipped = 0; Errors = 0 }
$Script:TaskStats = $null

# --- 4. Recursive Function ---
function Copy-FolderRecursively {
    param (
        [string]$DriveId,
        [string]$FolderItemId,
        [string]$LocalBasePath,
        [string]$UpdateAction = "Overwrite",
        [switch]$DryRun
    )

    if (!(Test-Path $LocalBasePath)) {
        if (-not $DryRun) {
            New-Item -ItemType Directory -Path $LocalBasePath -Force | Out-Null
        }
    }

    $NextLink = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$FolderItemId/children?`$top=200"

    do {
        # --- WRAPPED CALL 1: Getting File List ---
        try {
            $ApiResponse = Invoke-WithRetry -Command {
                return Invoke-MgGraphRequest -Method GET -Uri $NextLink
            }
            $Children = $ApiResponse.value
        }
        catch {
            Write-Log "Error retrieving folder contents for '$LocalBasePath'. Error: $_" "ERROR" "Red"
            $Script:TaskStats.Errors++
            break
        }

        foreach ($Item in $Children) {
            # FIX #5: Skip OneNote notebooks and other package-type items (cannot be downloaded as files)
            if ($null -ne $Item.package) {
                Write-Log "Skipping package item (e.g. OneNote): $($Item.name)" "WARN" "Yellow"
                $Script:TaskStats.Skipped++
                continue
            }

            $SafeName = $Item.name -replace '[\\/*?:"<>|]', '_'

            # Prevent data loss due to filename collisions (e.g. "Doc:1.txt" vs "Doc_1.txt")
            if ($SafeName -ne $Item.name) {
                # FIX #6: Use first 8 chars of item ID instead of full ID to keep paths short
                $IdSuffix = $Item.id.Substring(0, [Math]::Min(8, $Item.id.Length))
                $SafeName = "{0}_{1}" -f $SafeName, $IdSuffix

                # Log change to manifest
                $CsvLine = '"{0}","{1}","{2}","{3}","{4}"' -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), ($Item.name -replace '"','""'), $SafeName, $Item.id, $DriveId
                if (-not $DryRun) {
                    Add-Content -Path $ManifestPath -Value $CsvLine -Encoding UTF8
                }
            }

            $LocalItemPath = Join-Path -Path $LocalBasePath -ChildPath $SafeName

            if ($null -ne $Item.folder) {
                Write-Log "Folder: $SafeName" "INFO" "Yellow"
                Copy-FolderRecursively -DriveId $DriveId -FolderItemId $Item.id -LocalBasePath $LocalItemPath -UpdateAction $UpdateAction -DryRun:$DryRun
            }
            else {
                $ShouldDownload = $true
                $LocalFileExists = Test-Path $LocalItemPath

                if ($LocalFileExists) {
                    $LocalFile = Get-Item $LocalItemPath
                    $RemoteHash = $null
                    $LocalHashAlgo = $null

                    # Check for available hashes (SHA256 or SHA1)
                    if ($Item.file.hashes.sha256Hash) {
                        $RemoteHash = $Item.file.hashes.sha256Hash
                        $LocalHashAlgo = "SHA256"
                    }
                    elseif ($Item.file.hashes.sha1Hash) {
                        $RemoteHash = $Item.file.hashes.sha1Hash
                        $LocalHashAlgo = "SHA1"
                    }

                    if ($RemoteHash) {
                        $LocalHash = (Get-FileHash -Path $LocalItemPath -Algorithm $LocalHashAlgo).Hash
                        if ($LocalHash -eq $RemoteHash) {
                            Write-Log "Skipping: $SafeName (Hash Match)" "INFO" "DarkGray"
                            $ShouldDownload = $false
                            $Script:TaskStats.Skipped++
                        }
                        else {
                            Write-Log "Updating: $SafeName (Hash Mismatch)" "INFO" "Yellow"
                        }
                    }
                    else {
                        # FIX #1: Fallback — compare Size AND LastModified when no SHA hash is available.
                        # SPO/Teams files typically only provide quickXorHash which can't be computed locally,
                        # so size-only comparison was unreliable (same-size modified files were silently skipped).
                        $RemoteModified = $null
                        if ($Item.lastModifiedDateTime) {
                            $RemoteModified = ([DateTimeOffset]::Parse(
                                $Item.lastModifiedDateTime,
                                [System.Globalization.CultureInfo]::InvariantCulture)).UtcDateTime
                        }
                        $LocalModified = $LocalFile.LastWriteTimeUtc

                        $SizeMatch = $LocalFile.Length -eq $Item.size
                        # Allow 2-second tolerance for filesystem timestamp rounding
                        $DateMatch = $RemoteModified -and ([Math]::Abs(($LocalModified - $RemoteModified).TotalSeconds) -lt 2)

                        if ($SizeMatch -and $DateMatch) {
                            Write-Log "Skipping: $SafeName (Size+Date Match)" "INFO" "DarkGray"
                            $ShouldDownload = $false
                            $Script:TaskStats.Skipped++
                        }
                        elseif ($SizeMatch -and -not $RemoteModified) {
                            Write-Log "Skipping: $SafeName (Size Match - No Hash/Date Available)" "INFO" "DarkGray"
                            $ShouldDownload = $false
                            $Script:TaskStats.Skipped++
                        }
                        else {
                            $Reason = if (-not $SizeMatch) { "Size Mismatch" } else { "Date Mismatch" }
                            Write-Log "Updating: $SafeName ($Reason)" "INFO" "Yellow"
                        }
                    }
                }

                if ($ShouldDownload) {
                    if ($DryRun) {
                        if ($LocalFileExists -and $UpdateAction -eq "RenameNew") {
                            $Extension = [System.IO.Path]::GetExtension($SafeName)
                            $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($SafeName)
                            $PrevName = "{0}_prev_XXXXX{1}" -f $BaseName, $Extension
                            Write-Log "[DRYRUN] Would rename existing to: $PrevName and download new: $SafeName" "INFO" "Cyan"
                        } else {
                            Write-Log "[DRYRUN] Would download/update: $SafeName" "INFO" "Cyan"
                        }
                        $Script:TaskStats.Downloaded++
                        continue
                    }

                    try {
                        if (!(Test-Path (Split-Path $LocalItemPath))) {
                            New-Item -ItemType Directory -Path (Split-Path $LocalItemPath) -Force | Out-Null
                        }

                        $TargetFilePath = $LocalItemPath

                        # Handle Update Action (RenameNew) - Rename existing file, download new to original path
                        # This ensures the canonical filename always holds the latest version,
                        # so subsequent runs compare against the current remote and skip if unchanged.
                        if ($LocalFileExists -and $UpdateAction -eq "RenameNew") {
                            $Extension = [System.IO.Path]::GetExtension($SafeName)
                            $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($SafeName)
                            $Suffix = Get-Random -Minimum 10000 -Maximum 99999
                            $PrevName = "{0}_prev_{1}{2}" -f $BaseName, $Suffix, $Extension
                            $PrevPath = Join-Path -Path $LocalBasePath -ChildPath $PrevName
                            Move-Item -Path $LocalItemPath -Destination $PrevPath -Force
                            Write-Log "Preserved original as: $PrevName" "INFO" "Yellow"
                            # TargetFilePath stays as $LocalItemPath (the canonical filename)
                        }

                        # --- WRAPPED CALL 2: Downloading File ---
                        Invoke-WithRetry -Command {
                            Get-MgDriveItemContent -DriveId $DriveId -DriveItemId $Item.id -OutFile $TargetFilePath
                        }

                        # FIX #3b: Set local file timestamp to match remote so date-based
                        # incremental comparison works correctly on the next run.
                        if ($Item.lastModifiedDateTime) {
                            $RemoteTime = ([DateTimeOffset]::Parse(
                                $Item.lastModifiedDateTime,
                                [System.Globalization.CultureInfo]::InvariantCulture)).UtcDateTime
                            (Get-Item $TargetFilePath).LastWriteTimeUtc = $RemoteTime
                        }

                        Write-Log "Downloaded: $(Split-Path $TargetFilePath -Leaf)" "INFO" "Cyan"
                        $Script:TaskStats.Downloaded++
                    }
                    catch {
                        Write-Log "Failed to download: $($Item.name). Error: $_" "WARN" "Red"
                        $Script:TaskStats.Errors++
                    }
                }
            }
        }
        $NextLink = $ApiResponse.'@odata.nextLink'

    } while ($NextLink)
}

# --- 5. Main Batch Loop ---
Write-Log "Found $( $TaskList.Count ) tasks to process." "INFO" "Magenta"
$TaskIndex = 0

foreach ($Task in $TaskList) {
    $TaskIndex++
    $CurrentAction = if ($Task.UpdateAction) { $Task.UpdateAction } else { $UpdateAction }
    Write-Log "----------------------------------------"
    Write-Log "Starting Task #$TaskIndex : $($Task.Type) -> $($Task.LocalDownloadPath) [Mode: $CurrentAction]" "INFO" "White"

    # FIX #4: Validate config before doing any work
    if (-not (Test-TaskConfig -Task $Task -Index $TaskIndex)) {
        Write-Log "Skipping task #$TaskIndex due to config errors." "ERROR" "Red"
        $Script:TotalStats.Errors++
        continue
    }

    # FIX #7: Per-task stats
    $Script:TaskStats = @{ Downloaded = 0; Skipped = 0; Errors = 0 }
    $TargetDrive = $null

    try {
        if ($Task.Type -eq "SharePoint") {
            $Site = $null

            # FIX #2: Prefer exact lookup via SiteUrl if provided (e.g. "contoso.sharepoint.com:/sites/TeamSite")
            if ($Task.SiteUrl) {
                try {
                    $GraphSiteId = ConvertTo-GraphSiteId $Task.SiteUrl
                    $Site = Invoke-WithRetry -Command { Get-MgSite -SiteId $GraphSiteId }
                    Write-Log "Resolved site via SiteUrl: $($Site.DisplayName) (Id: $($Site.Id), Url: $($Site.WebUrl))" "INFO" "Cyan"
                }
                catch {
                    throw "Could not resolve SiteUrl '$($Task.SiteUrl)'. Ensure format is 'hostname:/sites/sitepath' or a full URL. Error: $_"
                }
            }
            else {
                # FIX #2: Fuzzy search — prefer exact DisplayName match, then fall back to first result
                $SearchResults = Invoke-WithRetry -Command { Get-MgSite -Search $Task.SiteName }
                if (-not $SearchResults) { throw "No sites found matching '$($Task.SiteName)'." }

                $Site = $SearchResults | Where-Object { $_.DisplayName -eq $Task.SiteName } | Select-Object -First 1
                if (-not $Site) {
                    $Site = $SearchResults | Select-Object -First 1
                    $AvailableNames = ($SearchResults | ForEach-Object { "'$($_.DisplayName)'" }) -join ", "
                    Write-Log "No exact DisplayName match for '$($Task.SiteName)'. Search returned: $AvailableNames. Using first result." "WARN" "Yellow"
                }

                # FIX #8: Always log which site was matched
                Write-Log "Matched site: $($Site.DisplayName) (Id: $($Site.Id), Url: $($Site.WebUrl))" "INFO" "Cyan"
            }

            $AllDrives = Invoke-WithRetry -Command { Get-MgSiteDrive -SiteId $Site.Id }
            # Try to match Library Name, "Shared Documents" alias, or WebUrl suffix (with URL-decoding)
            $TargetDrive = $AllDrives | Where-Object {
                $_.Name -eq $Task.LibraryName -or
                ($Task.LibraryName -eq "Documents" -and $_.Name -eq "Shared Documents") -or
                ($Task.LibraryName -eq "Shared Documents" -and $_.Name -eq "Documents") -or
                $_.WebUrl -like "*/$($Task.LibraryName)" -or
                ([Uri]::UnescapeDataString($_.WebUrl) -like "*/$($Task.LibraryName)")
            } | Select-Object -First 1
            if (-not $TargetDrive) {
                $DriveList = ($AllDrives | ForEach-Object { "  - Name: '$($_.Name)' | WebUrl: $($_.WebUrl)" }) -join "`n"
                Write-Log "Available drives on site '$($Site.DisplayName)':`n$DriveList" "WARN" "Yellow"
                throw "Library '$($Task.LibraryName)' not found on site '$($Site.DisplayName)'. Check the drive names listed above."
            }
        }
        elseif ($Task.Type -eq "OneDrive") {
            if ($Task.TargetUser) {
                try {
                    $TargetDrive = Invoke-WithRetry -Command { Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($Task.TargetUser)/drive" -ErrorAction Stop }
                    Write-Log "Found OneDrive for user: $($Task.TargetUser)" "INFO" "Cyan"
                }
                catch {
                    throw "Could not find OneDrive for TargetUser '$($Task.TargetUser)'. Error: $_"
                }
            }
            else {
                try {
                    $TargetDrive = Invoke-WithRetry -Command { Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me/drive" -ErrorAction Stop }
                }
                catch {
                    Write-Log "Could not access 'me/drive'. This usually happens in App-Only context or if the user lacks a license." "WARN" "Yellow"

                    Write-Log "Listing all available drives (Note: Personal drives may not appear here)..." "WARN" "Yellow"
                    $AllDrives = Invoke-WithRetry -Command { Get-MgDrive -All }
                    $TargetDrive = $AllDrives | Where-Object { $_.DriveType -in @("personal", "business") } | Select-Object -First 1

                    if (-not $TargetDrive -and $AllDrives.Count -gt 0) {
                        Write-Log "No Personal/Business drive found in available list. Found $($AllDrives.Count) drives (likely SharePoint libraries)." "WARN" "Yellow"
                        Write-Log "TIP: If running as an Application, add 'TargetUser': 'user@domain.com' to your config.json for OneDrive tasks." "WARN" "Yellow"

                        $TargetDrive = $AllDrives | Select-Object -First 1
                        Write-Log "Falling back to first available drive: $($TargetDrive.Id) (Name: $($TargetDrive.Name) | Url: $($TargetDrive.WebUrl))" "WARN" "Yellow"
                    }
                }
            }
            if (-not $TargetDrive) { throw "No OneDrive found. Please specify 'TargetUser' in config or ensure signed-in user has a OneDrive." }
        }

        $CleanPath = if ($Task.SourcePath) { $Task.SourcePath.TrimStart('/') } else { "" }

        if ([string]::IsNullOrWhiteSpace($CleanPath)) {
             $ApiUrl = "https://graph.microsoft.com/v1.0/drives/$($TargetDrive.Id)/root"
        }
        else {
             # Encode path segments to handle spaces/special characters
             $EncodedPath = ($CleanPath.Split('/') | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
             $ApiUrl = "https://graph.microsoft.com/v1.0/drives/$($TargetDrive.Id)/root:/$EncodedPath"
        }

        # --- WRAPPED CALL 3: Initial Folder Lookup ---
        try {
            $StartFolder = Invoke-WithRetry -Command {
                return Invoke-MgGraphRequest -Method GET -Uri $ApiUrl
            }
        }
        catch {
            $StatusCode = $null
            $ex = $_.Exception
            if ($ex.Message -match "403|Forbidden") { $StatusCode = 403 }
            elseif ($null -ne $ex.Data -and $ex.Data.Contains("StatusCode")) { $StatusCode = $ex.Data["StatusCode"] }

            if ($StatusCode -eq 403) {
                throw "Access denied to '$($Task.SourcePath)' on site '$($Site.DisplayName)'. Check that the signed-in user has access to this site/library."
            }

            $ErrorMsg = "Source path '$($Task.SourcePath)' not found."
            if ($Task.LibraryName -and $CleanPath -like "$($Task.LibraryName)/*") {
                $ErrorMsg += " (Hint: Your SourcePath starts with the Library Name. Try removing '$($Task.LibraryName)/' from the path.)"
            }
            throw $ErrorMsg
        }

        if ($null -ne $StartFolder.folder) {
            Copy-FolderRecursively -DriveId $TargetDrive.Id -FolderItemId $StartFolder.id -LocalBasePath $Task.LocalDownloadPath -UpdateAction $CurrentAction -DryRun:$DryRun
        }
        else {
            Write-Log "The SourcePath points to a file, not a folder." "ERROR" "Red"
        }
    }
    catch {
        Write-Log "TASK FAILED: $_" "ERROR" "Red"
        $Script:TaskStats.Errors++
    }

    # FIX #7: Per-task summary
    $DownloadLabel = if ($DryRun) { "Ready to Download" } else { "Downloaded" }
    $TaskColor = if ($Script:TaskStats.Errors -gt 0) { "Yellow" } else { "Green" }
    Write-Log "Task #$TaskIndex Summary: $DownloadLabel=$($Script:TaskStats.Downloaded), Skipped=$($Script:TaskStats.Skipped), Errors=$($Script:TaskStats.Errors)" "INFO" $TaskColor

    # Accumulate into global stats
    $Script:TotalStats.Downloaded += $Script:TaskStats.Downloaded
    $Script:TotalStats.Skipped += $Script:TaskStats.Skipped
    $Script:TotalStats.Errors += $Script:TaskStats.Errors
}

# --- 6. Cleanup & Summary ---
try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { Write-Log "Disconnect warning: $_" "WARN" "Yellow" }

Write-Log "========================================" "INFO" "Cyan"
Write-Log "Backup Summary Report (All Tasks)" "INFO" "Cyan"
$DownloadLabel = if ($DryRun) { "Files ready to Download" } else { "Files Downloaded" }
Write-Log "$DownloadLabel : $($Script:TotalStats.Downloaded)" "INFO" "Green"
Write-Log "Files Skipped    : $($Script:TotalStats.Skipped)" "INFO" "Gray"
Write-Log "Errors           : $($Script:TotalStats.Errors)" "INFO" "Red"
Write-Log "All tasks finished." "INFO" "Magenta"
