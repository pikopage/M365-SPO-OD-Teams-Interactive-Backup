<#
.SYNOPSIS
    WPF GUI wrapper for M365-SPO-OD-Teams Interactive Backup.
.DESCRIPTION
    Runs Backup-M365-Interactive.ps1 in a hidden process and displays
    real-time progress by tailing the script's log file.

    Features:
      - Dark-themed WPF window with color-coded log viewer
      - Real-time statistics (downloaded / skipped / errors / elapsed)
      - Dry-run toggle, cancel support, window-close guard
      - Config summary loaded from config.json

    Place this file in the same folder as Backup-M365-Interactive.ps1
    and config.json, then run:  .\Backup-GUI.ps1

.NOTES
    Authentication: The backup script calls Connect-MgGraph which opens a
    browser for interactive sign-in. This works even though the console
    window is hidden. If you experience auth issues, change CreateNoWindow
    to $false in the Start button handler.
#>

# ── Assemblies ──────────────────────────────────────────────
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ── Paths ───────────────────────────────────────────────────
$ScriptDir        = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$BackupScriptPath = Join-Path $ScriptDir 'Backup-M365-Interactive.ps1'
$ConfigPath       = Join-Path $ScriptDir 'config.json'
$PwshExe          = if ($PSVersionTable.PSVersion.Major -ge 7) { 'pwsh.exe' } else { 'powershell.exe' }

# ── WPF XAML ────────────────────────────────────────────────
[xml]$Xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="M365 Interactive Backup"
    Height="720" Width="980" MinHeight="480" MinWidth="660"
    WindowStartupLocation="CenterScreen" Background="#1E1E2E"
    ResizeMode="CanResizeWithGrip">

    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="FontSize"        Value="13"/>
            <Setter Property="Padding"         Value="14,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Margin"          Value="0,0,6,0"/>
        </Style>
    </Window.Resources>

    <Grid Margin="14">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title -->
        <TextBlock Grid.Row="0"
                   Text="M365 SPO-OD-Teams Interactive Backup"
                   FontSize="20" FontWeight="SemiBold" Foreground="#89B4FA"
                   Margin="0,0,0,10"/>

        <!-- Toolbar -->
        <WrapPanel Grid.Row="1" Margin="0,0,0,10">
            <Button x:Name="btnStart" Content="Start Backup"
                    Background="#16A34A" Foreground="White" FontWeight="Bold"/>
            <Button x:Name="btnStop" Content="Stop"
                    Background="#DC2626" Foreground="White" IsEnabled="False"/>
            <CheckBox x:Name="chkDryRun" Content=" Dry Run (no downloads)"
                      Foreground="#FFFFFF" VerticalAlignment="Center"
                      Margin="12,0,18,0" FontSize="13"/>
            <TextBlock Text="Update Mode:" Foreground="#E4E8F0"
                       VerticalAlignment="Center" Margin="6,0,4,0" FontSize="13"/>
            <Button x:Name="btnUpdateAction" Content="RenameNew"
                    Background="#2563EB" Foreground="White" FontWeight="Bold"
                    MinWidth="120" Margin="0,0,12,0"/>
            <Button x:Name="btnConfig" Content="Edit Config"
                    Background="#45475A" Foreground="#FFFFFF"/>
            <Button x:Name="btnFolder" Content="Open Folder"
                    Background="#45475A" Foreground="#FFFFFF"/>
        </WrapPanel>

        <!-- Task summary from config.json -->
        <Border Grid.Row="2" Background="#1E1E30" CornerRadius="6"
                Padding="12,8" Margin="0,0,0,10">
            <TextBlock x:Name="txtTasks" Text="Loading config..."
                       Foreground="#E0E4EC" FontFamily="Consolas" FontSize="12"
                       TextWrapping="Wrap"/>
        </Border>

        <!-- Log viewer -->
        <Border Grid.Row="3" Background="#11111B" CornerRadius="6" Padding="2">
            <RichTextBox x:Name="rtbLog" IsReadOnly="True"
                         Background="Transparent" Foreground="#CDD6F4"
                         FontFamily="Consolas" FontSize="12.5"
                         VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Disabled"
                         BorderThickness="0" Padding="4">
                <FlowDocument LineStackingStrategy="BlockLineHeight"
                              LineHeight="20" PagePadding="4"/>
            </RichTextBox>
        </Border>

        <!-- Status bar -->
        <Grid Grid.Row="4" Margin="0,10,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <TextBlock x:Name="txtStatus" Grid.Column="0"
                       Text="Ready" Foreground="#D0D4DC"
                       VerticalAlignment="Center" FontSize="12"/>
            <TextBlock x:Name="txtDownloaded" Grid.Column="1"
                       Text="Downloaded: 0" Foreground="#80F0A0"
                       Margin="14,0" VerticalAlignment="Center" FontSize="12"/>
            <TextBlock x:Name="txtSkipped" Grid.Column="2"
                       Text="Skipped: 0" Foreground="#A0A6B8"
                       Margin="14,0" VerticalAlignment="Center" FontSize="12"/>
            <TextBlock x:Name="txtErrors" Grid.Column="3"
                       Text="Errors: 0" Foreground="#FF8CA0"
                       Margin="14,0" VerticalAlignment="Center" FontSize="12"/>
            <TextBlock x:Name="txtElapsed" Grid.Column="4"
                       Text="" Foreground="#A0C8FF"
                       Margin="14,0,0,0" VerticalAlignment="Center" FontSize="12"/>
            <TextBlock x:Name="txtAuthUser" Grid.Column="5"
                       Text="" Foreground="#80F0D8"
                       Margin="20,0,0,0" VerticalAlignment="Center" FontSize="12"/>
        </Grid>
    </Grid>
</Window>
"@

# ── Load the window and resolve controls ────────────────────
$Reader = [System.Xml.XmlNodeReader]::new($Xaml)
$Window = [System.Windows.Markup.XamlReader]::Load($Reader)

$btnStart      = $Window.FindName('btnStart')
$btnStop       = $Window.FindName('btnStop')
$chkDryRun     = $Window.FindName('chkDryRun')
$btnConfig     = $Window.FindName('btnConfig')
$btnFolder     = $Window.FindName('btnFolder')
$cmbUpdateAction = $Window.FindName('btnUpdateAction')
$txtTasks      = $Window.FindName('txtTasks')
$rtbLog        = $Window.FindName('rtbLog')
$txtStatus     = $Window.FindName('txtStatus')
$txtDownloaded = $Window.FindName('txtDownloaded')
$txtSkipped    = $Window.FindName('txtSkipped')
$txtErrors     = $Window.FindName('txtErrors')
$txtElapsed    = $Window.FindName('txtElapsed')
$txtAuthUser   = $Window.FindName('txtAuthUser')

# ── Script-scope state ──────────────────────────────────────
$script:Proc             = $null
$script:LogFile          = $null
$script:LogPos           = 0
$script:Downloads        = 0
$script:Skips            = 0
$script:Errs             = 0
$script:StartTime        = $null
$script:KnownLogFiles    = @()
$script:LogFileFound     = $false
$script:ExitGraceTicks   = 0
$script:AuthFailed       = $false

# ── Helper: choose colour based on log-line content ─────────
function Get-LineColor([string]$Line) {
    if ($Line -match '^-{3,}' -or $Line -match '={3,}')         { return '#6C708A' }
    if ($Line -match '\[ERROR\]')                              { return '#FF8CA0' }
    if ($Line -match '\[WARN\]')                               { return '#FFE8A0' }
    if ($Line -match 'Downloaded:|DRYRUN.*[Ww]ould download')  { return '#80F0A0' }
    if ($Line -match 'Skipping:')                              { return '#8C90A4' }
    if ($Line -match 'Task #\d+ Summary')                      { return '#D4B8FF' }
    if ($Line -match 'Starting Task|Backup Summary|All tasks') { return '#A0C8FF' }
    if ($Line -match 'Connected|Matched|Resolved|Found')       { return '#80F0D8' }
    if ($Line -match '\[GUI\]')                                { return '#FFC8E8' }
    return '#E0E4F0'
}

# ── Helper: append a coloured line to the RichTextBox ────────
function Add-LogLine([string]$Text) {
    $para = New-Object System.Windows.Documents.Paragraph
    $para.Margin = [System.Windows.Thickness]::new(0)
    $run  = New-Object System.Windows.Documents.Run($Text)
    $run.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString(
                          (Get-LineColor $Text))
    $para.Inlines.Add($run)
    $rtbLog.Document.Blocks.Add($para)

    # Cap at 8 000 lines to prevent UI slowdown on huge backups
    while ($rtbLog.Document.Blocks.Count -gt 8000) {
        $rtbLog.Document.Blocks.Remove($rtbLog.Document.Blocks.FirstBlock)
    }
    $rtbLog.ScrollToEnd()
}

# ── Helper: update live counters from a log line ─────────────
function Update-Stats([string]$Line) {
    if     ($Line -match 'Downloaded:'  -and $Line -notmatch 'Files Downloaded') { $script:Downloads++ }
    elseif ($Line -match '\[DRYRUN\].*[Ww]ould download')                        { $script:Downloads++ }
    if     ($Line -match 'Skipping:')                                            { $script:Skips++ }
    if     ($Line -match '\[ERROR\]')                                            { $script:Errs++ }

    $txtDownloaded.Text = "Downloaded: $($script:Downloads)"
    $txtSkipped.Text    = "Skipped: $($script:Skips)"
    $txtErrors.Text     = "Errors: $($script:Errs)"

    if ($Line -match 'Starting Task #(\d+)') {
        $txtStatus.Text = "Running Task #$($Matches[1])..."
    }

    # Parse authenticated user emitted by backup script after Connect-MgGraph
    if ($Line -match '\[AUTH\] (.+?)\s+\|') {
        $txtAuthUser.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#80F0D8')
        $txtAuthUser.Text = "Signed in: $($Matches[1])"
    }

    # Auth failure — stop counting further errors as backup errors, surface clearly
    if ($Line -match '\[AUTH-FAILED\]') {
        $script:AuthFailed = $true
        $txtAuthUser.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF8CA0')
        $txtAuthUser.Text = 'Auth failed'
        $txtStatus.Text   = 'Authentication failed'
    }
}

# ── Helper: tail new content from the backup script's log ────
function Read-NewLogContent {
    # Discover the log file the backup script just created
    if (-not $script:LogFileFound) {
        $candidate = Get-ChildItem -Path $ScriptDir -Filter 'script_log_*.txt' `
                         -ErrorAction SilentlyContinue |
                     Where-Object { $_.Name -notin $script:KnownLogFiles } |
                     Sort-Object CreationTime -Descending |
                     Select-Object -First 1
        if ($candidate) {
            $script:LogFile      = $candidate.FullName
            $script:LogPos       = 0
            $script:LogFileFound = $true
        }
        else { return }
    }

    if (-not (Test-Path $script:LogFile)) { return }

    try {
        $fs = [System.IO.FileStream]::new(
                  $script:LogFile,
                  [System.IO.FileMode]::Open,
                  [System.IO.FileAccess]::Read,
                  [System.IO.FileShare]::ReadWrite)
        $fs.Position = $script:LogPos
        $sr = [System.IO.StreamReader]::new($fs)
        $chunk = $sr.ReadToEnd()
        $script:LogPos = $fs.Position
        $sr.Close(); $fs.Close()

        if ([string]::IsNullOrEmpty($chunk)) { return }

        foreach ($ln in ($chunk -split "`r?`n")) {
            if ($ln -eq '') { continue }
            Add-LogLine $ln
            Update-Stats $ln
        }
    }
    catch {
        # File might be momentarily locked; retry on next tick
    }
}

# ── Polling timer (runs on UI thread) ───────────────────────
$script:Timer = New-Object System.Windows.Threading.DispatcherTimer
$script:Timer.Interval = [TimeSpan]::FromMilliseconds(500)
$script:Timer.Add_Tick({

    Read-NewLogContent

    # Elapsed clock
    if ($script:StartTime) {
        $elapsed = (Get-Date) - $script:StartTime
        $txtElapsed.Text = $elapsed.ToString('hh\:mm\:ss')
    }

    # Detect process completion — wait a few ticks so the log file flushes
    if ($script:Proc) {
        $exited = $false
        try { $exited = $script:Proc.HasExited } catch { $exited = $true }

        if ($exited) {
            $script:ExitGraceTicks++
            if ($script:ExitGraceTicks -ge 4) {          # ~2 s grace
                Read-NewLogContent                        # final read
                $script:Timer.Stop()

                $exitCode = try { $script:Proc.ExitCode } catch { -1 }
                $script:Proc = $null

                if (-not $script:LogFileFound) {
                    Add-LogLine '[GUI] Process exited without creating a log file.'
                    Add-LogLine '[GUI] Ensure Microsoft.Graph modules are installed and config.json is valid.'
                }

                $resultText = if ($script:AuthFailed)    { 'Authentication failed — check credentials / consent' }
                              elseif ($script:Errs -gt 0) { 'Completed with errors' }
                              else                         { 'Completed successfully' }
                $txtStatus.Text = "$resultText  (exit $exitCode)"
                $Window.Title   = 'M365 Interactive Backup'
                $btnStart.IsEnabled = $true
                $btnStop.IsEnabled  = $false

                $Window.Activate()
            }
        }
    }
})

# ── Load and display config.json summary ────────────────────
function Show-ConfigSummary {
    if (-not (Test-Path $ConfigPath)) {
        $txtTasks.Text = "config.json not found in $ScriptDir"
        return
    }
    try {
        $tasks = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        $descs = @()
        $i = 0
        foreach ($t in $tasks) {
            $i++
            $src = if ($t.SourcePath) { $t.SourcePath } else { '(root)' }
            if ($t.Type -eq 'SharePoint') {
                $site = if ($t.SiteUrl) { $t.SiteUrl }
                        elseif ($t.SiteName) { $t.SiteName }
                        else { '?' }
                $descs += "  #$i  SP   $site / $($t.LibraryName)  [$src]  ->  $($t.LocalDownloadPath)"
            }
            elseif ($t.Type -eq 'OneDrive') {
                $user = if ($t.TargetUser) { $t.TargetUser } else { '(current user)' }
                $descs += "  #$i  OD   $user  [$src]  ->  $($t.LocalDownloadPath)"
            }
            else {
                $descs += "  #$i  ??   Unknown type: $($t.Type)"
            }
        }
        $txtTasks.Text = "$($tasks.Count) task(s) configured:`n" + ($descs -join "`n")
    }
    catch {
        $txtTasks.Text = "Error reading config.json: $_"
    }
}
Show-ConfigSummary

# ── Startup: check WAM / AAD join status ────────────────────
try {
    $dsreg = (& dsregcmd /status 2>$null) -join "`n"
    $aadJoined = $dsreg -match 'AzureAdJoined\s*:\s*YES'
    $hasPrt    = $dsreg -match 'AzureAdPrt\s*:\s*YES'

    if ($aadJoined -and $hasPrt) {
        $txtStatus.Text   = 'Ready  -  silent login (AAD joined, no popup)'
        $txtAuthUser.Text = 'WAM: available'
    }
    elseif ($aadJoined) {
        $txtStatus.Text   = 'Ready  -  AAD joined, PRT not yet issued'
        $txtAuthUser.Text = 'WAM: partial'
    }
    else {
        $txtStatus.Text   = 'Ready  -  login popup will appear on Start'
        $txtAuthUser.Text = 'WAM: unavailable'
    }
}
catch {
    $txtStatus.Text = 'Ready'
}

# ═════════════════════════════════════════════════════════════
#  Button handlers
# ═════════════════════════════════════════════════════════════

# ── Update Mode toggle ───────────────────────────────────────
$cmbUpdateAction.Add_Click({
    if ($cmbUpdateAction.Content -eq 'RenameNew') {
        $cmbUpdateAction.Content    = 'Overwrite'
        $cmbUpdateAction.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#DC8A16')
    }
    else {
        $cmbUpdateAction.Content    = 'RenameNew'
        $cmbUpdateAction.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#2563EB')
    }
})

# ── Start ────────────────────────────────────────────────────
$btnStart.Add_Click({
    if (-not (Test-Path $BackupScriptPath)) {
        [System.Windows.MessageBox]::Show(
            "Backup script not found:`n$BackupScriptPath",
            'Missing Script', 'OK', 'Error') | Out-Null
        return
    }
    if (-not (Test-Path $ConfigPath)) {
        [System.Windows.MessageBox]::Show(
            "config.json not found:`n$ConfigPath",
            'Missing Config', 'OK', 'Error') | Out-Null
        return
    }

    # Reset state
    $rtbLog.Document.Blocks.Clear()
    $script:Downloads      = 0
    $script:Skips          = 0
    $script:Errs           = 0
    $script:AuthFailed     = $false
    $script:LogFile        = $null
    $script:LogPos         = 0
    $script:LogFileFound   = $false
    $script:ExitGraceTicks = 0
    $script:StartTime      = Get-Date

    $txtDownloaded.Text = 'Downloaded: 0'
    $txtSkipped.Text    = 'Skipped: 0'
    $txtErrors.Text     = 'Errors: 0'
    $txtElapsed.Text    = '00:00:00'
    $txtStatus.Text     = 'Starting...'
    $Window.Title       = 'M365 Interactive Backup  -  Running'

    $btnStart.IsEnabled = $false
    $btnStop.IsEnabled  = $true

    # Snapshot existing log files so we only detect the NEW one
    $script:KnownLogFiles = @(
        Get-ChildItem $ScriptDir -Filter 'script_log_*.txt' -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty Name)

    # Launch the backup script as a hidden PowerShell process
    $selectedAction = $cmbUpdateAction.Content.ToString()
    $procArgs = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -UpdateAction "{1}"' -f $BackupScriptPath, $selectedAction
    if ($chkDryRun.IsChecked) { $procArgs += ' -DryRun' }

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName         = $PwshExe
        $psi.Arguments        = $procArgs
        $psi.WorkingDirectory = $ScriptDir
        $psi.UseShellExecute  = $false
        $psi.CreateNoWindow   = $true

        $script:Proc = [System.Diagnostics.Process]::Start($psi)
        $script:Timer.Start()

        $dryLabel = if ($chkDryRun.IsChecked) { 'Yes' } else { 'No' }
        Add-LogLine "[GUI] Backup process started  (PID $($script:Proc.Id), DryRun=$dryLabel, UpdateAction=$selectedAction)"
        Add-LogLine '[GUI] Waiting for log output...'
    }
    catch {
        Add-LogLine "[GUI] Failed to start process: $_"
        $txtStatus.Text     = 'Failed to launch'
        $btnStart.IsEnabled = $true
        $btnStop.IsEnabled  = $false
    }
})

# ── Stop ─────────────────────────────────────────────────────
$btnStop.Add_Click({
    if (-not $script:Proc) { return }
    $exited = $false
    try { $exited = $script:Proc.HasExited } catch { $exited = $true }
    if ($exited) { return }

    $answer = [System.Windows.MessageBox]::Show(
        'Cancel the running backup?', 'Confirm Cancel',
        'YesNo', 'Warning')

    if ($answer -eq 'Yes') {
        try { $script:Proc.Kill() } catch {}
        $script:Timer.Stop()
        Add-LogLine '[GUI] Backup cancelled by user.'
        $txtStatus.Text     = 'Cancelled'
        $Window.Title       = 'M365 Interactive Backup'
        $btnStart.IsEnabled = $true
        $btnStop.IsEnabled  = $false
    }
})

# ── Edit Config ──────────────────────────────────────────────
$btnConfig.Add_Click({
    if (Test-Path $ConfigPath) { Start-Process notepad.exe $ConfigPath }
    else {
        [System.Windows.MessageBox]::Show(
            "config.json not found:`n$ConfigPath", 'Error') | Out-Null
    }
})

# ── Open Folder ──────────────────────────────────────────────
$btnFolder.Add_Click({ Start-Process explorer.exe $ScriptDir })

# ── Window-close guard ───────────────────────────────────────
$Window.Add_Closing({
    param($s, $e)

    if ($script:Proc) {
        $exited = $false
        try { $exited = $script:Proc.HasExited } catch { $exited = $true }
        if (-not $exited) {
            $answer = [System.Windows.MessageBox]::Show(
                'A backup is still running. Stop it and close?',
                'Confirm Exit', 'YesNo', 'Warning')
            if ($answer -eq 'No') {
                $e.Cancel = $true
                return
            }
            try { $script:Proc.Kill() } catch {}
        }
    }
    $script:Timer.Stop()
})

# ── Show ─────────────────────────────────────────────────────
$Window.ShowDialog() | Out-Null
