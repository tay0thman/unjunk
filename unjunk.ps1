<#
.SYNOPSIS
    Master System Maintenance Script (v38 - Full Parity)
.DESCRIPTION
    v38 merges all cleanup categories from the original unjunk.ps1 into the
    v37 framework. Full feature list:

    FROM v37:  Dry-run preview, progress bars, deep-clean toggle (CrowdStrike
    safe default), auto-save log, framework-safe Appx pruning, WindowsApps
    orphan scanner, DISM component cleanup, accurate size reporting.

    NEW IN v38:
    - 50+ additional junk targets from unjunk.ps1 (V-Ray, Teams, Google Earth,
      DriveFS, VSCode, Elgato, Epic, WhatsApp, pip, qBittorrent, WDF, Maps,
      Remote Desktop, driver install roots, etc.)
    - Step 7: Windows Update cleanup (stops wuauserv, clears SoftwareDistribution)
    - Step 8: Windows Explorer & Privacy (registry MRU, recent items, jump lists,
      shim cache, AppCompat flags — kills & restarts explorer.exe)
    - Rhino installer scan: now removes Rhino + Rhino language packs but
      PRESERVES Rhino.Inside installations.
    - Expanded process kill list (all processes from unjunk.ps1)
    - Windows Defender scan history purge (ScanPurgeItemsAfterDelay = 1)

.PARAMETER Force
    Skips confirmation prompts — runs all steps including deep clean.
.PARAMETER DryRun
    Scan-only mode. Shows what WOULD be deleted without touching any files.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$Force,
    [switch]$DryRun
)

# --- CONFIGURATION ---
$ErrorActionPreference = "SilentlyContinue"
$Host.UI.RawUI.WindowTitle = "Master System Maintenance v38"

# --- GLOBAL REPORTING STATE ---
$global:CleanupReport     = [System.Collections.Generic.List[PSCustomObject]]::new()
$global:RebootQueue       = [System.Collections.Generic.List[string]]::new()
$global:TotalBytesDeleted = 0
$global:TotalBytesFailed  = 0
$global:TotalFilesDeleted = 0
$global:TotalFilesFailed  = 0
$global:TotalRebootQueued = 0
$global:ScriptStopwatch   = [System.Diagnostics.Stopwatch]::StartNew()
$global:StepTimings       = [System.Collections.Generic.List[PSCustomObject]]::new()
$global:IsDryRun          = $DryRun.IsPresent

# --- ADMIN CHECK ---
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)
if (-not $IsAdmin) {
    Write-Warning "You must run this script as Administrator!"
    Write-Warning "Right-click PowerShell and select 'Run as Administrator'."
    Break
}

# --- NATIVE API: MOVEFILEEX (DELETE ON REBOOT) ---
$MoveFileCode = @'
    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, int dwFlags);
'@
try {
    $null = Add-Type -MemberDefinition $MoveFileCode -Name "Kernel32Helper" -Namespace "Win32" -PassThru -ErrorAction Stop
} catch { }

function Register-DeleteOnReboot {
    param([string]$FilePath)
    $result = [Win32.Kernel32Helper]::MoveFileEx($FilePath, $null, 4)
    if ($result) {
        $global:RebootQueue.Add($FilePath)
        $global:TotalRebootQueued++
    } else {
        Write-Warning "  MoveFileEx failed for: $FilePath (Error: $([System.Runtime.InteropServices.Marshal]::GetLastWin32Error()))"
    }
}

# =====================================================================
#  HELPER FUNCTIONS
# =====================================================================

function Format-Size {
    param([double]$Bytes)
    if ($Bytes -lt 0) { return "-" + (Format-Size ([Math]::Abs($Bytes))) }
    switch ($Bytes) {
        { $_ -ge 1GB } { return "{0:N2} GB" -f ($_ / 1GB) }
        { $_ -ge 1MB } { return "{0:N2} MB" -f ($_ / 1MB) }
        { $_ -ge 1KB } { return "{0:N2} KB" -f ($_ / 1KB) }
        default         { return "{0:N0} B"  -f $_ }
    }
}

function Format-Elapsed {
    param([TimeSpan]$Span)
    if ($Span.TotalMinutes -ge 1) { return "{0:N1} min" -f $Span.TotalMinutes }
    return "{0:N1} sec" -f $Span.TotalSeconds
}

function Start-StepTimer {
    param([string]$StepName)
    [PSCustomObject]@{ Name = $StepName; Timer = [System.Diagnostics.Stopwatch]::StartNew() }
}

function Stop-StepTimer {
    param([PSCustomObject]$Step)
    $Step.Timer.Stop()
    $global:StepTimings.Add([PSCustomObject]@{
        Name    = $Step.Name
        Elapsed = $Step.Timer.Elapsed
    })
}

function Stop-TargetProcess {
    param([string]$ProcessName)
    $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($procs) {
        $count = @($procs).Count
        Stop-Process -Name $ProcessName -Force -Confirm:$false -ErrorAction SilentlyContinue
        return $count
    }
    return 0
}

function Write-StepHeader {
    param([string]$Title, [int]$Number)
    Write-Host ""
    Write-Host "  ┌─ STEP $Number ─────────────────────────────────────────────" -ForegroundColor DarkCyan
    Write-Host "  │ $Title" -ForegroundColor Cyan
    Write-Host "  └────────────────────────────────────────────────────────" -ForegroundColor DarkCyan
}

# =====================================================================
#  RESOLVE JUNK PATHS → returns file list with sizes (no deletion)
# =====================================================================
function Resolve-JunkFiles {
    param([string]$Path)
    $resolvedPaths = @()
    try {
        $expandedPath = [System.Environment]::ExpandEnvironmentVariables($Path)
        if ($expandedPath -match '[\*\?]') {
            $resolvedPaths = @(Resolve-Path -Path $expandedPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Path)
        } elseif (Test-Path -LiteralPath $expandedPath) {
            $resolvedPaths = @($expandedPath)
        }
    } catch { }

    $allFiles = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    foreach ($resolved in $resolvedPaths) {
        if (Test-Path -LiteralPath $resolved -PathType Leaf) {
            $f = Get-Item -LiteralPath $resolved -Force -ErrorAction SilentlyContinue
            if ($f) { $allFiles.Add($f) }
        } else {
            $children = Get-ChildItem -LiteralPath $resolved -Recurse -Force -File -ErrorAction SilentlyContinue
            foreach ($f in $children) { $allFiles.Add($f) }
        }
    }
    return $allFiles
}

# =====================================================================
#  DELETE FILES with progress bar (operates on pre-scanned file list)
# =====================================================================
function Remove-ScannedFiles {
    param(
        [string]$Desc,
        [System.Collections.Generic.List[System.IO.FileInfo]]$Files,
        [string]$Path
    )
    if ($Files.Count -eq 0) {
        # No files, but still attempt to remove empty directory trees
        $resolvedPaths = @()
        try {
            $expandedPath = [System.Environment]::ExpandEnvironmentVariables($Path)
            if ($expandedPath -match '[\*\?]') {
                $resolvedPaths = @(Resolve-Path -Path $expandedPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Path)
            } elseif (Test-Path -LiteralPath $expandedPath) {
                $resolvedPaths = @($expandedPath)
            }
        } catch { }
        foreach ($resolved in $resolvedPaths) {
            if (Test-Path -LiteralPath $resolved -PathType Container) {
                try { Remove-Item -LiteralPath $resolved -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue } catch { }
            }
        }
        return
    }

    $entryBytesDeleted = 0; $entryBytesFailed = 0
    $entryFilesDeleted = 0; $entryFilesFailed = 0; $entryRebootQueued = 0

    $totalBytes = ($Files | Measure-Object -Property Length -Sum).Sum
    $processedBytes = 0
    $i = 0

    foreach ($file in $Files) {
        $i++
        $fileSize = $file.Length
        $pct = if ($totalBytes -gt 0) { [math]::Min(100, [int](($processedBytes / $totalBytes) * 100)) } else { 0 }

        Write-Progress -Activity "Deleting: $Desc" `
                       -Status "$i / $($Files.Count) files  |  $(Format-Size $global:TotalBytesDeleted) recovered total" `
                       -PercentComplete $pct -Id 1

        try {
            Remove-Item -LiteralPath $file.FullName -Force -Confirm:$false -ErrorAction Stop
            $entryBytesDeleted += $fileSize
            $entryFilesDeleted++
            $global:TotalBytesDeleted += $fileSize
            $global:TotalFilesDeleted++
        } catch {
            Register-DeleteOnReboot -FilePath $file.FullName
            $entryBytesFailed += $fileSize
            $entryFilesFailed++
            $entryRebootQueued++
            $global:TotalBytesFailed += $fileSize
            $global:TotalFilesFailed++
        }
        $processedBytes += $fileSize
    }

    Write-Progress -Activity "Deleting: $Desc" -Completed -Id 1

    # Clean up empty directories
    $resolvedPaths = @()
    try {
        $expandedPath = [System.Environment]::ExpandEnvironmentVariables($Path)
        if ($expandedPath -match '[\*\?]') {
            $resolvedPaths = @(Resolve-Path -Path $expandedPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Path)
        } elseif (Test-Path -LiteralPath $expandedPath) {
            $resolvedPaths = @($expandedPath)
        }
    } catch { }

    foreach ($resolved in $resolvedPaths) {
        $allDirs = @(Get-ChildItem -LiteralPath $resolved -Recurse -Force -Directory -ErrorAction SilentlyContinue)
        $allDirs | Sort-Object { $_.FullName.Length } -Descending | ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -Confirm:$false -ErrorAction Stop } catch { }
        }
        if (Test-Path -LiteralPath $resolved -PathType Container) {
            try { Remove-Item -LiteralPath $resolved -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue } catch { }
        }
    }

    if (($entryFilesDeleted + $entryFilesFailed) -gt 0) {
        $global:CleanupReport.Add([PSCustomObject]@{
            Description  = $Desc
            FilesDeleted = $entryFilesDeleted; BytesDeleted = $entryBytesDeleted
            FilesFailed  = $entryFilesFailed;  BytesFailed  = $entryBytesFailed
            RebootQueued = $entryRebootQueued
        })
    }
}

# =====================================================================
#  JUNK PATH DEFINITIONS (data-driven)
# =====================================================================

# Build Revit year-based paths dynamically
$RevitPaths = @()
2018..2030 | ForEach-Object {
    $RevitPaths += @{ Desc = "Revit $_ Collab";    Path = "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit $_\CollaborationCache\*" }
    $RevitPaths += @{ Desc = "Revit $_ Journals";  Path = "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit $_\Journals\*" }
}

# Build Rhino autosave paths dynamically
$RhinoPaths = @()
6..10 | ForEach-Object {
    $RhinoPaths += @{ Desc = "Rhino $_.0 AutoSave"; Path = "$env:LOCALAPPDATA\McNeel\Rhinoceros\$_.0\AutoSave\*" }
}

$JunkTargets = @(
    # ── System Temps & Crash Data ──
    @{ Desc = "System Temp";              Path = "$env:SYSTEMROOT\Temp\*" }
    @{ Desc = "User Temp";                Path = "$env:TEMP\*" }
    @{ Desc = "Local Temp";               Path = "$env:LOCALAPPDATA\Temp\*" }
    @{ Desc = "Roaming Temp";             Path = "$env:APPDATA\Temp\*" }
    @{ Desc = "LocalLow Temp";            Path = "$env:USERPROFILE\AppData\LocalLow\Temp\*" }
    @{ Desc = "Prefetch";                 Path = "$env:SYSTEMROOT\Prefetch\*" }
    @{ Desc = "Live Kernel Reports";      Path = "$env:SYSTEMROOT\LiveKernelReports\*" }
    @{ Desc = "Minidumps";                Path = "$env:SYSTEMROOT\Minidump\*" }
    @{ Desc = "Memory Dump";              Path = "$env:SYSTEMROOT\MEMORY.DMP" }
    @{ Desc = "Crash Dumps";              Path = "$env:LOCALAPPDATA\CrashDumps\*" }
    @{ Desc = "WER Archives";             Path = "C:\ProgramData\Microsoft\Windows\WER\ReportArchive\*" }
    @{ Desc = "WER Queue";                Path = "C:\ProgramData\Microsoft\Windows\WER\ReportQueue\*" }
    @{ Desc = "CHKDSK Fragments";         Path = "$env:SystemDrive\File*.chk" }
    @{ Desc = "MATS Troubleshooter";      Path = "C:\MATS" }
    @{ Desc = "SRU Monitor";              Path = "C:\Windows\System32\sru\*" }
    @{ Desc = "DO Service Logs";          Path = "C:\WINDOWS\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Logs\*" }

    # ── Autodesk / Revit ──
    @{ Desc = "Revit PacCache";           Path = "$env:LOCALAPPDATA\Autodesk\Revit\PacCache\*" }
    @{ Desc = "Autodesk Install Root";    Path = "C:\Autodesk" }
    @{ Desc = "Autodesk ADPSDK JSON";     Path = "$env:APPDATA\Autodesk\ADPSDK\JSON" }
    @{ Desc = "Revit InterProcess";       Path = "C:\ProgramData\RevitInterProcess\*" }
    @{ Desc = "ACCDocs";                  Path = "$env:HOMEPATH\ACCDOCS\*" }

    # ── Adobe ──
    @{ Desc = "Adobe Temp Root";          Path = "C:\adobeTemp" }
    @{ Desc = "Adobe CC Libraries";       Path = "$env:APPDATA\Adobe\Creative Cloud Libraries\*" }
    @{ Desc = "Adobe Dunamis";            Path = "$env:APPDATA\com.adobe.dunamis\*" }
    @{ Desc = "Illustrator ACPL Logs";    Path = "$env:APPDATA\Adobe\Logs\Adobe Illustrator\*\Adobe Illustrator\ACPLLogs\*" }
    @{ Desc = "Lightroom Cache";          Path = "$env:LOCALAPPDATA\Adobe\Lightroom\Caches\*" }
    @{ Desc = "InDesign Caches";          Path = "$env:LOCALAPPDATA\Adobe\InDesign\*\*\Caches\*" }
    @{ Desc = "Adobe Media Cache";        Path = "$env:APPDATA\Adobe\Common\Media Cache Files\*" }
    @{ Desc = "Adobe Installer Logs";     Path = "C:\Program Files (x86)\Common Files\Adobe\Installers\*.log" }

    # ── Chaos Group / V-Ray ──
    @{ Desc = "Chaos Vantage Cache";      Path = "$env:LOCALAPPDATA\Chaos Group\Vantage\cache\*" }
    @{ Desc = "Chaos Cosmos Updates";     Path = "$env:LOCALAPPDATA\Chaos\Cosmos\Updates\*" }
    @{ Desc = "V-Ray Rhino Logs";         Path = "$env:APPDATA\Chaos Group\V-Ray for Rhinoceros\vrayneui\*.log" }

    # ── McNeel / Rhino ──
    @{ Desc = "Rhino System Update Cache"; Path = "C:\ProgramData\McNeel\McNeelUpdate\DownloadCache\*" }
    @{ Desc = "Rhino User Update Cache";  Path = "$env:LOCALAPPDATA\McNeel\McNeelUpdate\DownloadCache\*" }

    # ── Bluebeam ──
    @{ Desc = "Bluebeam Sessions";        Path = "$env:LOCALAPPDATA\Revu\data\Sessions\studio.bluebeam.com\*" }
    @{ Desc = "Bluebeam WebCache";        Path = "$env:LOCALAPPDATA\Bluebeam\Revu\*\WebCache" }
    @{ Desc = "Bluebeam Recovery";        Path = "$env:LOCALAPPDATA\Bluebeam\Revu\*\Recovery\*" }
    @{ Desc = "Bluebeam Logs";            Path = "$env:LOCALAPPDATA\Bluebeam\Revu\*\Logs\*" }

    # ── Honeybee / Ladybug ──
    @{ Desc = "Honeybee Simulations";     Path = "$env:USERPROFILE\simulation\*" }

    # ── Microsoft Teams (New Store App) ──
    @{ Desc = "Teams WebStorage";         Path = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_tfw\WebStorage\*" }
    @{ Desc = "Teams Cache";              Path = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_tfw\Cache\*" }
    @{ Desc = "Teams SW CacheStorage";    Path = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_tfw\Service Worker\CacheStorage\*" }
    @{ Desc = "Teams SW ScriptCache";     Path = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_tfw\Service Worker\ScriptCache\*" }
    @{ Desc = "Teams Logs";               Path = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs\*" }
    # Classic Teams (legacy desktop app)
    @{ Desc = "Teams Classic SW Cache";   Path = "$env:APPDATA\Microsoft\Teams\Service Worker\CacheStorage\*" }
    @{ Desc = "Teams Classic Cache";      Path = "$env:APPDATA\Microsoft\Teams\Cache\*" }

    # ── Google ──
    @{ Desc = "Google Earth Cache";       Path = "$env:USERPROFILE\AppData\LocalLow\Google\GoogleEarth\Cache\*" }
    @{ Desc = "Google DriveFS Photos";    Path = "$env:LOCALAPPDATA\Google\DriveFS\*\photos_cache_temp" }
    @{ Desc = "Google DriveFS Logs";      Path = "$env:LOCALAPPDATA\Google\DriveFS\logs\*" }
    @{ Desc = "Google Updater CRX Cache"; Path = "C:\Program Files (x86)\Google\GoogleUpdater\crx_cache\*" }
    @{ Desc = "Google Update Downloads";  Path = "C:\Program Files (x86)\Google\Update\Download\*" }

    # ── Browsers: Chrome ──
    @{ Desc = "Chrome Default Cache";     Path = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache\*" }
    @{ Desc = "Chrome Profile 1 Cache";   Path = "$env:LOCALAPPDATA\Google\Chrome\User Data\Profile 1\Cache\Cache_Data\*" }
    @{ Desc = "Chrome Profile 1 Code";    Path = "$env:LOCALAPPDATA\Google\Chrome\User Data\Profile 1\Code Cache\*" }
    @{ Desc = "Chrome Profile 2 Code";    Path = "$env:LOCALAPPDATA\Google\Chrome\User Data\Profile 2\Code Cache\*" }

    # ── Browsers: Edge ──
    @{ Desc = "Edge Cache";               Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache\Cache_Data\*" }
    @{ Desc = "Edge Code Cache JS";       Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache\js\*" }
    @{ Desc = "Edge Code Cache WASM";     Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache\wasm\*" }
    @{ Desc = "Edge Service Workers";     Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Service Worker\CacheStorage\*" }
    @{ Desc = "Edge ScriptCache";         Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Service Worker\ScriptCache\*" }
    @{ Desc = "Edge IndexedDB";           Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\IndexedDB\*" }
    @{ Desc = "Edge BrowserMetrics";      Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\BrowserMetrics\*" }
    @{ Desc = "Edge BrowserMetrics PMA";  Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\BrowserMetrics-spare.pma" }
    @{ Desc = "Edge Profile 1 Cache";     Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Profile 1\Cache\Cache_Data\*" }
    @{ Desc = "Edge Update Installers";   Path = "C:\Program Files (x86)\Microsoft\EdgeUpdate\Download\*" }

    # ── Browsers: Misc ──
    @{ Desc = "Windows WebCache";         Path = "$env:LOCALAPPDATA\Microsoft\Windows\WebCache\*" }
    @{ Desc = "Temporary Internet Files"; Path = "$env:LOCALAPPDATA\Microsoft\Windows\Temporary Internet Files\*" }
    @{ Desc = "pip Cache";                Path = "$env:LOCALAPPDATA\pip\*" }

    # ── VSCode ──
    @{ Desc = "VSCode Cache";             Path = "$env:APPDATA\Code\Cache\*" }
    @{ Desc = "VSCode CachedData";        Path = "$env:APPDATA\Code\CachedData\*" }

    # ── Communication / Remote ──
    @{ Desc = "Zoom Logs";                Path = "$env:APPDATA\Zoom\logs\*" }
    @{ Desc = "Zoom WebView Cache";       Path = "$env:APPDATA\Zoom\data\WebviewCacheX64\*\EBWebView\*" }
    @{ Desc = "TeamViewer Cache";         Path = "$env:LOCALAPPDATA\TeamViewer\*" }
    @{ Desc = "Remote Desktop Cache";     Path = "$env:LOCALAPPDATA\Microsoft\Terminal Server Client\Cache\*" }
    @{ Desc = "WhatsApp SW Cache";        Path = "$env:LOCALAPPDATA\Packages\5319275A.WhatsAppDesktop_*\LocalCache\*\WhatsApp\Service Worker\CacheStorage" }

    # ── IDM ──
    @{ Desc = "IDM Download Data";        Path = "$env:APPDATA\IDM\DwnlData\*" }

    # ── NVIDIA ──
    @{ Desc = "NVIDIA DX Cache";          Path = "$env:LOCALAPPDATA\NVIDIA\DXCache\*" }
    @{ Desc = "NVIDIA GL Cache";          Path = "$env:LOCALAPPDATA\NVIDIA\GLCache\*" }
    @{ Desc = "NVIDIA PerDriver DXCache"; Path = "$env:USERPROFILE\AppData\LocalLow\NVIDIA\PerDriverVersion\DXCache\*" }
    @{ Desc = "NVIDIA Compute Cache";     Path = "$env:APPDATA\NVIDIA\ComputeCache\*" }

    # ── Maxon / Cinebench ──
    @{ Desc = "Maxon Redshift Cache";     Path = "$env:APPDATA\Maxon\*\Redshift\Cache\*" }
    @{ Desc = "Maxon Redshift Textures";  Path = "$env:APPDATA\Maxon\*\Redshift\Cache\Textures\*" }
    @{ Desc = "Cinebench Cache";          Path = "$env:APPDATA\Maxon\Cinebench*\cache\*" }
    @{ Desc = "Maxon Assets Cache";       Path = "$env:APPDATA\Maxon\*\assets\*" }

    # ── Logitech ──
    @{ Desc = "Logitech GHub Cache";      Path = "C:\ProgramData\LGHUB\cache\*" }
    @{ Desc = "Logitech GHub Depots";     Path = "C:\ProgramData\LGHUB\depots\*" }

    # ── Elgato ──
    @{ Desc = "Elgato CameraHub Logs";    Path = "$env:APPDATA\Elgato\CameraHub\logs\*" }
    @{ Desc = "Elgato CameraHub SW";      Path = "$env:APPDATA\Elgato\CameraHub\SW\*" }
    @{ Desc = "Elgato CameraHub Tmp";     Path = "$env:APPDATA\Elgato\CameraHub\Tmp\*" }

    # ── Discord ──
    @{ Desc = "Discord Cache";            Path = "$env:APPDATA\discord\Cache\*" }
    @{ Desc = "Discord Code Cache";       Path = "$env:APPDATA\discord\Code Cache\*" }
    @{ Desc = "Discord GPUCache";         Path = "$env:APPDATA\discord\GPUCache\*" }

    # ── Game Launchers ──
    @{ Desc = "Steam Logs";               Path = "C:\Program Files (x86)\Steam\logs\*" }
    @{ Desc = "Steam Web Cache";          Path = "$env:LOCALAPPDATA\Steam\htmlcache\*" }
    @{ Desc = "Ubisoft Cache";            Path = "C:\Program Files (x86)\Ubisoft\Ubisoft Game Launcher\cache\*" }
    @{ Desc = "Epic WebCache";            Path = "$env:LOCALAPPDATA\EpicGamesLauncher\Saved\webcache_*\*" }
    @{ Desc = "Epic Crashes";             Path = "$env:LOCALAPPDATA\EpicGamesLauncher\Saved\Crashes\*" }
    @{ Desc = "Datasmith Crashes";        Path = "$env:LOCALAPPDATA\UnrealDatasmithExporter\Saved\Crashes\*" }

    # ── Media / Torrent ──
    @{ Desc = "Eibolsoft Cache";          Path = "$env:LOCALAPPDATA\Eibolsoft\*" }
    @{ Desc = "qBittorrent Cache";        Path = "$env:LOCALAPPDATA\qBittorrent\*" }
    @{ Desc = "Gameloft Cache";           Path = "$env:LOCALAPPDATA\Gameloft\*\Cache\*" }

    # ── System Caches ──
    @{ Desc = "WDF User Data";            Path = "$env:LOCALAPPDATA\Microsoft\Windows\WDF\*" }
    @{ Desc = "WDF System Data";          Path = "C:\ProgramData\Microsoft\WDF\*" }
    @{ Desc = "Microsoft Maps Data";      Path = "C:\ProgramData\Microsoft\MapData\*" }
    @{ Desc = "Windows Notifications";    Path = "$env:LOCALAPPDATA\Microsoft\Windows\Notifications\*" }
    @{ Desc = "Action Center Cache";      Path = "$env:LOCALAPPDATA\Microsoft\Windows\ActionCenterCache\*" }
    @{ Desc = "Windows Explorer Cache";   Path = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\*" }
    @{ Desc = "Windows Caches";           Path = "$env:LOCALAPPDATA\Microsoft\Windows\Caches\*" }
    @{ Desc = "Elevated Diagnostics";     Path = "$env:LOCALAPPDATA\ElevatedDiagnostics\*" }
    @{ Desc = "Defender Scan History";    Path = "C:\ProgramData\Microsoft\Windows Defender\Scans\History\Service" }

    # ── Windows Logs ──
    @{ Desc = "NetSetup Logs";            Path = "C:\Windows\Logs\NetSetup\*" }
    @{ Desc = "SIH Logs";                 Path = "C:\Windows\Logs\SIH\*" }
    @{ Desc = "Windows Update Logs";      Path = "C:\Windows\Logs\WindowsUpdate\*" }
    @{ Desc = "USO Logs";                 Path = "C:\ProgramData\USOShared\Logs\*" }

    # ── Installers & Install Leftovers ──
    @{ Desc = "WinRE Agent";              Path = "C:\`$WinREAgent" }
    @{ Desc = "InstallShield Leftovers";  Path = "C:\Program Files (x86)\InstallShield Installation Information\*" }
    @{ Desc = "VS Installer Packages";    Path = "C:\ProgramData\Microsoft\VisualStudio\Packages\*" }
    @{ Desc = "Downloaded Installations"; Path = "$env:LOCALAPPDATA\Downloaded Installations" }
    @{ Desc = "NuGet Packages";           Path = "$env:USERPROFILE\.nuget\packages\*" }

    # ── Driver Install Extraction Roots ──
    @{ Desc = "AMD Driver Install";       Path = "$env:SystemDrive\AMD" }
    @{ Desc = "NVIDIA Driver Install";    Path = "$env:SystemDrive\NVIDIA" }
    @{ Desc = "Intel Driver Install";     Path = "$env:SystemDrive\INTEL" }

    # ── Desktop App Caches ──
    @{ Desc = "Upscayl Cache";            Path = "$env:LOCALAPPDATA\upscayl-updater\*" }
    @{ Desc = "PowerToys Updates";        Path = "$env:LOCALAPPDATA\Microsoft\PowerToys\Updates\*" }
    @{ Desc = "UniGetUI Cache";           Path = "$env:LOCALAPPDATA\UniGetUI\CachedMedia\*" }

) + $RevitPaths + $RhinoPaths

# =====================================================================
#  INTERACTIVE MENU
# =====================================================================
Clear-Host
Write-Host @"

   __  __          _____ _______ ______ _____  
  |  \/  |   /\   / ____|__   __|  ____|  __ \ 
  | \  / |  /  \ | (___    | |  | |__  | |__) |
  | |\/| | / /\ \ \___ \   | |  |  __| |  _  / 
  | |  | |/ ____ \____) |  | |  | |____| | \ \ 
  |_|  |_/_/    \_\_____/  |_|  |______|_|  \_\

  Master Maintenance v38 — By Tay Othman
"@ -ForegroundColor Green

if ($DryRun) {
    Write-Host "  ╔═══════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "  ║  DRY-RUN MODE — nothing will be deleted  ║" -ForegroundColor Yellow
    Write-Host "  ╚═══════════════════════════════════════╝" -ForegroundColor Yellow
}

# --- CUSTOM FILE DESTROYER ---
Write-Host "`n  ── CUSTOM FILE DESTROYER ──" -ForegroundColor Cyan
Write-Host "  Paste a path to force-delete, or press Enter to skip." -ForegroundColor Gray
$CustomPath = Read-Host "  Path"
$HasCustomTarget = $false
if ($CustomPath -and (Test-Path -LiteralPath $CustomPath)) {
    $HasCustomTarget = $true
    $customSize = 0
    if (Test-Path -LiteralPath $CustomPath -PathType Leaf) {
        $customSize = (Get-Item -LiteralPath $CustomPath -Force).Length
    } else {
        $customSize = (Get-ChildItem -LiteralPath $CustomPath -Recurse -Force -File -ErrorAction SilentlyContinue |
                       Measure-Object -Property Length -Sum).Sum
    }
    Write-Host "  Target: $CustomPath ($(Format-Size $customSize))" -ForegroundColor Red
} elseif ($CustomPath) {
    Write-Host "  Path not found. Skipping." -ForegroundColor DarkGray
}

# --- DEEP CLEAN TOGGLE (CrowdStrike Falcon alert risk) ---
Write-Host "`n  ── DEEP CLEANING (Shadow Copies, Event Logs, CleanMgr) ──" -ForegroundColor Cyan
Write-Host "  ⚠  May trigger CrowdStrike Falcon alerts." -ForegroundColor Yellow
$DeepClean = $false
if ($Force) {
    $DeepClean = $true
    Write-Host "  -Force: Deep cleaning enabled." -ForegroundColor Yellow
} else {
    $deepResponse = Read-Host "  Enable deep cleaning? (y/N)"
    $DeepClean = ($deepResponse -eq "y")
}

$stepsLabel = "Steps 1-8 (standard)"
if ($DeepClean) { $stepsLabel += " + Step 9 (deep clean)" }

# =====================================================================
#  PHASE 1: DRY-RUN SCAN (preview sizes before any deletion)
# =====================================================================
Write-Host "`n"
Write-Host "  ╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║               SCANNING — Calculating cleanup size...         ║" -ForegroundColor Cyan
Write-Host "  ╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

$previewTotalBytes = 0
$previewTotalFiles = 0
$previewTable = [System.Collections.Generic.List[PSCustomObject]]::new()

$targetCount = $JunkTargets.Count
$idx = 0
foreach ($target in $JunkTargets) {
    $idx++
    Write-Progress -Activity "Scanning junk targets" -Status "$idx / $targetCount  $($target.Desc)" -PercentComplete ([int]($idx / $targetCount * 100)) -Id 2
    $files = Resolve-JunkFiles -Path $target.Path
    $fileCount = $files.Count
    $fileBytes = 0
    if ($fileCount -gt 0) {
        $fileBytes = ($files | Measure-Object -Property Length -Sum).Sum
    }
    # Include in preview if files found OR if path exists as a directory (for empty dir cleanup)
    $pathExists = $false
    if ($fileCount -eq 0) {
        try {
            $exp = [System.Environment]::ExpandEnvironmentVariables($target.Path) -replace '[\*\?].*$', ''
            if ($exp -and (Test-Path -LiteralPath $exp -PathType Container)) { $pathExists = $true }
        } catch { }
    }
    if ($fileCount -gt 0 -or $pathExists) {
        $previewTable.Add([PSCustomObject]@{
            Desc      = $target.Desc
            Path      = $target.Path
            FileCount = $fileCount
            Bytes     = $fileBytes
            Files     = $files
        })
        $previewTotalBytes += $fileBytes
        $previewTotalFiles += $fileCount
    }
}
Write-Progress -Activity "Scanning junk targets" -Completed -Id 2

# Custom target preview
if ($HasCustomTarget) {
    $previewTable.Add([PSCustomObject]@{
        Desc      = "Custom: $(Split-Path $CustomPath -Leaf)"
        Path      = $CustomPath
        FileCount = 1
        Bytes     = $customSize
        Files     = $null
    })
    $previewTotalBytes += $customSize
    $previewTotalFiles++
}

# --- Display Preview Table ---
Write-Host ""
if ($previewTable.Count -gt 0) {
    $sorted = $previewTable | Sort-Object Bytes -Descending
    Write-Host "  {0,-36} {1,>10} {2,>8}" -f "TARGET", "SIZE", "FILES" -ForegroundColor White
    Write-Host "  $('─' * 58)" -ForegroundColor DarkGray

    foreach ($row in $sorted) {
        $nameDisplay = $row.Desc
        if ($nameDisplay.Length -gt 36) { $nameDisplay = $nameDisplay.Substring(0, 33) + "..." }
        $color = if ($row.Bytes -ge 100MB) { "Red" } elseif ($row.Bytes -ge 10MB) { "Yellow" } else { "White" }
        Write-Host ("  {0,-36} {1,>10} {2,>8}" -f $nameDisplay, (Format-Size $row.Bytes), $row.FileCount) -ForegroundColor $color
    }

    Write-Host "  $('─' * 58)" -ForegroundColor DarkGray
    Write-Host "  JUNK FILES TOTAL: $(Format-Size $previewTotalBytes)  ($previewTotalFiles files)" -ForegroundColor Cyan
} else {
    Write-Host "  No junk files found to clean." -ForegroundColor DarkGray
}

$queuedSteps = "1:Kill Procs → 2:Prune Apps → 3:Bloatware → 4:System Opt → 5:Junk → 6:Rhino → 7:WU Cleanup → 8:Explorer/Privacy"
if ($DeepClean) { $queuedSteps += " → 9:Deep Clean ⚠" }
Write-Host "`n  Steps: $queuedSteps" -ForegroundColor Gray

# --- DRY RUN EXIT ---
if ($global:IsDryRun) {
    Write-Host "`n  DRY RUN COMPLETE — no files were deleted." -ForegroundColor Yellow
    Write-Host "  Run without -DryRun to proceed with cleanup.`n" -ForegroundColor Gray
    Break
}

# --- CONFIRMATION ---
if (-not $Force) {
    Write-Host ""
    $confirm = Read-Host "  Proceed with cleanup? (y/N)"
    if ($confirm -ne "y") {
        Write-Host "  Cancelled." -ForegroundColor Yellow
        Break
    }
}

# =====================================================================
#  PHASE 2: EXECUTE CLEANUP
# =====================================================================

# Snapshot disk before
$DiskInfoBefore = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
$FreeSpaceBefore = $DiskInfoBefore.FreeSpace

# --- Custom Target ---
if ($HasCustomTarget) {
    Write-Host ""
    Write-Host "  Destroying custom target..." -ForegroundColor Red
    Set-ItemProperty -LiteralPath $CustomPath -Name Attributes -Value "Normal" -ErrorAction SilentlyContinue
    try {
        Remove-Item -LiteralPath $CustomPath -Recurse -Force -Confirm:$false -ErrorAction Stop
        Write-Host "  ✓ Destroyed. ($(Format-Size $customSize))" -ForegroundColor Green
        $global:CleanupReport.Add([PSCustomObject]@{
            Description = "Custom: $(Split-Path $CustomPath -Leaf)"
            FilesDeleted = 1; BytesDeleted = $customSize
            FilesFailed = 0; BytesFailed = 0; RebootQueued = 0
        })
        $global:TotalBytesDeleted += $customSize
        $global:TotalFilesDeleted++
    } catch {
        Write-Warning "  Locked. Queued for reboot."
        Register-DeleteOnReboot -FilePath $CustomPath
    }
}

# ------------------------------------------------------------------
# STEP 1: Kill Processes (expanded from unjunk.ps1)
# ------------------------------------------------------------------
Write-StepHeader "Stopping Background Processes" 1
$step = Start-StepTimer "Kill Processes"
$KillList = @(
    "ms-teams", "idman", "upc", "steam", "EpicGamesLauncher",
    "OneDrive", "Dropbox", "chrome", "msedge", "firefox",
    "discord", "cb_launcher", "PowerToys.Run", "upscayl"
)
$totalKilled = 0
foreach ($proc in $KillList) {
    $killed = Stop-TargetProcess $proc
    if ($killed -gt 0) { Write-Host "    Stopped: $proc ($killed)" -ForegroundColor Yellow }
    $totalKilled += $killed
}
if ($totalKilled -eq 0) { Write-Host "    No target processes running." -ForegroundColor DarkGray }
else { Write-Host "    Total: $totalKilled process(es) stopped." -ForegroundColor Green }
Stop-StepTimer $step

# ------------------------------------------------------------------
# STEP 2: Prune App Versions
# ------------------------------------------------------------------
Write-StepHeader "Pruning Old App Versions" 2
$step = Start-StepTimer "App Pruning"

$pruneListApps = @(
    "*Microsoft.DesktopAppInstaller*",     "*Microsoft.SecHealthUI*",
    "*NVIDIACorp.NVIDIAControlPanel*",     "*RealtekSemiconductorCorp.Realtek*",
    "*Microsoft.MicrosoftStickyNotes*",    "*Microsoft.WindowsTerminal*",
    "*Microsoft.WindowsNotepad*",          "*Microsoft.LanguageExperiencePack*",
    "*Microsoft.HEVCVideoExtension*",      "*Microsoft.HEIFImageExtension*",
    "*Microsoft.VP9VideoExtensions*",      "*Microsoft.WebMediaExtensions*",
    "*Microsoft.WebpImageExtension*",      "*Microsoft.MPEG2VideoExtension*",
    "*Microsoft.AVCEncoderVideoExtension*",
    "*Microsoft.WindowsStore*",            "*Microsoft.StorePurchaseApp*"
)
$pruneListFrameworks = @(
    "*Microsoft.WindowsAppRuntime*",       "*Microsoft.UI.Xaml*",
    "*Microsoft.NET.Native.Framework*",    "*Microsoft.NET.Native.Runtime*",
    "*Microsoft.VCLibs*",                  "*Microsoft.DirectXRuntime*"
)

$prunedCount  = 0
$skippedCount = 0

function Invoke-AppxPrune {
    param([string[]]$Patterns, [int]$KeepCount = 1)
    foreach ($appPattern in $Patterns) {
        $packages = Get-AppxPackage -Name $appPattern -AllUsers -PackageTypeFilter Main, Framework, Resource -ErrorAction SilentlyContinue
        if (-not $packages) { continue }
        $packages | Group-Object Name | ForEach-Object {
            $_.Group | Group-Object Architecture | ForEach-Object {
                $sorted = @($_.Group | Sort-Object { [version]$_.Version } -Descending)
                if ($sorted.Count -le $KeepCount) { return }
                $sorted | Select-Object -Skip $KeepCount | ForEach-Object {
                    $before = Get-AppxPackage -Name $_.Name -PackageTypeFilter Main, Framework, Resource -ErrorAction SilentlyContinue | Measure-Object
                    Write-Host "    Removing: $($_.Name) v$($_.Version) [$($_.Architecture)]" -ForegroundColor DarkGray
                    Remove-AppxPackage -Package $_.PackageFullName -AllUsers -Confirm:$false -ErrorAction SilentlyContinue
                    $after = Get-AppxPackage -Name $_.Name -PackageTypeFilter Main, Framework, Resource -ErrorAction SilentlyContinue | Measure-Object
                    if ($after.Count -lt $before.Count) { $script:prunedCount++ }
                    else { $script:skippedCount++; Write-Host "      ↳ Kept (dependency)" -ForegroundColor DarkYellow }
                }
            }
        }
    }
}

Invoke-AppxPrune -Patterns $pruneListApps -KeepCount 1
Invoke-AppxPrune -Patterns $pruneListFrameworks -KeepCount 2

# Provisioned
$allProvisioned = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
$allPrunePatterns = $pruneListApps + $pruneListFrameworks
foreach ($pattern in $allPrunePatterns) {
    $isFramework = $pattern -in $pruneListFrameworks
    $keepCount = if ($isFramework) { 2 } else { 1 }
    $matchedPackages = $allProvisioned | Where-Object { $_.DisplayName -like $pattern }
    if (-not $matchedPackages) { continue }
    $matchedPackages = $matchedPackages | Select-Object *, @{
        N = 'Architecture'; E = {
            if ($_.PackageName -match '_(?<arch>x64|x86|arm64|arm|neutral)_') { $Matches['arch'] } else { 'unknown' }
        }
    }
    $matchedPackages | Group-Object DisplayName | ForEach-Object {
        $_.Group | Group-Object Architecture | ForEach-Object {
            $sorted = @($_.Group | Sort-Object { [version]$_.Version } -Descending)
            if ($sorted.Count -le $keepCount) { return }
            $sorted | Select-Object -Skip $keepCount | ForEach-Object {
                Write-Host "    Removing Provisioned: $($_.DisplayName) v$($_.Version) [$($_.Architecture)]" -ForegroundColor DarkGray
                Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction SilentlyContinue
                $prunedCount++
            }
        }
    }
}

Write-Host "    Pruned: $prunedCount  |  Skipped: $skippedCount" -ForegroundColor Gray
Stop-StepTimer $step

# ------------------------------------------------------------------
# STEP 3: Bloatware
# ------------------------------------------------------------------
Write-StepHeader "Removing Bloatware" 3
$step = Start-StepTimer "Bloatware Removal"
$bloatApps = @(
    "*Clipchamp.Clipchamp*",              "*Microsoft.MixedReality.Portal*",
    "*Microsoft.XboxGamingOverlay*",      "*Microsoft.MicrosoftOfficeHub*",
    "*Microsoft.WinDbg*",                 "*Microsoft.GetHelp*",
    "*Microsoft.People*",                 "*Microsoft.GetStarted*",
    "*Microsoft.YourPhone*",              "*Microsoft.BingNews*",
    "*Microsoft.Todos*",                  "*Microsoft.Wallet*",
    "*Microsoft.PowerAutomateDesktop*",   "*Microsoft.Windows.DevHome*"
)
$bloatRemoved = 0
foreach ($app in $bloatApps) {
    $found = Get-AppxPackage $app -AllUsers -ErrorAction SilentlyContinue
    if ($found) {
        $found | Remove-AppxPackage -AllUsers -Confirm:$false -ErrorAction SilentlyContinue
        $bloatRemoved += @($found).Count
    }
    Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
        Where-Object { $_.PackageName -like $app } |
        Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
}
Write-Host "    Removed $bloatRemoved package(s)." -ForegroundColor Gray
Stop-StepTimer $step

# ------------------------------------------------------------------
# STEP 4: System Optimization
# ------------------------------------------------------------------
Write-StepHeader "System Optimization" 4
$step = Start-StepTimer "System Optimization"

# 4a. Delivery Optimization cache
Get-DeliveryOptimizationStatus | Remove-DeliveryOptimizationStatus -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "    Delivery Optimization cache cleared." -ForegroundColor Gray

# 4b. Built-in orphan cleanup
Start-Process -FilePath "rundll32.exe" -ArgumentList "AppxDeploymentClient.dll,AppxCleanupOrphanPackages" -Wait
Write-Host "    Built-in Appx orphan cleanup done." -ForegroundColor Gray

# 4c. Windows Defender — set scan purge to 1 day
Set-MpPreference -ScanPurgeItemsAfterDelay 1 -ErrorAction SilentlyContinue
Write-Host "    Defender scan purge set to 1 day." -ForegroundColor Gray

# 4d. DISM Component Cleanup
Write-Host "    Running DISM component cleanup (may take 1-3 min)..." -ForegroundColor Yellow
$dismProc = Start-Process -FilePath "dism.exe" `
    -ArgumentList "/Online /Cleanup-Image /StartComponentCleanup" `
    -PassThru -WindowStyle Hidden -RedirectStandardOutput "$env:TEMP\dism_cleanup.log"
$dismProc | Wait-Process -Timeout 300 -ErrorAction SilentlyContinue
if ($dismProc.HasExited -and $dismProc.ExitCode -eq 0) {
    Write-Host "    DISM component cleanup completed." -ForegroundColor Green
} elseif (-not $dismProc.HasExited) {
    Write-Host "    DISM timed out (5 min). Continuing — it may finish in background." -ForegroundColor Yellow
} else {
    Write-Host "    DISM returned exit code $($dismProc.ExitCode). Check $env:TEMP\dism_cleanup.log" -ForegroundColor Yellow
}

# 4e. WindowsApps Orphan Scanner
Write-Host "    Scanning WindowsApps for orphaned packages..." -ForegroundColor Yellow
$windowsAppsPath = "$env:ProgramFiles\WindowsApps"
$orphanBytesTotal = 0
$orphanFolders = [System.Collections.Generic.List[PSCustomObject]]::new()

try {
    $installedSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | ForEach-Object { $null = $installedSet.Add($_.PackageFullName) }
    Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | ForEach-Object { $null = $installedSet.Add($_.PackageName) }

    $appFolders = Get-ChildItem -LiteralPath $windowsAppsPath -Directory -ErrorAction Stop
    foreach ($folder in $appFolders) {
        if ($folder.Name -match '^(MutableBackup|MovedPackages|Deleted|\.staging)') { continue }
        if ($folder.Name -match '^Microsoft\.(NET|VCLibs|UI\.Xaml|Services\.Store)') { continue }
        if (-not $installedSet.Contains($folder.Name)) {
            $folderSize = 0
            try {
                $folderSize = (Get-ChildItem -LiteralPath $folder.FullName -Recurse -Force -File -ErrorAction SilentlyContinue |
                               Measure-Object -Property Length -Sum).Sum
            } catch { }
            if ($folderSize -gt 1MB) {
                $orphanFolders.Add([PSCustomObject]@{ Name = $folder.Name; Path = $folder.FullName; Bytes = $folderSize })
                $orphanBytesTotal += $folderSize
            }
        }
    }
} catch {
    Write-Host "    Could not enumerate WindowsApps (access denied)." -ForegroundColor DarkGray
}

if ($orphanFolders.Count -gt 0) {
    Write-Host "    Found $($orphanFolders.Count) orphaned package(s) totaling $(Format-Size $orphanBytesTotal):" -ForegroundColor Yellow
    $orphansRemoved = 0; $orphanBytesRemoved = 0
    foreach ($orphan in $orphanFolders) {
        try {
            $matchPkg = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object { $_.InstallLocation -eq $orphan.Path }
            if ($matchPkg) {
                $matchPkg | Remove-AppxPackage -AllUsers -Confirm:$false -ErrorAction Stop
                $orphansRemoved++; $orphanBytesRemoved += $orphan.Bytes; continue
            }
            $null = takeown.exe /F $orphan.Path /R /D Y 2>&1
            $null = icacls.exe $orphan.Path /grant "Administrators:F" /T /C /Q 2>&1
            Remove-Item -LiteralPath $orphan.Path -Recurse -Force -Confirm:$false -ErrorAction Stop
            $orphansRemoved++; $orphanBytesRemoved += $orphan.Bytes
        } catch {
            Write-Host "    Could not remove: $($orphan.Name)" -ForegroundColor DarkGray
        }
    }
    if ($orphansRemoved -gt 0) {
        Write-Host "    Removed $orphansRemoved orphan(s), freed $(Format-Size $orphanBytesRemoved)." -ForegroundColor Green
        $global:CleanupReport.Add([PSCustomObject]@{
            Description = "WindowsApps Orphans"
            FilesDeleted = $orphansRemoved; BytesDeleted = $orphanBytesRemoved
            FilesFailed = ($orphanFolders.Count - $orphansRemoved); BytesFailed = ($orphanBytesTotal - $orphanBytesRemoved)
            RebootQueued = 0
        })
        $global:TotalBytesDeleted += $orphanBytesRemoved
        $global:TotalFilesDeleted += $orphansRemoved
    }
} else {
    Write-Host "    No orphaned packages found." -ForegroundColor DarkGray
}

Stop-StepTimer $step

# ------------------------------------------------------------------
# STEP 5: Junk File Removal (pre-scanned + progress bar)
# ------------------------------------------------------------------
Write-StepHeader "Junk File Removal ($previewTotalFiles files, $(Format-Size $previewTotalBytes))" 5
$step = Start-StepTimer "Junk File Removal"

$groupIdx = 0
$groupTotal = ($previewTable | Where-Object { $_.Desc -notlike "Custom:*" }).Count
foreach ($entry in ($previewTable | Where-Object { $_.Desc -notlike "Custom:*" })) {
    $groupIdx++
    Write-Progress -Activity "Cleaning junk ($groupIdx/$groupTotal)" `
                   -Status "$($entry.Desc)  —  $(Format-Size $global:TotalBytesDeleted) recovered" `
                   -PercentComplete ([int]($groupIdx / [Math]::Max(1, $groupTotal) * 100)) -Id 0

    Remove-ScannedFiles -Desc $entry.Desc -Files $entry.Files -Path $entry.Path
}
Write-Progress -Activity "Cleaning junk" -Completed -Id 0

Stop-StepTimer $step

# ------------------------------------------------------------------
# STEP 6: Rhino Installer Scanner
#   Removes: Rhino + Rhino language packs
#   KEEPS:   Rhino.Inside (Rhino Inside Revit, etc.)
# ------------------------------------------------------------------
Write-StepHeader "Rhino Installer Scan (preserves Rhino.Inside)" 6
$step = Start-StepTimer "Rhino Scan"

$targetFolder = "C:\Windows\Installer"
$shell = $null
try {
    $shell = New-Object -ComObject Shell.Application
    $files = Get-ChildItem -Path $targetFolder -Recurse -File -ErrorAction SilentlyContinue
    $rhinoDeleted = 0; $rhinoBytes = 0; $rhinoSkipped = 0
    $fileTotal = @($files).Count; $fileIdx = 0

    foreach ($file in $files) {
        $fileIdx++
        if ($fileIdx % 50 -eq 0) {
            Write-Progress -Activity "Scanning Installer folder" -Status "$fileIdx / $fileTotal" -PercentComplete ([int]($fileIdx / [Math]::Max(1, $fileTotal) * 100)) -Id 3
        }
        try {
            $folder = $shell.Namespace($file.DirectoryName)
            if (-not $folder) { continue }
            $item = $folder.ParseName($file.Name)
            if (-not $item) { continue }
            $subject = $null
            foreach ($i in @(22, 11, 24)) {
                if ($folder.GetDetailsOf($folder.Items, $i) -eq "Subject") {
                    $subject = $folder.GetDetailsOf($item, $i); break
                }
            }
            if ($null -eq $subject) {
                for ($i = 0; $i -lt 266; $i++) {
                    if ($folder.GetDetailsOf($folder.Items, $i) -eq "Subject") {
                        $subject = $folder.GetDetailsOf($item, $i); break
                    }
                }
            }

            if ($subject -and $subject -match "Rhino") {
                # *** KEEP Rhino.Inside installations ***
                if ($subject -match "Rhino\.Inside") {
                    $rhinoSkipped++
                    continue
                }
                $size = $file.Length
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                $rhinoDeleted++; $rhinoBytes += $size
            }
        } catch { }
    }
    Write-Progress -Activity "Scanning Installer folder" -Completed -Id 3

    if ($rhinoDeleted -gt 0) {
        $global:CleanupReport.Add([PSCustomObject]@{
            Description = "Rhino Installers"
            FilesDeleted = $rhinoDeleted; BytesDeleted = $rhinoBytes
            FilesFailed = 0; BytesFailed = 0; RebootQueued = 0
        })
        $global:TotalBytesDeleted += $rhinoBytes
        $global:TotalFilesDeleted += $rhinoDeleted
        Write-Host "    Deleted $rhinoDeleted Rhino installer(s) ($(Format-Size $rhinoBytes))" -ForegroundColor Yellow
    } else {
        Write-Host "    No Rhino remnants found." -ForegroundColor DarkGray
    }
    if ($rhinoSkipped -gt 0) {
        Write-Host "    Preserved $rhinoSkipped Rhino.Inside file(s)." -ForegroundColor Green
    }
} finally {
    if ($shell) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null }
}
Stop-StepTimer $step

# ------------------------------------------------------------------
# STEP 7: Windows Update Cleanup
# ------------------------------------------------------------------
Write-StepHeader "Windows Update Cleanup" 7
$step = Start-StepTimer "WU Cleanup"

# Stop Windows Update service
Write-Host "    Stopping wuauserv..." -ForegroundColor DarkGray
Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# Clear SoftwareDistribution download cache
$wuDlPath = "C:\WINDOWS\SoftwareDistribution\Download"
if (Test-Path -LiteralPath $wuDlPath) {
    $wuSize = (Get-ChildItem -LiteralPath $wuDlPath -Recurse -Force -File -ErrorAction SilentlyContinue |
               Measure-Object -Property Length -Sum).Sum
    Remove-Item -LiteralPath $wuDlPath -Recurse -Force -ErrorAction SilentlyContinue
    if ($wuSize -gt 0) {
        $global:CleanupReport.Add([PSCustomObject]@{
            Description = "WU SoftwareDistribution"
            FilesDeleted = 1; BytesDeleted = $wuSize
            FilesFailed = 0; BytesFailed = 0; RebootQueued = 0
        })
        $global:TotalBytesDeleted += $wuSize
        $global:TotalFilesDeleted++
        Write-Host "    Cleared SoftwareDistribution ($(Format-Size $wuSize))" -ForegroundColor Yellow
    }
}

# Clear servicing LCU rollup
Remove-Item "C:\Windows\servicing\LCU\*.*" -Recurse -Force -ErrorAction SilentlyContinue

# Restart Windows Update service
Start-Service -Name wuauserv -ErrorAction SilentlyContinue
Write-Host "    wuauserv restarted." -ForegroundColor Gray

# Clear IE / legacy browser tracks
RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 1

Stop-StepTimer $step

# ------------------------------------------------------------------
# STEP 8: Windows Explorer & Privacy Cleanup
# ------------------------------------------------------------------
Write-StepHeader "Windows Explorer & Privacy" 8
$step = Start-StepTimer "Explorer & Privacy"

# Kill explorer for clean registry/file access
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

# Recent Items (history only — preserves pinned Quick Access & taskbar jump lists)
Write-Host "    Clearing recent items (preserving pins)..." -ForegroundColor DarkGray
Get-ChildItem "$env:APPDATA\Microsoft\Windows\Recent\*" -File -Force -Exclude desktop.ini -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue
# AutomaticDestinations: exclude f01b4d95cf55d32a = Quick Access pinned folders
Get-ChildItem "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\*" -File -Force `
    -Exclude desktop.ini, "f01b4d95cf55d32a.automaticDestinations-ms" -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue
# CustomDestinations: SKIP entirely — these are user-pinned taskbar jump list entries
Remove-Item "$env:APPDATA\Microsoft\Windows\Recent\*.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\History" -Recurse -Force -ErrorAction SilentlyContinue

# Registry MRU / History cleanup
Write-Host "    Clearing registry history keys..." -ForegroundColor DarkGray
$registryTargets = @(
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\WordWheelQuery"
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs"
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths"
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FeatureUsage"
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32"
    "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache"
    "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\BagMRU"
    "HKCU:\SOFTWARE\MPC-HC\MPC-HC\Recent File List"
    "HKCU:\Software\MPC-HC\MPC-HC\MediaHistory"
    "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Persisted"
    "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Store"
)
$regCleaned = 0
foreach ($regPath in $registryTargets) {
    if (Test-Path -LiteralPath $regPath) {
        Remove-Item -LiteralPath $regPath -Recurse -Force -ErrorAction SilentlyContinue
        $regCleaned++
    }
}

# AppCompatCache (Shim Cache) — system-level
foreach ($cs in @("ControlSet001", "CurrentControlSet")) {
    $shimPath = "HKLM:\SYSTEM\$cs\Control\Session Manager\AppCompatCache"
    if (Test-Path -LiteralPath $shimPath) {
        Remove-Item -LiteralPath $shimPath -Recurse -Force -ErrorAction SilentlyContinue
        $regCleaned++
    }
}

Write-Host "    Cleaned $regCleaned registry key(s)." -ForegroundColor Gray

# IDM download history (registry)
1..9 | ForEach-Object {
    Remove-Item "HKCU:\Software\DownloadManager\$_*" -Recurse -ErrorAction SilentlyContinue
}

# Restart explorer
Start-Process explorer.exe
Write-Host "    Explorer restarted." -ForegroundColor Green

Stop-StepTimer $step

# ------------------------------------------------------------------
# STEP 9: Deep Cleaning (gated by $DeepClean)
# ------------------------------------------------------------------
if ($DeepClean) {
    Write-StepHeader "Deep System Cleaning ⚠" 9
    $step = Start-StepTimer "Deep Cleaning"

    if ($PSCmdlet.ShouldProcess("Shadow Copies", "Delete")) {
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c vssadmin delete shadows /all /quiet" -Wait -WindowStyle Hidden
        Write-Host "    Shadow Copies deleted." -ForegroundColor Yellow
    }
    if ($PSCmdlet.ShouldProcess("CleanMgr", "Run Cleanup")) {
        $volumeCachesKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
        $cleanupCategories = @(
            "Active Setup Temp Folders",          "BranchCache"
            "D3D Shader Cache",                   "Delivery Optimization Files"
            "Device Driver Packages",             "Diagnostic Data Viewer database files"
            "Downloaded Program Files",           "Internet Cache Files"
            "Language Pack",                       "Old ChkDsk Files"
            "Previous Installations",             "Recycle Bin"
            "RetailDemo Offline Content",         "Service Pack Cleanup"
            "Setup Log Files",                     "System error memory dump files"
            "System error minidump files",        "Temporary Files"
            "Temporary Setup Files",              "Thumbnail Cache"
            "Update Cleanup",                     "Upgrade Discarded Files"
            "User file versions",                 "Windows Defender"
            "Windows Error Reporting Files",      "Windows Error Reporting Archive Files"
            "Windows Error Reporting Queue Files", "Windows Error Reporting System Archive Files"
            "Windows Error Reporting System Queue Files"
            "Windows ESD installation files",     "Windows Upgrade Log Files"
        )

        Write-Host "    Configuring Disk Cleanup profile..." -ForegroundColor DarkGray
        $configuredCount = 0
        foreach ($category in $cleanupCategories) {
            $catKey = Join-Path $volumeCachesKey $category
            if (Test-Path -LiteralPath $catKey) {
                Set-ItemProperty -LiteralPath $catKey -Name "StateFlags1221" -Value 2 -Type DWord -ErrorAction SilentlyContinue
                $configuredCount++
            }
        }
        Write-Host "    Enabled $configuredCount cleanup categories." -ForegroundColor DarkGray

        Write-Host "    Running Disk Cleanup (this may take a while)..." -ForegroundColor Yellow
        $cleanMgrProc = Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1221 /d C:" -PassThru
        try { $cleanMgrProc | Wait-Process -Timeout 10 -ErrorAction SilentlyContinue } catch { }
        $timeout = 600; $elapsed = 0
        while ($elapsed -lt $timeout) {
            Start-Sleep -Seconds 3; $elapsed += 3
            $workers = Get-Process -Name "cleanmgr", "DismHost" -ErrorAction SilentlyContinue
            if (-not $workers) { break }
            if ($elapsed % 15 -eq 0) { Write-Host "    Still cleaning... ($elapsed sec)" -ForegroundColor DarkGray }
        }
        if ($elapsed -ge $timeout) { Write-Host "    Disk Cleanup timed out after $timeout seconds." -ForegroundColor Yellow }
        else { Write-Host "    Disk Cleanup completed." -ForegroundColor Yellow }
    }
    if ($PSCmdlet.ShouldProcess("Event Logs", "Clear all")) {
        $logs = wevtutil.exe el
        $cleared = 0
        foreach ($log in $logs) { wevtutil.exe cl "$log" 2>$null; if ($LASTEXITCODE -eq 0) { $cleared++ } }
        Write-Host "    Event Logs cleared ($cleared)." -ForegroundColor Yellow
    }
    Stop-StepTimer $step
} else {
    Write-Host "`n  Step 9: Deep Cleaning skipped (CrowdStrike safe)." -ForegroundColor Green
}

# =====================================================================
#  PHASE 3: REPORT
# =====================================================================
$global:ScriptStopwatch.Stop()
$DiskInfoAfter  = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
$FreeSpaceAfter = $DiskInfoAfter.FreeSpace
$Recovered      = $FreeSpaceAfter - $FreeSpaceBefore

# Aggregate report
$AggregatedReport = $global:CleanupReport |
    Group-Object Description |
    Select-Object @{N='Description'; E={$_.Name}},
                  @{N='BytesDeleted'; E={($_.Group | Measure-Object BytesDeleted -Sum).Sum}},
                  @{N='FilesDeleted'; E={($_.Group | Measure-Object FilesDeleted -Sum).Sum}},
                  @{N='BytesFailed';  E={($_.Group | Measure-Object BytesFailed  -Sum).Sum}},
                  @{N='FilesFailed';  E={($_.Group | Measure-Object FilesFailed  -Sum).Sum}},
                  @{N='RebootQueued'; E={($_.Group | Measure-Object RebootQueued -Sum).Sum}} |
    Sort-Object BytesDeleted -Descending

# --- Build report string (console + log file) ---
$reportLines = [System.Collections.Generic.List[string]]::new()

function Add-ReportLine {
    param([string]$Line = "", [string]$Color = "White", [switch]$NoConsole)
    $reportLines.Add($Line)
    if (-not $NoConsole) { Write-Host $Line -ForegroundColor $Color }
}

Add-ReportLine "" -NoConsole
Add-ReportLine "╔══════════════════════════════════════════════════════════════════════╗" -Color Green
Add-ReportLine "║                         CLEANUP REPORT                             ║" -Color Green
Add-ReportLine "╠══════════════════════════════════════════════════════════════════════╣" -Color Green
Add-ReportLine ""
Add-ReportLine ("  {0,-36} {1,>12} {2,>8} {3,>8}" -f "LOCATION", "DELETED", "FILES", "FAILED") -Color Gray
Add-ReportLine "  $('─' * 68)" -Color DarkGray

foreach ($row in $AggregatedReport) {
    $failStr = if ($row.FilesFailed -gt 0) { "$($row.FilesFailed) !!" } else { "-" }
    $color = if ($row.FilesFailed -gt 0) { "Yellow" } else { "White" }
    $name = $row.Description
    if ($name.Length -gt 36) { $name = $name.Substring(0, 33) + "..." }
    Add-ReportLine ("  {0,-36} {1,>12} {2,>8} {3,>8}" -f $name, (Format-Size $row.BytesDeleted), $row.FilesDeleted, $failStr) -Color $color
}

if ($AggregatedReport.Count -eq 0) {
    Add-ReportLine "  (No junk files found — system is clean!)" -Color DarkGray
}

Add-ReportLine ""
Add-ReportLine "  $('═' * 68)" -Color Green
Add-ReportLine "  TOTAL DELETED:       $(Format-Size $global:TotalBytesDeleted)  ($($global:TotalFilesDeleted) files)" -Color Cyan
if ($global:TotalFilesFailed -gt 0) {
    Add-ReportLine "  FAILED (LOCKED):     $(Format-Size $global:TotalBytesFailed)  ($($global:TotalFilesFailed) files)" -Color Yellow
}
if ($global:TotalRebootQueued -gt 0) {
    Add-ReportLine "  QUEUED FOR REBOOT:   $($global:TotalRebootQueued) file(s)" -Color Magenta
}

Add-ReportLine ""
Add-ReportLine "  $('─' * 68)" -Color DarkGray
Add-ReportLine "  C: FREE BEFORE:      $(Format-Size $FreeSpaceBefore)" -Color Gray
Add-ReportLine "  C: FREE AFTER:       $(Format-Size $FreeSpaceAfter)" -Color Gray
$deltaSign = if ($Recovered -ge 0) { "+" } else { "" }
Add-ReportLine "  NET CHANGE:          $deltaSign$(Format-Size $Recovered)" -Color $(if ($Recovered -ge 0) { "Green" } else { "Red" })

if ($global:RebootQueue.Count -gt 0) {
    Add-ReportLine ""
    Add-ReportLine "  $('─' * 68)" -Color DarkGray
    Add-ReportLine "  FILES QUEUED FOR DELETION ON NEXT REBOOT:" -Color Magenta
    foreach ($qf in $global:RebootQueue) { Add-ReportLine "    * $qf" -Color DarkMagenta }
    Add-ReportLine "  !! A reboot is required to complete cleanup." -Color Yellow
}

Add-ReportLine ""
Add-ReportLine "  $('─' * 68)" -Color DarkGray
Add-ReportLine "  STEP TIMINGS:" -Color Gray
foreach ($t in $global:StepTimings) {
    Add-ReportLine ("    {0,-30} {1,>10}" -f $t.Name, (Format-Elapsed $t.Elapsed)) -Color DarkGray
}
Add-ReportLine ("    {0,-30} {1,>10}" -f "TOTAL RUNTIME", (Format-Elapsed $global:ScriptStopwatch.Elapsed)) -Color White

if ($AggregatedReport.Count -gt 0) {
    Add-ReportLine ""
    Add-ReportLine "  $('─' * 68)" -Color DarkGray
    Add-ReportLine "  TOP 5 SPACE CONSUMERS:" -Color Gray
    $rank = 1
    foreach ($item in ($AggregatedReport | Select-Object -First 5)) {
        $pct = if ($global:TotalBytesDeleted -gt 0) { [math]::Round(($item.BytesDeleted / $global:TotalBytesDeleted) * 100, 1) } else { 0 }
        $barLen = [math]::Max(1, [math]::Round($pct / 5))
        $bar = ([char]0x2588).ToString() * $barLen + ([char]0x2591).ToString() * (20 - $barLen)
        Add-ReportLine "    $rank. $bar $pct%  $($item.Description) ($(Format-Size $item.BytesDeleted))" -Color Cyan
        $rank++
    }
}

Add-ReportLine ""
Add-ReportLine "╚══════════════════════════════════════════════════════════════════════╝" -Color Green

# =====================================================================
#  AUTO-SAVE LOG FILE
# =====================================================================
$logDir = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "MaintenanceLogs")
if (-not (Test-Path -LiteralPath $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$logPath = Join-Path $logDir "Maintenance_$timestamp.log"

$logHeader = @(
    "Master System Maintenance v38 — Log"
    "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    "User: $env:USERNAME@$env:COMPUTERNAME"
    "Steps Run: $stepsLabel"
    "Junk Targets: $($JunkTargets.Count)"
    ""
)
($logHeader + $reportLines) | Out-File -FilePath $logPath -Encoding UTF8

Write-Host ""
Write-Host "  Log saved: $logPath" -ForegroundColor Green
Write-Host ""
Write-Host "  Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")