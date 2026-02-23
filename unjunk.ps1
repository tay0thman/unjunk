<#
.SYNOPSIS
    Master System Maintenance Script (v37 - Interactive UX)
.DESCRIPTION
    v37 UX overhaul:
    1. DRY-RUN PREVIEW: Scans all targets, shows a size summary table,
       and asks for confirmation BEFORE deleting anything.
    2. PROGRESS BARS: Write-Progress with live bytes-recovered counter during
       deletion. Replaces the scrolling text wall.
    3. DEEP CLEAN TOGGLE: Single y/N prompt for risky operations (shadow
       copies, event logs, cleanmgr) that may trigger CrowdStrike Falcon.
       Steps 1-6 always run.
    4. AUTO-SAVE LOG: Full report is saved as a timestamped .log file on the
       Desktop after completion.
    
    Inherits all v36 fixes (wildcard resolution, accurate size tracking,
    framework-safe pruning, COM cleanup, bottom-up directory removal).

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
$Host.UI.RawUI.WindowTitle = "Master System Maintenance v37"

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

# Preview-mode scan results (populated during dry-run scan, used as delete manifest)
$global:PreviewResults    = [System.Collections.Generic.List[PSCustomObject]]::new()

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
        [string]$Path   # original path for directory cleanup
    )
    if ($Files.Count -eq 0) { return }

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
                       -PercentComplete $pct `
                       -Id 1

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
#  JUNK PATH DEFINITIONS (data-driven: description + path pairs)
# =====================================================================

# Build Revit year-based paths dynamically
$RevitPaths = @()
2018..2030 | ForEach-Object {
    $RevitPaths += @{ Desc = "Revit $_ Collab";    Path = "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit $_\CollaborationCache\*" }
    $RevitPaths += @{ Desc = "Revit $_ Journals";  Path = "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit $_\Journals\*" }
}

$JunkTargets = @(
    # --- System Temps & Crash Data ---
    @{ Desc = "System Temp";         Path = "$env:SYSTEMROOT\Temp\*" }
    @{ Desc = "User Temp";           Path = "$env:TEMP\*" }
    @{ Desc = "Local Temp";          Path = "$env:LOCALAPPDATA\Temp\*" }
    @{ Desc = "Prefetch";            Path = "$env:SYSTEMROOT\Prefetch\*" }
    @{ Desc = "Live Kernel Reports"; Path = "$env:SYSTEMROOT\LiveKernelReports\*" }
    @{ Desc = "Minidumps";           Path = "$env:SYSTEMROOT\Minidump\*" }
    @{ Desc = "Memory Dump";         Path = "$env:SYSTEMROOT\MEMORY.DMP" }
    @{ Desc = "Crash Dumps";         Path = "$env:LOCALAPPDATA\CrashDumps\*" }
    @{ Desc = "WER Archives";        Path = "C:\ProgramData\Microsoft\Windows\WER\ReportArchive\*" }
    @{ Desc = "WER Queue";           Path = "C:\ProgramData\Microsoft\Windows\WER\ReportQueue\*" }

    # --- Browsers ---
    @{ Desc = "Chrome Cache";        Path = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache\*" }
    @{ Desc = "Edge Cache";          Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache\Cache_Data\*" }
    @{ Desc = "Windows WebCache";    Path = "$env:LOCALAPPDATA\Microsoft\Windows\WebCache\*" }
    @{ Desc = "Edge Service Workers"; Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Service Worker\CacheStorage\*" }
    @{ Desc = "Edge ScriptCache";    Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Service Worker\ScriptCache\*" }
    @{ Desc = "Edge IndexedDB";      Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\IndexedDB\*" }
    @{ Desc = "Edge Code Cache";     Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache\*" }

    # --- Installers ---
    @{ Desc = "Autodesk Install Root"; Path = "C:\Autodesk" }
    @{ Desc = "Adobe Temp Root";     Path = "C:\adobeTemp" }
    @{ Desc = "WinRE Agent";         Path = "C:\`$WinREAgent" }

    # --- App Data ---
    @{ Desc = "IDM Download Data";   Path = "$env:APPDATA\IDM\DwnlData\*" }
    @{ Desc = "Ubisoft Cache";       Path = "C:\Program Files (x86)\Ubisoft\Ubisoft Game Launcher\cache\*" }
    @{ Desc = "Edge Update Installers"; Path = "C:\Program Files (x86)\Microsoft\EdgeUpdate\Download\*" }
    @{ Desc = "Steam Logs";          Path = "C:\Program Files (x86)\Steam\logs\*" }
    @{ Desc = "Adobe Installer Logs"; Path = "C:\Program Files (x86)\Common Files\Adobe\Installers\*.log" }
    @{ Desc = "VS Installer Packages"; Path = "C:\ProgramData\Microsoft\VisualStudio\Packages\*" }
    @{ Desc = "USO Logs";            Path = "C:\ProgramData\USOShared\Logs\*" }
    @{ Desc = "Upscayl Cache";       Path = "$env:LOCALAPPDATA\upscayl-updater\*" }
    @{ Desc = "PowerToys Updates";   Path = "$env:LOCALAPPDATA\Microsoft\PowerToys\Updates\*" }
    @{ Desc = "UniGetUI Cache";      Path = "$env:LOCALAPPDATA\UniGetUI\CachedMedia\*" }
    @{ Desc = "Chaos Cosmos Updates"; Path = "$env:LOCALAPPDATA\Chaos\Cosmos\Updates\*" }
    @{ Desc = "Chaos Vantage Cache"; Path = "$env:LOCALAPPDATA\Chaos Group\Vantage\cache\*" }
    @{ Desc = "Rhino User Update Cache"; Path = "$env:LOCALAPPDATA\McNeel\McNeelUpdate\DownloadCache\*" }
    @{ Desc = "Rhino System Update Cache"; Path = "C:\ProgramData\McNeel\McNeelUpdate\DownloadCache\*" }
    @{ Desc = "Maxon Redshift Cache"; Path = "$env:APPDATA\Maxon\*\Redshift\Cache\*" }
    @{ Desc = "Maxon Redshift Textures"; Path = "$env:APPDATA\Maxon\*\Redshift\Cache\Textures\*" }
    @{ Desc = "Cinebench Cache";     Path = "$env:APPDATA\Maxon\Cinebench*\cache\*" }
    @{ Desc = "Maxon Assets Cache";  Path = "$env:APPDATA\Maxon\*\assets\*" }
    @{ Desc = "InDesign Caches";     Path = "$env:LOCALAPPDATA\Adobe\InDesign\*\*\Caches\*" }
    @{ Desc = "Adobe Media Cache";   Path = "$env:APPDATA\Adobe\Common\Media Cache Files\*" }
    @{ Desc = "Lightroom Cache";     Path = "$env:LOCALAPPDATA\Adobe\Lightroom\Caches\*" }
    @{ Desc = "Gameloft Cache";      Path = "$env:LOCALAPPDATA\Gameloft\*\Cache\*" }
    @{ Desc = "Bluebeam Logs";       Path = "$env:LOCALAPPDATA\Bluebeam\Revu\*\Logs\*" }
    @{ Desc = "NuGet Packages";      Path = "$env:USERPROFILE\.nuget\packages\*" }
    @{ Desc = "Discord Cache";       Path = "$env:APPDATA\discord\Cache\*" }
    @{ Desc = "Discord Code Cache";  Path = "$env:APPDATA\discord\Code Cache\*" }
    @{ Desc = "Discord GPUCache";    Path = "$env:APPDATA\discord\GPUCache\*" }
    @{ Desc = "Revit PacCache";      Path = "$env:LOCALAPPDATA\Autodesk\Revit\PacCache\*" }
    @{ Desc = "ACCDocs";             Path = "$env:HOMEPATH\ACCDOCS\*" }
    @{ Desc = "Logitech GHub Cache"; Path = "C:\ProgramData\LGHUB\cache\*" }
    @{ Desc = "Zoom Logs";           Path = "$env:APPDATA\Zoom\logs\*" }
    @{ Desc = "NVIDIA DX Cache";     Path = "$env:LOCALAPPDATA\NVIDIA\DXCache\*" }
    @{ Desc = "Steam Web Cache";     Path = "$env:LOCALAPPDATA\Steam\htmlcache\*" }
    @{ Desc = "InstallShield Leftovers"; Path = "C:\Program Files (x86)\InstallShield Installation Information\*" }
) + $RevitPaths

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

  Master Maintenance v37
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

$stepsLabel = "Steps 1-6 (standard)"
if ($DeepClean) { $stepsLabel += " + Step 7 (deep clean)" }

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
    if ($fileCount -gt 0) {
        $previewTable.Add([PSCustomObject]@{
            Desc      = $target.Desc
            Path      = $target.Path
            FileCount = $fileCount
            Bytes     = $fileBytes
            Files     = $files   # carry forward for deletion phase
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
        Files     = $null  # handled separately
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

# Show what will run
$queuedSteps = "Kill Processes → Prune Apps → Remove Bloatware → System Optimization → Junk Removal → Rhino Scan"
if ($DeepClean) { $queuedSteps += " → Deep Cleaning ⚠" }
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
# STEP 1: Kill Processes
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
    "*Microsoft.AVCEncoderVideoExtension*"
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
    $matches = $allProvisioned | Where-Object { $_.DisplayName -like $pattern }
    if (-not $matches) { continue }
    $matchesWithArch = $matches | Select-Object *, @{
        N = 'Architecture'; E = {
            if ($_.PackageName -match '_(?<arch>x64|x86|arm64|arm|neutral)_') { $Matches['arch'] } else { 'unknown' }
        }
    }
    $matchesWithArch | Group-Object DisplayName | ForEach-Object {
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

# 4b. Built-in orphan cleanup (limited but fast)
Start-Process -FilePath "rundll32.exe" -ArgumentList "AppxDeploymentClient.dll,AppxCleanupOrphanPackages" -Wait
Write-Host "    Built-in Appx orphan cleanup done." -ForegroundColor Gray

# 4c. DISM Component Cleanup (removes superseded component versions from WinSxS)
#     This is the single highest-impact system size reducer. Safe — only removes
#     components that have been fully replaced by newer versions.
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

# 4d. WindowsApps Orphan Scanner
#     Cross-references folder contents against registered packages.
#     Folders with no matching installation record are orphans.
Write-Host "    Scanning WindowsApps for orphaned packages..." -ForegroundColor Yellow

$windowsAppsPath = "$env:ProgramFiles\WindowsApps"
$orphanBytesTotal = 0
$orphanFolders = [System.Collections.Generic.List[PSCustomObject]]::new()

try {
    # Build lookup sets of all legitimately installed package folder names
    $installedSet = [System.Collections.Generic.HashSet[string]]::new(
        [StringComparer]::OrdinalIgnoreCase
    )
    Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | ForEach-Object {
        $null = $installedSet.Add($_.PackageFullName)
    }
    # Provisioned packages use a slightly different naming — add those too
    Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | ForEach-Object {
        $null = $installedSet.Add($_.PackageName)
    }

    # Enumerate actual folders in WindowsApps
    # This requires admin; folder ACL allows SYSTEM + TrustedInstaller by default.
    # We read via Get-ChildItem which works with admin rights on most builds.
    $appFolders = Get-ChildItem -LiteralPath $windowsAppsPath -Directory -ErrorAction Stop

    foreach ($folder in $appFolders) {
        # Skip known system folders
        if ($folder.Name -match '^(MutableBackup|MovedPackages|Deleted|\.staging)') { continue }
        # Skip Microsoft runtime/framework infrastructure folders
        if ($folder.Name -match '^Microsoft\.(NET|VCLibs|UI\.Xaml|Services\.Store)') { continue }

        if (-not $installedSet.Contains($folder.Name)) {
            # Not registered — potential orphan. Measure its size.
            $folderSize = 0
            try {
                $folderSize = (Get-ChildItem -LiteralPath $folder.FullName -Recurse -Force -File -ErrorAction SilentlyContinue |
                               Measure-Object -Property Length -Sum).Sum
            } catch { }

            # Only flag folders > 1 MB to avoid noise from tiny metadata remnants
            if ($folderSize -gt 1MB) {
                $orphanFolders.Add([PSCustomObject]@{
                    Name = $folder.Name
                    Path = $folder.FullName
                    Bytes = $folderSize
                })
                $orphanBytesTotal += $folderSize
            }
        }
    }
} catch {
    Write-Host "    Could not enumerate WindowsApps (access denied or not found)." -ForegroundColor DarkGray
}

if ($orphanFolders.Count -gt 0) {
    Write-Host ""
    Write-Host "    Found $($orphanFolders.Count) orphaned package(s) totaling $(Format-Size $orphanBytesTotal):" -ForegroundColor Yellow
    Write-Host "    $('─' * 60)" -ForegroundColor DarkGray

    foreach ($orphan in ($orphanFolders | Sort-Object Bytes -Descending)) {
        # Truncate long package names for display
        $displayName = $orphan.Name
        if ($displayName.Length -gt 55) { $displayName = $displayName.Substring(0, 52) + "..." }
        Write-Host ("    {0,-55} {1,>10}" -f $displayName, (Format-Size $orphan.Bytes)) -ForegroundColor White
    }

    Write-Host ""
    Write-Host "    Removing orphaned packages..." -ForegroundColor Yellow
    $orphansRemoved = 0; $orphanBytesRemoved = 0

    foreach ($orphan in $orphanFolders) {
        try {
            # First try: ask the package manager to remove it properly
            $matchPkg = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue |
                        Where-Object { $_.InstallLocation -eq $orphan.Path }
            if ($matchPkg) {
                $matchPkg | Remove-AppxPackage -AllUsers -Confirm:$false -ErrorAction Stop
                $orphansRemoved++; $orphanBytesRemoved += $orphan.Bytes
                continue
            }

            # Second try: take ownership and force-remove the folder
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
            Description  = "WindowsApps Orphans"
            FilesDeleted = $orphansRemoved; BytesDeleted = $orphanBytesRemoved
            FilesFailed  = ($orphanFolders.Count - $orphansRemoved); BytesFailed = ($orphanBytesTotal - $orphanBytesRemoved)
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
# STEP 5: Junk File Removal (using pre-scanned results + progress bar)
# ------------------------------------------------------------------
Write-StepHeader "Junk File Removal ($previewTotalFiles files, $(Format-Size $previewTotalBytes))" 5
$step = Start-StepTimer "Junk File Removal"

$groupIdx = 0
$groupTotal = $previewTable.Count
foreach ($entry in ($previewTable | Where-Object { $_.Desc -notlike "Custom:*" })) {
    $groupIdx++
    Write-Progress -Activity "Cleaning junk ($groupIdx/$groupTotal)" `
                   -Status "$($entry.Desc)  —  $(Format-Size $global:TotalBytesDeleted) recovered" `
                   -PercentComplete ([int]($groupIdx / $groupTotal * 100)) `
                   -Id 0

    Remove-ScannedFiles -Desc $entry.Desc -Files $entry.Files -Path $entry.Path
}
Write-Progress -Activity "Cleaning junk" -Completed -Id 0

Stop-StepTimer $step

# ------------------------------------------------------------------
# STEP 6: Rhino Installer Scanner
# ------------------------------------------------------------------
Write-StepHeader "Rhino Installer Scan" 6
$step = Start-StepTimer "Rhino Scan"

$targetFolder = "C:\Windows\Installer"
$shell = $null
try {
    $shell = New-Object -ComObject Shell.Application
    $files = Get-ChildItem -Path $targetFolder -Recurse -File -ErrorAction SilentlyContinue
    $rhinoDeleted = 0; $rhinoBytes = 0
    $fileTotal = @($files).Count; $fileIdx = 0

    foreach ($file in $files) {
        $fileIdx++
        if ($fileIdx % 50 -eq 0) {
            Write-Progress -Activity "Scanning Installer folder" -Status "$fileIdx / $fileTotal" -PercentComplete ([int]($fileIdx / $fileTotal * 100)) -Id 3
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
} finally {
    if ($shell) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null }
}
Stop-StepTimer $step

# ------------------------------------------------------------------
# STEP 7: Deep Cleaning
# ------------------------------------------------------------------
if ($DeepClean) {
    Write-StepHeader "Deep System Cleaning ⚠" 7
    $step = Start-StepTimer "Deep Cleaning"

    if ($PSCmdlet.ShouldProcess("Shadow Copies", "Delete")) {
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c vssadmin delete shadows /all /quiet" -Wait -WindowStyle Hidden
        Write-Host "    Shadow Copies deleted." -ForegroundColor Yellow
    }
    if ($PSCmdlet.ShouldProcess("CleanMgr", "Run Cleanup")) {
        # Configure cleanup categories for sagerun profile 1221.
        # Without this, /sagerun:1221 exits instantly (no profile = nothing to do).
        $volumeCachesKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
        $cleanupCategories = @(
            "Active Setup Temp Folders"
            "BranchCache"
            "D3D Shader Cache"
            "Delivery Optimization Files"
            "Device Driver Packages"
            "Diagnostic Data Viewer database files"
            "Downloaded Program Files"
            "Internet Cache Files"
            "Language Pack"
            "Old ChkDsk Files"
            "Previous Installations"
            "Recycle Bin"
            "RetailDemo Offline Content"
            "Service Pack Cleanup"
            "Setup Log Files"
            "System error memory dump files"
            "System error minidump files"
            "Temporary Files"
            "Temporary Setup Files"
            "Thumbnail Cache"
            "Update Cleanup"
            "Upgrade Discarded Files"
            "User file versions"
            "Windows Defender"
            "Windows Error Reporting Files"
            "Windows ESD installation files"
            "Windows Upgrade Log Files"
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
        $cleanMgrProc = Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1221" -PassThru
        # cleanmgr spawns a child process and the parent exits immediately.
        # Wait for the actual cleanup window (DismHost or cleanmgr) to finish.
        try {
            $cleanMgrProc | Wait-Process -Timeout 10 -ErrorAction SilentlyContinue
        } catch { }
        # Now wait for the real worker process
        $timeout = 600  # 10 minute max
        $elapsed = 0
        while ($elapsed -lt $timeout) {
            Start-Sleep -Seconds 3
            $elapsed += 3
            $workers = Get-Process -Name "cleanmgr", "DismHost" -ErrorAction SilentlyContinue
            if (-not $workers) { break }
            if ($elapsed % 15 -eq 0) {
                Write-Host "    Still cleaning... ($elapsed sec)" -ForegroundColor DarkGray
            }
        }
        if ($elapsed -ge $timeout) {
            Write-Host "    Disk Cleanup timed out after $timeout seconds." -ForegroundColor Yellow
        } else {
            Write-Host "    Disk Cleanup completed." -ForegroundColor Yellow
        }
    }
    if ($PSCmdlet.ShouldProcess("Event Logs", "Clear all")) {
        $logs = wevtutil.exe el
        $cleared = 0
        foreach ($log in $logs) { wevtutil.exe cl "$log" 2>$null; if ($LASTEXITCODE -eq 0) { $cleared++ } }
        Write-Host "    Event Logs cleared ($cleared)." -ForegroundColor Yellow
    }
    Stop-StepTimer $step
} else {
    Write-Host "`n  Step 7: Deep Cleaning skipped (CrowdStrike safe)." -ForegroundColor Green
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

# --- Build report string (used for both console and log file) ---
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

# Reboot queue
if ($global:RebootQueue.Count -gt 0) {
    Add-ReportLine ""
    Add-ReportLine "  $('─' * 68)" -Color DarkGray
    Add-ReportLine "  FILES QUEUED FOR DELETION ON NEXT REBOOT:" -Color Magenta
    foreach ($qf in $global:RebootQueue) { Add-ReportLine "    * $qf" -Color DarkMagenta }
    Add-ReportLine "  !! A reboot is required to complete cleanup." -Color Yellow
}

# Step timings
Add-ReportLine ""
Add-ReportLine "  $('─' * 68)" -Color DarkGray
Add-ReportLine "  STEP TIMINGS:" -Color Gray
foreach ($t in $global:StepTimings) {
    Add-ReportLine ("    {0,-30} {1,>10}" -f $t.Name, (Format-Elapsed $t.Elapsed)) -Color DarkGray
}
Add-ReportLine ("    {0,-30} {1,>10}" -f "TOTAL RUNTIME", (Format-Elapsed $global:ScriptStopwatch.Elapsed)) -Color White

# Top 5
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
    "Master System Maintenance v37 — Log"
    "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    "User: $env:USERNAME@$env:COMPUTERNAME"
    "Steps Run: $stepsLabel"
    ""
)
($logHeader + $reportLines) | Out-File -FilePath $logPath -Encoding UTF8

Write-Host ""
Write-Host "  Log saved: $logPath" -ForegroundColor Green
Write-Host ""
Write-Host "  Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
