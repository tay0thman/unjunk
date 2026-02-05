<#
.SYNOPSIS
    Master System Maintenance Script (v34 - Silent / Yes to All)
.DESCRIPTION
    1. Interactive Toggle for "Deep Clean" (Shadow Copies/Logs).
    2. Interactive "Custom File Destroyer" with Reboot Support.
    3. SILENT EXECUTION: Removes all individual "Are you sure?" prompts.
    4. NEW: Queues locked files for deletion on next REBOOT (MoveFileEx).
    5. Smart-Prunes Apps, Cleans Junk, Installer Caches (Autodesk/Adobe).
.PARAMETER Force
    Skips confirmation prompts.
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param()

# --- CONFIGURATION ---
$ErrorActionPreference = "SilentlyContinue"
$Host.UI.RawUI.WindowTitle = "Master System Maintenance v34"

# --- ADMIN CHECK ---
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Warning "You must run this script as Administrator!"
    Write-Warning "Right-click PowerShell and select 'Run as Administrator'."
    Break
}

# --- NATIVE API WRAPPER FOR MOVEFILEEX (DELETE ON REBOOT) ---
$MoveFileCode = @'
    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, int dwFlags);
'@
try {
    $Kernel32 = Add-Type -MemberDefinition $MoveFileCode -Name "Kernel32Helper" -Namespace "Win32" -PassThru
} catch {
    # Type already exists from previous run, ignore error
}

function Register-DeleteOnReboot {
    param([string]$FilePath)
    # 0x4 = MOVEFILE_DELAY_UNTIL_REBOOT
    try {
        [Win32.Kernel32Helper]::MoveFileEx($FilePath, $null, 4)
    } catch {
        Write-Warning "Failed to register $FilePath for reboot deletion."
    }
}

# --- HELPER FUNCTIONS ---

function Stop-TargetProcess {
    param([string]$ProcessName)
    if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) {
        Write-Host "Stopping process: $ProcessName" -ForegroundColor Yellow
        Stop-Process -Name $ProcessName -Force -Confirm:$false -ErrorAction SilentlyContinue
    }
}

function Remove-JunkPath {
    param([string]$Path, [string]$Desc)
    $ExpandedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    
    if (Test-Path $ExpandedPath) {
        # PROMPT REMOVED: Auto-confirming deletion
        Write-Host "Cleaning $Desc..." -ForegroundColor Cyan
        
        # Get all items, including hidden
        $items = Get-ChildItem -Path $ExpandedPath -Recurse -Force -ErrorAction SilentlyContinue
        
        foreach ($item in $items) {
            try {
                # Added -Confirm:$false to suppress file-level prompts
                $item | Remove-Item -Recurse -Force -Confirm:$false -ErrorAction Stop
            } catch {
                # IF DELETE FAILS: Queue for Reboot
                # Only show warning for files, not folders (to reduce noise)
                if (-not $item.PSIsContainer) {
                    Write-Warning "Locked: $($item.Name). Queueing for deletion on REBOOT."
                    Register-DeleteOnReboot -FilePath $item.FullName
                }
            }
        }
        
        # Try to remove the root folder itself if empty now
        try { Remove-Item -Path $ExpandedPath -Force -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    }
}

# Improved MSI Reader using WindowsInstaller COM Object
function Get-MsiProductName {
    param ([string]$filePath)
    try {
        $wi = New-Object -ComObject WindowsInstaller.Installer
        # 0 = ReadOnly Mode
        $db = $wi.OpenDatabase($filePath, 0) 
        $view = $db.OpenView("SELECT Value FROM Property WHERE Property = 'ProductName'")
        $view.Execute()
        $record = $view.Fetch()
        if ($record) { 
            return $record.StringData(1) 
        }
        $view.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($db) | Out-Null
    } catch {
        return $null
    }
    return $null
}

# --- INTERACTIVE MENU ---
Clear-Host
Write-Host "
   __  __          _____ _______ ______ _____  
  |  \/  |   /\   / ____|__   __|  ____|  __ \ 
  | \  / |  /  \ | (___    | |  | |__  | |__) |
  | |\/| | / /\ \ \___ \   | |  |  __| |  _  / 
  | |  | |/ ____ \____) |  | |  | |____| | \ \ 
  |_|  |_/_/    \_\_____/  |_|  |______|_|  \_\
                                               
  Master Maintenance (v34 - Silent Mode)
" -ForegroundColor Green

# 1. CUSTOM FILE DESTROYER
Write-Host "--- 1. CUSTOM FILE DESTROYER ---" -ForegroundColor Cyan
Write-Host "Paste a path to force delete (or press Enter to skip)." -ForegroundColor Gray
$CustomPath = Read-Host "Path"
if ($CustomPath -and (Test-Path $CustomPath)) {
    Write-Host "Target Acquired: $CustomPath" -ForegroundColor Red
    Set-ItemProperty -Path $CustomPath -Name Attributes -Value "Normal" -ErrorAction SilentlyContinue
    
    try {
        Remove-Item -Path $CustomPath -Recurse -Force -Confirm:$false -ErrorAction Stop
        Write-Host "Target Destroyed." -ForegroundColor Green
    } catch {
        Write-Warning "File is locked. Scheduling destruction for NEXT REBOOT."
        Register-DeleteOnReboot -FilePath $CustomPath
    }
}

# 2. DEEP CLEAN TOGGLE
Write-Host "`n--- 2. DEEP SYSTEM CLEANING (CrowdStrike Alert Risk) ---" -ForegroundColor Cyan
Write-Host "Includes: Shadow Copies, Event Logs, Disk Cleanup." -ForegroundColor Gray
$DeepResponse = Read-Host "Enable Deep Cleaning? (y/N)"
if ($DeepResponse -eq "y") { $DeepClean = $true } else { $DeepClean = $false }

# --- START AUTOMATED CLEANUP ---

# Measure Disk Space
$DiskInfoBefore = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
$FreeSpaceBefore = $DiskInfoBefore.FreeSpace
Write-Host "`nStarting Free Space: $([math]::round($FreeSpaceBefore/1GB, 2)) GB" -ForegroundColor White

# Kill Processes
Write-Host "`n--- STEP 1: Stopping Background Processes ---" -ForegroundColor Yellow
$KillList = @("ms-teams", "idman", "upc", "steam", "EpicGamesLauncher", "OneDrive", "Dropbox", "chrome", "msedge", "firefox", "discord", "cb_launcher", "PowerToys.Run", "upscayl")
foreach ($proc in $KillList) { 
    Stop-TargetProcess $proc
}

# 3. Windows Apps Pruning
Write-Host "`n--- STEP 2: Pruning Apps (User & System) ---" -ForegroundColor Yellow
$pruneList = @(
    "*Microsoft.DesktopAppInstaller*", "*Microsoft.SecHealthUI*", "*NVIDIACorp.NVIDIAControlPanel*",       
    "*RealtekSemiconductorCorp.Realtek*", "*Microsoft.MicrosoftStickyNotes*", "*Microsoft.WindowsTerminal*",           
    "*Microsoft.WindowsNotepad*", "*Microsoft.LanguageExperiencePack*", "*Microsoft.WindowsAppRuntime*",
    "*Microsoft.UI.Xaml*", "*Microsoft.NET.Native.Framework*", "*Microsoft.NET.Native.Runtime*",
    "*Microsoft.VCLibs*", "*Microsoft.DirectXRuntime*", "*Microsoft.HEVCVideoExtension*",        
    "*Microsoft.HEIFImageExtension*", "*Microsoft.VP9VideoExtensions*", "*Microsoft.WebMediaExtensions*",        
    "*Microsoft.WebpImageExtension*", "*Microsoft.MPEG2VideoExtension*", "*Microsoft.AVCEncoderVideoExtension*"   
)

# User Installed Pruning
foreach ($appPattern in $pruneList) {
    $packages = Get-AppxPackage -Name $appPattern -AllUsers -PackageTypeFilter Main, Framework, Resource -ErrorAction SilentlyContinue
    if ($packages) {
        $groupedByName = $packages | Group-Object Name
        foreach ($nameGroup in $groupedByName) {
            $groupedByArch = $nameGroup.Group | Group-Object Architecture
            foreach ($archGroup in $groupedByArch) {
                $sorted = $archGroup.Group | Sort-Object { [version]$_.Version } -Descending
                if ($sorted.Count -gt 1) {
                    $latest = $sorted[0]
                    $olderItems = $sorted | Select-Object -Skip 1
                    foreach ($oldPkg in $olderItems) {
                        if ([version]$oldPkg.Version -lt [version]$latest.Version) {
                            Write-Host "Removing Old: $($oldPkg.Name) v$($oldPkg.Version)" -ForegroundColor DarkGray
                            Remove-AppxPackage -Package $oldPkg.PackageFullName -AllUsers -Confirm:$false -ErrorAction SilentlyContinue
                        }
                    }
                }
            }
        }
    }
}

# Provisioned Pruning
$allProvisioned = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
foreach ($pattern in $pruneList) {
    $matches = $allProvisioned | Where-Object { $_.DisplayName -like $pattern }
    if ($matches) {
        $grouped = $matches | Group-Object DisplayName
        foreach ($group in $grouped) {
             $sorted = $group.Group | Sort-Object { [version]$_.Version } -Descending
             if ($sorted.Count -gt 1) {
                 $latest = $sorted[0]
                 $olderItems = $sorted | Select-Object -Skip 1
                 foreach ($oldPkg in $olderItems) {
                     if ([version]$oldPkg.Version -lt [version]$latest.Version) {
                         Write-Host "Removing Provisioned: $($oldPkg.DisplayName) v$($oldPkg.Version)" -ForegroundColor DarkGray
                         Remove-AppxProvisionedPackage -Online -PackageName $oldPkg.PackageName -ErrorAction SilentlyContinue
                     }
                 }
             }
        }
    }
}

# 4. Bloatware Removal
Write-Host "`n--- STEP 3: Removing Bloatware ---" -ForegroundColor Yellow
$bloatApps = @(
    "*Clipchamp.Clipchamp*", "*Microsoft.MixedReality.Portal*", "*Microsoft.XboxGamingOverlay*",         
    "*Microsoft.MicrosoftOfficeHub*", "*Microsoft.WinDbg*", "*Microsoft.GetHelp*", "*Microsoft.People*",
    "*Microsoft.GetStarted*", "*Microsoft.YourPhone*", "*Microsoft.BingNews*", "*Microsoft.Todos*",                
    "*Microsoft.Wallet*", "*Microsoft.PowerAutomateDesktop*", "*Microsoft.Windows.DevHome*"       
)
foreach ($app in $bloatApps) {
    Get-AppxPackage $app -AllUsers | Remove-AppxPackage -AllUsers -Confirm:$false -ErrorAction SilentlyContinue
    Get-AppxProvisionedPackage -Online | Where-Object {$_.PackageName -like $app} | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
}

# 5. Orphans & Delivery Optimization
Write-Host "`n--- STEP 4: System Optimization ---" -ForegroundColor Yellow
Get-DeliveryOptimizationStatus | Remove-DeliveryOptimizationStatus -Confirm:$false -ErrorAction SilentlyContinue
Start-Process -FilePath "rundll32.exe" -ArgumentList "AppxDeploymentClient.dll,AppxCleanupOrphanPackages" -Wait

# 6. File Cleanup (Junk & Temp)
Write-Host "`n--- STEP 5: Removing Junk Files ---" -ForegroundColor Yellow
Remove-JunkPath "$env:SYSTEMROOT\Temp\*" "System Temp"
Remove-JunkPath "$env:TEMP\*" "User Temp"
Remove-JunkPath "$env:LOCALAPPDATA\Temp\*" "Local Temp"
Remove-JunkPath "$env:SYSTEMROOT\Prefetch\*" "Prefetch"
Remove-JunkPath "$env:SYSTEMROOT\LiveKernelReports\*" "Live Kernel Reports" 
Remove-JunkPath "$env:SYSTEMROOT\Minidump\*" "Minidumps"
Remove-JunkPath "$env:SYSTEMROOT\MEMORY.DMP" "Memory Dump"
Remove-JunkPath "$env:LOCALAPPDATA\CrashDumps\*" "Crash Dumps"
Remove-JunkPath "C:\ProgramData\Microsoft\Windows\WER\ReportArchive\*" "WER Archives"
Remove-JunkPath "C:\ProgramData\Microsoft\Windows\WER\ReportQueue\*" "WER Queue"

# 7. Browser & Edge Aggressive
Remove-JunkPath "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache\*" "Chrome Cache"
Remove-JunkPath "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache\Cache_Data\*" "Edge Cache"
Remove-JunkPath "$env:LOCALAPPDATA\Microsoft\Windows\WebCache\*" "Windows WebCache"
$EdgeBase = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default"
Remove-JunkPath "$EdgeBase\Service Worker\CacheStorage\*" "Edge Service Workers"
Remove-JunkPath "$EdgeBase\Service Worker\ScriptCache\*" "Edge Service ScriptCache"
Remove-JunkPath "$EdgeBase\IndexedDB\*" "Edge IndexedDB"
Remove-JunkPath "$EdgeBase\Code Cache\*" "Edge Code Cache"

# 8. Installer Folders
Write-Host "--- Cleaning Installers ---" -ForegroundColor Cyan
Remove-JunkPath "C:\Autodesk" "Autodesk Install Root"
Remove-JunkPath "C:\adobeTemp" "Adobe Temp Root"
Remove-JunkPath "C:\`$WinREAgent" "WinRE Agent"

# TARGETED: Known User File (Explicit Removal)
$StubbornFile = "C:\Windows\Installer\5ee8fe6.msi"
if (Test-Path $StubbornFile) {
    Write-Host "Deleting persistent target: 5ee8fe6.msi" -ForegroundColor Red
    Set-ItemProperty -Path $StubbornFile -Name Attributes -Value "Normal" -ErrorAction SilentlyContinue
    try {
        Remove-Item -Path $StubbornFile -Force -Confirm:$false -ErrorAction Stop
        Write-Host "Target Deleted." -ForegroundColor Green
    } catch {
        Write-Warning "File locked. Queueing for reboot."
        Register-DeleteOnReboot -FilePath $StubbornFile
    }
}

# --- SMART SCANNER: RHINO/MCNEEL DETECTION (SYSTEM ONLY) ---
Write-Host "`n--- STEP 6: Smart Scanning for Rhino Installers ---" -ForegroundColor Yellow

$InstallerTarget = "C:\Windows\Installer"
if (Test-Path $InstallerTarget) {
    Write-Host "Deep Scanning C:\Windows\Installer..." -ForegroundColor Gray
    $files = Get-ChildItem -Path $InstallerTarget -Recurse -File -Force -ErrorAction SilentlyContinue
    foreach ($file in $files) {
        $prodName = Get-MsiProductName -filePath $file.FullName
        
        if ($prodName -match "Rhino" -and $prodName -notmatch "Rhino.Inside" -or $prodName -match "McNeel") {
            Write-Host "Deleting: $($file.Name) - $prodName" -ForegroundColor Red
            Set-ItemProperty -Path $file.FullName -Name Attributes -Value "Normal" -ErrorAction SilentlyContinue
            try {
                Remove-Item -Path $file.FullName -Force -Confirm:$false -ErrorAction Stop
            } catch {
                Write-Warning "Locked. Queueing for reboot."
                Register-DeleteOnReboot -FilePath $file.FullName
            }
        }
    }
}

# 9. App Data Cleanup
Remove-JunkPath "$env:APPDATA\IDM\DwnlData\*" "IDM Download List"
Remove-JunkPath "C:\Program Files (x86)\Ubisoft\Ubisoft Game Launcher\cache\*" "Ubisoft Cache"
Remove-JunkPath "C:\Program Files (x86)\Microsoft\EdgeUpdate\Download\*" "Edge Update Installers"
Remove-JunkPath "C:\Program Files (x86)\Steam\logs\*" "Steam Logs"
Remove-JunkPath "C:\Program Files (x86)\Common Files\Adobe\Installers\*.log" "Adobe Installer Logs"
Remove-JunkPath "C:\ProgramData\Microsoft\VisualStudio\Packages\*" "VS Installer Packages"
Remove-JunkPath "C:\ProgramData\USOShared\Logs\*" "USO Logs"
Remove-JunkPath "$env:LOCALAPPDATA\upscayl-updater\*" "Upscayl Cache"
Remove-JunkPath "$env:LOCALAPPDATA\Microsoft\PowerToys\Updates\*" "PowerToys Updates"
Remove-JunkPath "$env:LOCALAPPDATA\UniGetUI\CachedMedia\*" "UniGetUI Cache"
Remove-JunkPath "$env:LOCALAPPDATA\Chaos\Cosmos\Updates\*" "Chaos Cosmos Updates"
Remove-JunkPath "$env:LOCALAPPDATA\McNeel\McNeelUpdate\DownloadCache\*" "Rhino User Update Cache"
Remove-JunkPath "C:\ProgramData\McNeel\McNeelUpdate\DownloadCache\*" "Rhino System Update Cache"
Remove-JunkPath "$env:APPDATA\Maxon\*\Redshift\Cache\*" "Maxon Redshift Cache"
Remove-JunkPath "$env:APPDATA\Maxon\*\Redshift\Cache\Textures\*" "Maxon Redshift Textures"
Remove-JunkPath "$env:APPDATA\Maxon\Cinebench*\cache\*" "Cinebench R23/2024 Cache"
Remove-JunkPath "$env:APPDATA\Maxon\*\assets\*" "Maxon Assets Cache"
Remove-JunkPath "$env:LOCALAPPDATA\Adobe\InDesign\*\*\Caches\*" "InDesign Caches"
Remove-JunkPath "$env:LOCALAPPDATA\Gameloft\*\Cache\*" "Gameloft Cache"
Remove-JunkPath "$env:LOCALAPPDATA\Bluebeam\Revu\*\Logs\*" "Bluebeam Logs"
Remove-JunkPath "$env:USERPROFILE\.nuget\packages\*" "NuGet Packages"
Remove-JunkPath "$env:APPDATA\discord\Cache\*" "Discord Cache"
Remove-JunkPath "$env:APPDATA\discord\Code Cache\*" "Discord Code Cache"
Remove-JunkPath "$env:APPDATA\discord\GPUCache\*" "Discord GPUCache"
Remove-JunkPath "$env:LOCALAPPDATA\Autodesk\Revit\PacCache\*" "Revit PacCache"
2018..2030 | ForEach-Object {
    Remove-JunkPath "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit $_\CollaborationCache\*" "Revit $_ Collab"
    Remove-JunkPath "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit $_\Journals\*" "Revit $_ Journals"
}
Remove-JunkPath "$env:HOMEPATH\ACCDOCS\*" "ACCDocs"
Remove-JunkPath "$env:LOCALAPPDATA\Chaos Group\Vantage\cache\*" "Chaos Vantage"
Remove-JunkPath "$env:APPDATA\Adobe\Common\Media Cache Files\*" "Adobe Media Cache"
Remove-JunkPath "$env:LOCALAPPDATA\Adobe\Lightroom\Caches\*" "Lightroom Cache"
Remove-JunkPath "C:\ProgramData\LGHUB\cache\*" "Logitech GHub"
Remove-JunkPath "$env:APPDATA\Zoom\logs\*" "Zoom Logs"
Remove-JunkPath "$env:LOCALAPPDATA\NVIDIA\DXCache\*" "Nvidia DX Cache"
Remove-JunkPath "$env:LOCALAPPDATA\Steam\htmlcache\*" "Steam Web Cache"
Remove-JunkPath "C:\Program Files (x86)\InstallShield Installation Information\*" "InstallShield Leftovers"

# 10. Deep Cleaning (CONDITIONAL)
if ($DeepClean) {
    Write-Host "`n--- STEP 8: Deep System Cleaning (Enabled) ---" -ForegroundColor Red
    if ($PSCmdlet.ShouldProcess("Shadow Copies", "Delete (User Requested)")) {
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c vssadmin delete shadows /all /quiet" -Wait -WindowStyle Hidden
    }
    if ($PSCmdlet.ShouldProcess("CleanMgr", "Run Driver Cleanup")) {
        Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1221" -Wait -WindowStyle Hidden
    }
    if ($PSCmdlet.ShouldProcess("All Event Logs", "Clear via wevtutil")) {
        $logs = wevtutil.exe el
        foreach ($log in $logs) { wevtutil.exe cl "$log" 2>$null }
        Write-Host "Event Logs Cleared." -ForegroundColor Green
    }
} else {
    Write-Host "`n--- STEP 8: Deep Cleaning Skipped (Safe Mode) ---" -ForegroundColor Green
}

# --- SUMMARY ---
$DiskInfoAfter = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
$FreeSpaceAfter = $DiskInfoAfter.FreeSpace
$Recovered = ($FreeSpaceAfter - $FreeSpaceBefore) / 1GB

Write-Host "`n============================================" -ForegroundColor Green
Write-Host " CLEANUP FINISHED" -ForegroundColor Green
Write-Host " Space Before: $([math]::round($FreeSpaceBefore/1GB, 2)) GB"
Write-Host " Space After:  $([math]::round($FreeSpaceAfter/1GB, 2)) GB"
if ($Recovered -gt 0) {
    Write-Host " RECOVERED:    $([math]::round($Recovered, 2)) GB" -ForegroundColor Magenta
}
Write-Host "============================================" -ForegroundColor Green
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")