<#
.SYNOPSIS
    Master System Maintenance Script (v10 - Store Safe)
.DESCRIPTION
    1. Smart-Prunes old App Versions (Keeps Latest).
    * EXCLUDES Microsoft Store from pruning.
    2. Removes Bloatware.
    3. Cleans Junk, Dumps, and Event Logs.
    4. Cleans Rhino Installers.
    5. Interactive Shadow Copy Toggle.
    * PROTECTS Pinned Taskbar Items (JumpLists).
.PARAMETER Force
    Skips confirmation prompts.
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param()

# --- CONFIGURATION ---
$ErrorActionPreference = "SilentlyContinue"
$Host.UI.RawUI.WindowTitle = "Master System Maintenance v10"

# --- ADMIN CHECK ---
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Warning "You must run this script as Administrator!"
    Write-Warning "Right-click PowerShell and select 'Run as Administrator'."
    Break
}

# --- INTERACTIVE TOGGLES ---
Clear-Host
Write-Host "
   __  __          _____ _______ ______ _____  
  |  \/  |   /\   / ____|__   __|  ____|  __ \ 
  | \  / |  /  \ | (___    | |  | |__  | |__) |
  | |\/| | / /\ \ \___ \   | |  |  __| |  _  / 
  | |  | |/ ____ \____) |  | |  | |____| | \ \ 
  |_|  |_/_/    \_\_____/  |_|  |______|_|  \_\
                                               
  Master Maintenance (v10 - Store Safe)
" -ForegroundColor Green

# --- TOGGLE: SHADOW COPIES ---
Write-Host "Shadow Copies (System Restore Points) can take up 10GB+ of space." -ForegroundColor Gray
$ShadowResponse = Read-Host "Do you want to DELETE all Shadow Copies? (y/N)"
if ($ShadowResponse -eq "y") { 
    $CleanShadows = $true 
    Write-Host " -> Shadow Copies will be DELETED." -ForegroundColor Red
} else { 
    $CleanShadows = $false
    Write-Host " -> Shadow Copies will be PRESERVED." -ForegroundColor Green
}
Start-Sleep -Seconds 1

# --- HELPER FUNCTIONS ---

function Stop-TargetProcess {
    param([string]$ProcessName)
    if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) {
        Write-Host "Stopping process: $ProcessName" -ForegroundColor Yellow
        Stop-Process -Name $ProcessName -Force -ErrorAction SilentlyContinue
    }
}

function Remove-JunkPath {
    param([string]$Path, [string]$Desc)
    $ExpandedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if (Test-Path $ExpandedPath) {
        if ($PSCmdlet.ShouldProcess($ExpandedPath, "Clean $Desc")) {
            Write-Host "Cleaning $Desc..." -ForegroundColor Cyan
            Get-ChildItem -Path $ExpandedPath -Recurse -Force -ErrorAction SilentlyContinue | 
                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-FileSubject {
    param ([string]$filePath)
    try {
        $shell = New-Object -ComObject Shell.Application
        $folderPath = Split-Path $filePath
        $fileName = Split-Path $filePath -Leaf
        $folder = $shell.Namespace($folderPath)
        $item = $folder.ParseName($fileName)
        for ($i = 0; $i -lt 300; $i++) {
            $name = $folder.GetDetailsOf($folder.Items, $i)
            if ($name -eq "Subject") { return $folder.GetDetailsOf($item, $i) }
        }
    } catch {}
    return $null
}

# --- START CLEANUP ---

# 1. Measure Disk Space
$DiskInfoBefore = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
$FreeSpaceBefore = $DiskInfoBefore.FreeSpace
Write-Host "`nStarting Free Space: $([math]::round($FreeSpaceBefore/1GB, 2)) GB" -ForegroundColor White

# 2. Kill Processes
Write-Host "`n--- STEP 1: Stopping Background Processes ---" -ForegroundColor Yellow
$KillList = @("ms-teams", "idman", "upc", "steam", "EpicGamesLauncher", "OneDrive", "Dropbox", "chrome", "msedge", "firefox")
foreach ($proc in $KillList) { Stop-TargetProcess $proc }

# 3. Windows Apps Pruning (SMART VERSION)
Write-Host "`n--- STEP 2: Smart-Pruning Old App Versions ---" -ForegroundColor Yellow
Write-Host "Scanning for duplicate versions (Keeping newest only)..." -ForegroundColor Gray

# REMOVED: Microsoft.WindowsStore and Microsoft.StorePurchaseApp
$pruneList = @(
    "*Microsoft.DesktopAppInstaller*",       
    "*Microsoft.SecHealthUI*",               
    "*NVIDIACorp.NVIDIAControlPanel*",       
    "*RealtekSemiconductorCorp.Realtek*",    
    "*Microsoft.MicrosoftStickyNotes*",      
    "*Microsoft.WindowsTerminal*",           
    "*Microsoft.WindowsNotepad*",            
    "*Microsoft.LanguageExperiencePack*",    
    "*Microsoft.WindowsAppRuntime*"          
)

foreach ($appPattern in $pruneList) {
    # Get all packages matching the pattern
    $packages = Get-AppxPackage -Name $appPattern -AllUsers -ErrorAction SilentlyContinue

    if ($packages) {
        # Group by Name (in case wildcard catches multiple distinct apps)
        $grouped = $packages | Group-Object Name

        foreach ($group in $grouped) {
            # Sort by Version Descending (Newest is top)
            $sorted = $group.Group | Sort-Object { [version]$_.Version } -Descending

            if ($sorted.Count -gt 1) {
                $latest = $sorted[0]
                $old = $sorted | Select-Object -Skip 1

                Write-Host "Checking $($latest.Name):" -ForegroundColor Cyan
                Write-Host "  [KEEP] Latest: $($latest.Version)" -ForegroundColor Green
                
                foreach ($oldPkg in $old) {
                    if ($PSCmdlet.ShouldProcess("$($oldPkg.Name) v$($oldPkg.Version)", "Remove Old Version")) {
                        Write-Host "  [DEL]  Old:    $($oldPkg.Version)" -ForegroundColor Red
                        Remove-AppxPackage -Package $oldPkg.PackageFullName -AllUsers -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    }
}

# 4. Bloatware Removal
Write-Host "`n--- STEP 3: Removing Bloatware ---" -ForegroundColor Yellow
$bloatApps = @(
    "*Clipchamp.Clipchamp*",                 
    "*Microsoft.MixedReality.Portal*",       
    "*Microsoft.XboxGamingOverlay*",         
    "*Microsoft.MicrosoftOfficeHub*",        
    "*Microsoft.WinDbg*",                    
    "*Microsoft.GetHelp*", 
    "*Microsoft.People*",
    "*Microsoft.GetStarted*",
    "*Microsoft.YourPhone*"
)

foreach ($app in $bloatApps) {
    if ($PSCmdlet.ShouldProcess($app, "Remove Bloatware")) {
        Write-Host "Removing $app..." -ForegroundColor Cyan
        Get-AppxPackage $app -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        Get-AppxProvisionedPackage -Online | Where-Object {$_.PackageName -like $app} | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
    }
}

# 5. App Orphans & Delivery Optimization
Write-Host "`n--- STEP 4: Cleaning App Orphans ---" -ForegroundColor Yellow
if ($PSCmdlet.ShouldProcess("DeliveryOptimization", "Clean")) {
    Get-DeliveryOptimizationStatus | Remove-DeliveryOptimizationStatus -ErrorAction SilentlyContinue
}
if ($PSCmdlet.ShouldProcess("AppxDeploymentClient", "Cleanup Orphan Packages")) {
    Start-Process -FilePath "rundll32.exe" -ArgumentList "AppxDeploymentClient.dll,AppxCleanupOrphanPackages" -Wait
}

# 6. File Cleanup
Write-Host "`n--- STEP 5: Removing Junk Files ---" -ForegroundColor Yellow

# System Temp & Prefetch
Remove-JunkPath "$env:SYSTEMROOT\Temp\*" "System Temp"
Remove-JunkPath "$env:TEMP\*" "User Temp"
Remove-JunkPath "$env:LOCALAPPDATA\Temp\*" "Local Temp"
Remove-JunkPath "$env:SYSTEMROOT\Prefetch\*" "Prefetch"

# Memory Dumps & Watchdog
Remove-JunkPath "$env:SYSTEMROOT\LiveKernelReports\*" "Live Kernel Reports (Watchdog)" 
Remove-JunkPath "$env:SYSTEMROOT\Minidump\*" "Blue Screen Minidumps"
Remove-JunkPath "$env:SYSTEMROOT\MEMORY.DMP" "Full Memory Dump"
Remove-JunkPath "$env:LOCALAPPDATA\CrashDumps\*" "Application Crash Dumps"
Remove-JunkPath "C:\ProgramData\Microsoft\Windows\WER\ReportArchive\*" "WER Archives"
Remove-JunkPath "C:\ProgramData\Microsoft\Windows\WER\ReportQueue\*" "WER Queue"

# Browsers & Recent - FIXED LOGIC TO PROTECT JUMP LISTS
Remove-JunkPath "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache\*" "Chrome Cache"
Remove-JunkPath "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache\Cache_Data\*" "Edge Cache"
Remove-JunkPath "$env:LOCALAPPDATA\Microsoft\Windows\WebCache\*" "Windows WebCache"

# Specific Recent Files Clean (Protects AutomaticDestinations)
$RecentPath = "$env:APPDATA\Microsoft\Windows\Recent"
if (Test-Path $RecentPath) {
    if ($PSCmdlet.ShouldProcess($RecentPath, "Clean Recent Shortcuts (Protecting JumpLists)")) {
        Write-Host "Cleaning Recent Shortcuts (Protecting Pinned Items)..." -ForegroundColor Cyan
        Get-ChildItem -Path $RecentPath -File -Force -ErrorAction SilentlyContinue | 
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
}

# Autodesk / CAD / 3D
Remove-JunkPath "C:\Autodesk\*" "Autodesk Installers"
Remove-JunkPath "$env:LOCALAPPDATA\Autodesk\Revit\PacCache\*" "Revit PacCache"
2018..2026 | ForEach-Object {
    Remove-JunkPath "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit $_\CollaborationCache\*" "Revit $_ Collab"
    Remove-JunkPath "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit $_\Journals\*" "Revit $_ Journals"
}
Remove-JunkPath "$env:HOMEPATH\ACCDOCS\*" "ACCDocs"
Remove-JunkPath "$env:LOCALAPPDATA\Chaos Group\Vantage\cache\*" "Chaos Vantage"
Remove-JunkPath "C:\ProgramData\McNeel\McNeelUpdate\DownloadCache\*" "Rhino Update Cache"

# Adobe
Remove-JunkPath "C:\adobeTemp" "Adobe Temp Root"
Remove-JunkPath "$env:APPDATA\Adobe\Common\Media Cache Files\*" "Adobe Media Cache"
Remove-JunkPath "$env:LOCALAPPDATA\Adobe\Lightroom\Caches\*" "Lightroom Cache"

# Gaming & Tools
Remove-JunkPath "C:\ProgramData\LGHUB\cache\*" "Logitech GHub"
Remove-JunkPath "$env:APPDATA\Zoom\logs\*" "Zoom Logs"
Remove-JunkPath "$env:LOCALAPPDATA\NVIDIA\DXCache\*" "Nvidia DX Cache"
Remove-JunkPath "$env:LOCALAPPDATA\Steam\htmlcache\*" "Steam Web Cache"
Remove-JunkPath "$env:LOCALAPPDATA\EpicGamesLauncher\Saved\webcache_*" "Epic Games WebCache"
Remove-JunkPath "C:\Program Files (x86)\InstallShield Installation Information\*" "InstallShield Leftovers"

# 7. Old Rhino Installers (Subject Check)
Write-Host "`n--- STEP 6: Scanning for old Rhino Installers ---" -ForegroundColor Yellow
$InstallerTarget = "C:\Windows\Installer"
if (Test-Path $InstallerTarget) {
    $files = Get-ChildItem -Path $InstallerTarget -Recurse -File -ErrorAction SilentlyContinue
    foreach ($file in $files) {
        $subject = Get-FileSubject -filePath $file.FullName
        if ($subject -and $subject -match "Rhino" -and $subject -notmatch "Rhino\.Inside") {
            if ($PSCmdlet.ShouldProcess($file.FullName, "Delete Rhino Installer ($subject)")) {
                try {
                    Write-Host "Deleting: $($file.FullName)" -ForegroundColor Red
                    Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
                } catch { Write-Warning "Failed to delete: $($file.FullName)" }
            }
        }
    }
}

# 8. Deep System Cleaning (Shadows + Update + Drivers)
Write-Host "`n--- STEP 7: Deep System Cleaning ---" -ForegroundColor Yellow

if ($CleanShadows) {
    if ($PSCmdlet.ShouldProcess("Shadow Copies", "Delete (User Requested)")) {
        Write-Host "Deleting VSS Shadow Copies (Silent)..." -ForegroundColor Red
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c vssadmin delete shadows /all /quiet" -Wait -WindowStyle Hidden
    }
}

if ($PSCmdlet.ShouldProcess("WindowsUpdate", "Stop Service & Clean")) {
    Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
    Remove-JunkPath "C:\Windows\SoftwareDistribution\Download\*" "Windows Update Downloads"
    Start-Service -Name wuauserv -ErrorAction SilentlyContinue
}

if ($PSCmdlet.ShouldProcess("DISM", "Component Store Cleanup")) {
    Write-Host "Running DISM Component Cleanup (Background)..." -ForegroundColor Cyan
    Start-Process -FilePath "dism.exe" -ArgumentList "/online /Cleanup-Image /StartComponentCleanup /Quiet /NoRestart" -Wait -WindowStyle Hidden
}

# Driver Cleanup (Minimized)
$CleanMgrKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
$StateFlags = @("Device Driver Packages", "Temporary Files", "Update Cleanup", "Windows Defender")
if ($PSCmdlet.ShouldProcess("CleanMgr", "Run Driver Cleanup")) {
    foreach ($flag in $StateFlags) {
        if (Test-Path "$CleanMgrKey\$flag") {
            Set-ItemProperty -Path "$CleanMgrKey\$flag" -Name StateFlags1221 -Type DWORD -Value 2
        }
    }
    Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1221" -Wait -WindowStyle Hidden
}

# 9. Event Viewer Cleaner
Write-Host "`n--- STEP 8: Clearing Event Viewer Logs ---" -ForegroundColor Yellow
if ($PSCmdlet.ShouldProcess("All Event Logs", "Clear via wevtutil")) {
    Write-Host "Fetching log list..." -ForegroundColor Gray
    $logs = wevtutil.exe el
    $count = 0
    foreach ($log in $logs) {
        $count++
        Write-Progress -Activity "Clearing Event Logs" -Status "$log" -PercentComplete (($count / $logs.Count) * 100)
        wevtutil.exe cl "$log" 2>$null
    }
    Write-Progress -Activity "Clearing Event Logs" -Completed
    Write-Host "All Event Logs have been cleared." -ForegroundColor Green
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