<#
.SYNOPSIS
    Master System Maintenance Script
.DESCRIPTION
    1. Cleans temporary files (App, Browser, System).
    2. Removes specified "Bloatware" (Windows Store Apps).
    3. Cleans old Drivers and Windows Update files.
    4. Clears All Windows Event Viewer Logs.
    5. Optimizes Component Store.
.PARAMETER Force
    Skips confirmation prompts.
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param()

# --- CONFIGURATION ---
$ErrorActionPreference = "SilentlyContinue"
$Host.UI.RawUI.WindowTitle = "Master System Maintenance"

# --- ADMIN CHECK ---
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Warning "You must run this script as Administrator!"
    Write-Warning "Right-click PowerShell and select 'Run as Administrator'."
    Break
}

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

function Remove-StoreApp {
    param([string]$AppName)
    Write-Host "Checking for App: $AppName" -ForegroundColor Gray
    Get-AppxPackage $AppName -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
    Get-AppxProvisionedPackage -Online | Where-Object {$_.PackageName -like $AppName} | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
}

# --- START SCRIPT ---

Write-Host "
   __  __          _____ _______ ______ _____  
  |  \/  |   /\   / ____|__   __|  ____|  __ \ 
  | \  / |  /  \ | (___    | |  | |__  | |__) |
  | |\/| | / /\ \ \___ \   | |  |  __| |  _  / 
  | |  | |/ ____ \____) |  | |  | |____| | \ \ 
  |_|  |_/_/    \_\_____/  |_|  |______|_|  \_\
                                               
  Full System Maintenance + Event Log Cleaner
" -ForegroundColor Green

# 1. Measure Disk Space
$DiskInfoBefore = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
$FreeSpaceBefore = $DiskInfoBefore.FreeSpace
Write-Host "Starting Free Space: $([math]::round($FreeSpaceBefore/1GB, 2)) GB" -ForegroundColor White

# 2. Kill Processes
Write-Host "`n--- STEP 1: Stopping Background Processes ---" -ForegroundColor Yellow
$KillList = @("ms-teams", "idman", "upc", "steam", "EpicGamesLauncher", "OneDrive", "Dropbox", "chrome", "msedge", "firefox")
foreach ($proc in $KillList) { Stop-TargetProcess $proc }

# 3. File Cleanup
Write-Host "`n--- STEP 2: Removing Junk Files ---" -ForegroundColor Yellow

# System & Temp
Remove-JunkPath "$env:SYSTEMROOT\Temp\*" "System Temp"
Remove-JunkPath "$env:TEMP\*" "User Temp"
Remove-JunkPath "$env:LOCALAPPDATA\Temp\*" "Local Temp"
Remove-JunkPath "$env:SYSTEMROOT\Prefetch\*" "Prefetch"
Remove-JunkPath "$env:LOCALAPPDATA\CrashDumps\*" "Crash Dumps"
Remove-JunkPath "$env:SYSTEMROOT\memory.dmp" "Memory Dump"

# Browsers
Remove-JunkPath "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache\*" "Chrome Cache"
Remove-JunkPath "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache\Cache_Data\*" "Edge Cache"
Remove-JunkPath "$env:LOCALAPPDATA\Microsoft\Windows\WebCache\*" "Windows WebCache"
Remove-JunkPath "$env:APPDATA\Microsoft\Windows\Recent\*" "Recent Documents"

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

# 4. Old Rhino Installers
Write-Host "`n--- STEP 3: Scanning for old Rhino Installers ---" -ForegroundColor Yellow
if (Test-Path "C:\Windows\Installer") {
    $files = Get-ChildItem "C:\Windows\Installer" -Recurse -File -ErrorAction SilentlyContinue
    $shell = New-Object -ComObject Shell.Application
    foreach ($file in $files) {
        if ($PSCmdlet.ShouldProcess($file.FullName, "Check Subject Property")) {
            try {
                $folder = $shell.Namespace((Split-Path $file.FullName))
                $item = $folder.ParseName((Split-Path $file.FullName -Leaf))
                if ($folder.GetDetailsOf($item, 2) -match "Rhino" -and $folder.GetDetailsOf($item, 2) -notmatch "Rhino\.Inside") {
                    Write-Host "Deleting Old Rhino Installer: $($file.Name)" -ForegroundColor Red
                    Remove-Item $file.FullName -Force
                }
            } catch {}
        }
    }
}

# 5. Bloatware Removal
Write-Host "`n--- STEP 4: Removing Bloatware Apps ---" -ForegroundColor Yellow
if ($PSCmdlet.ShouldProcess("System Apps", "Remove Bloatware")) {
    $BloatList = @(
        "*Clipchamp.Clipchamp*", 
        "*Microsoft.MixedReality.Portal*", 
        "*Microsoft.XboxGamingOverlay*", 
        "*Microsoft.MicrosoftOfficeHub*", 
        "*Microsoft.GetHelp*", 
        "*Microsoft.People*",
        "*Microsoft.GetStarted*",
        "*Microsoft.YourPhone*"
    )
    foreach ($app in $BloatList) { Remove-StoreApp $app }
}

# 6. Deep System Cleaning
Write-Host "`n--- STEP 5: Deep System Cleaning ---" -ForegroundColor Yellow
if ($PSCmdlet.ShouldProcess("WindowsUpdate", "Stop Service & Clean")) {
    Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
    Remove-JunkPath "C:\Windows\SoftwareDistribution\Download\*" "Windows Update Downloads"
    Start-Service -Name wuauserv -ErrorAction SilentlyContinue
}

if ($PSCmdlet.ShouldProcess("DISM", "Component Store Cleanup")) {
    Write-Host "Running DISM Component Cleanup (This takes time)..." -ForegroundColor Cyan
    Start-Process -FilePath "dism.exe" -ArgumentList "/online /Cleanup-Image /StartComponentCleanup" -Wait -NoNewWindow
}

# 7. Driver Cleanup
Write-Host "`n--- STEP 6: Driver Cleanup ---" -ForegroundColor Yellow
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

# 8. Event Viewer Cleaner (Added functionality)
Write-Host "`n--- STEP 8: Clearing Event Viewer Logs ---" -ForegroundColor Yellow
if ($PSCmdlet.ShouldProcess("All Event Logs", "Clear via wevtutil")) {
    Write-Host "Fetching log list..." -ForegroundColor Gray
    $logs = wevtutil.exe el
    $count = 0
    foreach ($log in $logs) {
        $count++
        # Show a progress bar because there are hundreds of logs
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