<#
.SYNOPSIS
    Unjunk v52: selected-file and Windows maintenance cleanup, with a native desktop GUI.
.DESCRIPTION
    No arguments opens the GUI. -DryRun or -Console scans without deleting.
    CLI cleanup requires -Clean and explicit -TargetId values from a scan.
    -Force only bypasses the confirmation for those selected targets.
#>
#requires -Version 5.1
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
param(
    [switch]$Gui,
    [switch]$Console,
    [switch]$DryRun,
    [switch]$Clean,
    [switch]$Force,
    [switch]$StandardCleaning,
    [switch]$CloseApps,
    [Alias('DeleteAfterReboot')][switch]$QueueOnReboot,
    [string[]]$TargetId=@(),
    [ValidateRange(1,365)][int]$TempAgeDays=7,
    [string[]]$ExcludePath=@(),
    [string]$ReportPath
)
$ErrorActionPreference='Stop'
if($env:OS -ne 'Windows_NT') { throw 'Unjunk requires Windows.' }
if($QueueOnReboot -and (-not $Clean -or -not $StandardCleaning)) { throw '-QueueOnReboot requires -Clean -StandardCleaning and explicit -TargetId values.' }
if($Force -and -not $Clean) { throw '-Force now requires -Clean and explicit -TargetId values. Use -Gui or -DryRun to review targets.' }
if($Clean -and -not $TargetId.Count) { throw 'Specify -TargetId values from a preview scan. There is no clean-everything mode.' }
if($Gui -and ($Console -or $DryRun -or $Clean -or $WhatIfPreference)) { throw 'Choose GUI mode or console mode.' }
$useGui=$Gui -or (-not $Console -and -not $DryRun -and -not $Clean -and -not $WhatIfPreference)
if($useGui) {
    if([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        Start-Process "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -WindowStyle Hidden -ArgumentList ('-NoProfile -STA -ExecutionPolicy Bypass -File "'+$PSCommandPath+'" -Gui')
        return
    }
    try { & (Join-Path $PSScriptRoot 'Unjunk.Gui.ps1') }
    catch {
        Add-Type -AssemblyName PresentationFramework
        [Windows.MessageBox]::Show($_.Exception.Message,'Unjunk could not start') | Out-Null
        throw
    }
    return
}
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Core.psm1') -Force
$scan=Invoke-UnjunkScan -TempAgeDays $TempAgeDays -Exclusions $ExcludePath -CrowdStrikeSafe:(-not $StandardCleaning)
Write-Host ('CrowdStrike Safe cleaning: '+$scan.CrowdStrikeSafe+' (conservative profile; no guarantee against security alerts)')
$scan.Targets | Where-Object {$_.FileCount -gt 0 -or $_.Issues.Count -gt 0} |
    Sort-Object Bytes -Descending | Select-Object Id,Name,Size,FileCount,Risk,Status | Format-Table -AutoSize
if($Clean -and -not $DryRun) {
    if($Force -and -not $WhatIfPreference) {
        $report=Invoke-UnjunkCleanup -Scan $scan -TargetIds $TargetId -QueueOnReboot:$QueueOnReboot -CloseApps:$CloseApps -Confirm:$false
    } else {
        $report=Invoke-UnjunkCleanup -Scan $scan -TargetIds $TargetId -QueueOnReboot:$QueueOnReboot -CloseApps:$CloseApps -WhatIf:$WhatIfPreference
    }
    $kind='cleanup'
    Write-Host ('Deleted {0} in {1} files; {2} skipped.' -f (Format-UnjunkSize $report.DeletedBytes),$report.DeletedFiles,$report.SkippedFiles)
    if($report.QueuedFiles) { Write-Host ('Queued for next Windows restart: {0} files ({1}). This space has not been reclaimed yet.' -f $report.QueuedFiles,(Format-UnjunkSize $report.QueuedBytes)) }
} else { $report=$scan; $kind='scan'; Write-Host 'Preview complete. No cleanup was performed.' }
$log=Save-UnjunkReport -Report $report -Kind $kind -Path $ReportPath
Write-Host "Report: $log"
