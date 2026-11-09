#requires -Version 5.1
# Synthetic process handles only; no real application is closed or terminated.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Core.psm1') -Force
$module=Get-Module Unjunk.Core
$fixture=Join-Path $PSScriptRoot ('CloseAppsFixture-'+[Guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $fixture
$passed=0
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
function ProcessFixture([string]$Name='Claude',[bool]$Graceful=$true,[int]$Session=1,[int]$ProcessId=45678,[bool]$FailKill=$false){
    $p=[pscustomobject]@{ProcessName=$Name;Id=$ProcessId;SessionId=$Session;Handle=123;MainWindowHandle=[IntPtr]1;HasExited=$false;Graceful=$Graceful;FailKill=$FailKill;CloseRequested=$false;Killed=$false;Disposed=$false;CachePresentAtClose=$false;CachePath=(Join-Path $fixture 'cache.bin')}
    $p | Add-Member ScriptMethod CloseMainWindow {$this.CloseRequested=$true;$this.CachePresentAtClose=Test-Path -LiteralPath $this.CachePath;if($this.Graceful){$this.HasExited=$true};return $true}
    $p | Add-Member ScriptMethod Kill {if($this.FailKill){throw 'Access denied fixture'};$this.Killed=$true;$this.HasExited=$true}
    $p | Add-Member ScriptMethod WaitForExit {param($Milliseconds) return $this.HasExited}
    $p | Add-Member ScriptMethod Dispose {$this.Disposed=$true}
    return $p
}
function ScanFixture($Process,[string]$Risk='Cache'){
    Set-Content -LiteralPath (Join-Path $fixture 'cache.bin') -Value 'disposable cache'
    & $module {param($p,$risk) $script:fixtureProcess=$p;$script:fixtureTarget.Risk=$risk} $Process $Risk
    Invoke-UnjunkScan -CrowdStrikeSafe $false
}
try {
    & $module {
        param($root)
        $script:fixtureTarget=New-UnjunkTarget 'Fixture Claude cache' 'Desktop apps' $root @('Claude')
        $script:realClose=(Get-Item Function:Close-UnjunkTargetApps).ScriptBlock
        function script:Get-UnjunkTargets {@($script:fixtureTarget)}
        function script:Get-UnjunkInteractiveSession {1}
        function script:Get-UnjunkAppProcesses {param($Target) if(-not $script:fixtureProcess.HasExited){$script:fixtureProcess}}
        function script:Get-UnjunkRunningApps {param($Target) if(-not $script:fixtureProcess.HasExited){$script:fixtureProcess.ProcessName}}
        function script:Close-UnjunkTargetApps {param($Target,$State) & $script:realClose -Target $Target -State $State -GraceSeconds 0}
    } $fixture
    $p=ProcessFixture;$scan=ScanFixture $p
    $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets.Id -Confirm:$false
    Assert ($result.DeletedFiles -eq 0 -and -not $p.CloseRequested -and -not $p.Killed) 'App-closing is opt-in at the cleanup API'
    $preview=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets.Id -CloseApps -WhatIf
    Assert (-not $p.CloseRequested -and -not $p.Killed -and (Test-Path $p.CachePath)) 'WhatIf never closes an app or deletes its cache'
    $cancel=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets.Id -CloseApps -State @{Cancel=$true} -Confirm:$false
    Assert ($cancel.Cancelled -and -not $p.CloseRequested) 'Cancellation prevents both closing and deletion'
    Set-UnjunkCategorySelection -Rows $scan.Targets -CloseApps
    Assert $scan.Targets[0].Selected 'Select all includes running caches when app-closing is enabled'
    $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets.Id -CloseApps -Confirm:$false
    Assert ($p.CloseRequested -and $p.CachePresentAtClose -and -not $p.Killed -and $result.DeletedFiles -eq 1) 'A graceful close completes before cache deletion'
    Assert ($result.ClosedProcesses -eq 1 -and $result.AppClosures[0].Status -eq 'Closed' -and $p.Disposed) 'Graceful exits are counted and process handles are released'
    $p=ProcessFixture -Graceful $false;$scan=ScanFixture $p
    $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets.Id -CloseApps -Confirm:$false
    Assert ($p.CloseRequested -and $p.Killed -and $result.DeletedFiles -eq 1 -and $result.AppClosures[0].Status -eq 'Terminated') 'An app that refuses normal close is stopped before cleanup'
    $p=ProcessFixture -Graceful $false -FailKill $true;$scan=ScanFixture $p
    $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets.Id -CloseApps -Confirm:$false
    Assert ($result.DeletedFiles -eq 0 -and (Test-Path $p.CachePath) -and $result.AppClosures[0].Reason -match 'denied') 'Failure to stop an app preserves its cache and reports the reason'
    foreach($case in @(@{Name='CSFalconService'},@{Name='Dism'},@{Name='cleanmgr'},@{Name='msiexec'},@{Name='Unjunk'},@{Name='Claude';Session=0},@{Name='Claude';Session=9},@{Name='Claude';ProcessId=$PID})){
        $p=ProcessFixture @case;$scan=ScanFixture $p
        $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets.Id -CloseApps -Confirm:$false
        Assert (-not $p.CloseRequested -and -not $p.Killed -and $result.DeletedFiles -eq 0) ('Protected/session process remains running: '+$p.ProcessName+'/'+$p.SessionId)
    }
    $p=ProcessFixture;$scan=ScanFixture $p -Risk Review
    Set-UnjunkCategorySelection -Rows $scan.Targets -CloseApps
    $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets.Id -CloseApps -Confirm:$false
    Assert (-not $scan.Targets[0].Selected -and -not $p.CloseRequested -and $result.DeletedFiles -eq 0) 'Review/installer targets never trigger app termination'
    $p=ProcessFixture;$scan=ScanFixture $p;Remove-Item -LiteralPath $p.CachePath
    $scan=Invoke-UnjunkScan -CrowdStrikeSafe $false
    $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets.Id -CloseApps -Confirm:$false
    Assert (-not $p.CloseRequested -and -not $p.Killed) 'An empty target does not close any app'
    Write-Host "All $passed app-closing checks passed."
} finally {
    Remove-Module Unjunk.Core -Force
    $path=[IO.Path]::GetFullPath($fixture);$parent=[IO.Path]::GetFullPath($PSScriptRoot)+'\'
    if(-not $path.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $path -Leaf) -notlike 'CloseAppsFixture-*'){throw 'Unsafe fixture path'}
    Remove-Item -LiteralPath $path -Recurse -Force
}
