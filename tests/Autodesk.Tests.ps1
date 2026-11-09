#requires -Version 5.1
# Only disposable installer fixtures; no real C:\Autodesk files are touched.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Core.psm1') -Force
$module=Get-Module Unjunk.Core;$passed=0
$fixture=Join-Path $PSScriptRoot ('AutodeskFixture-'+[Guid]::NewGuid().ToString('N'))
$previousDrive=$env:SystemDrive;$env:SystemDrive=$fixture;$junction=$null
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
try {
    $root=Join-Path $fixture 'Autodesk';$installer=Join-Path $root 'WI\Installer.msi';$license=Join-Path $root 'Network License Manager\license.lic';$excluded=Join-Path $root 'Access\keep.dat'
    foreach($file in @($installer,$license,$excluded)){$null=New-Item -ItemType Directory -Path (Split-Path $file) -Force;Set-Content -LiteralPath $file -Value 'fixture'}
    $outside=Join-Path $fixture 'Outside';$null=New-Item -ItemType Directory -Path $outside
    Set-Content -LiteralPath (Join-Path $outside 'keep.txt') -Value 'outside fixture'
    $junction=Join-Path $root 'linked-folder';$null=New-Item -ItemType Junction -Path $junction -Target $outside
    $target=@(Get-UnjunkTargets | Where-Object Root -eq $root)[0]
    Assert ($target.Risk -eq 'Review' -and 'msiexec' -in $target.Processes -and 'Setup' -in $target.Processes) 'Autodesk extracted installers remain a review target with installer process guards'
    & $module {
        param($t)
        $script:fixtureTarget=$t;$script:installerRunning=$false;$script:accessRunning=$false;$script:closeCalls=0
        function script:Get-UnjunkTargets {@($script:fixtureTarget)}
        function script:Get-UnjunkRunningApps {param($Target) if($script:installerRunning){'Setup'};if($script:accessRunning){'AdskAccessUIHost';'AdskAccessServiceHost'}}
        function script:Close-UnjunkTargetApps {
            $script:closeCalls++;$script:accessRunning=$false
            [pscustomobject]@{Name='Autodesk Access Service Host';ProcessId=0;Status='ServiceStopped';Reason='Fixture'}
            [pscustomobject]@{Name='AdskAccessUIHost';ProcessId=123;Status='Terminated';Reason='Fixture'}
        }
    } $target
    $scan=Invoke-UnjunkSystemFileScan -TempAgeDays 0 -Exclusions @($excluded)
    Assert ($scan.Targets.Count -eq 1 -and $scan.Targets[0].SystemOption -eq 'autodesk-installers' -and $scan.Targets[0].MinimumAgeDays -eq 0) 'The Windows/one-click Autodesk option includes extracted installers of all ages'
    Assert ($scan.Targets[0].Files.Count -eq 1 -and $scan.Targets[0].Files[0].Path -eq $installer) 'Network License Manager and saved exclusions are never selected'
    $main=Invoke-UnjunkScan -CrowdStrikeSafe $false -Exclusions @($excluded)
    Assert ($main.Targets[0].FileCount -eq 1 -and $main.Targets[0].MinimumAgeDays -eq 0) 'The main Standard scan no longer hides Autodesk leftovers behind a 30-day cutoff'
    & $module {$script:installerRunning=$true}
    $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $target.Id -CloseApps -Confirm:$false
    Assert ($result.DeletedFiles -eq 0 -and (Test-Path $installer) -and (& $module {$script:closeCalls}) -eq 0) 'An active Autodesk/Windows installer blocks cleanup and is never terminated'
    & $module {$script:installerRunning=$false;$script:accessRunning=$true}
    $running=Invoke-UnjunkSystemFileScan -TempAgeDays 0 -Exclusions @($excluded)
    Set-UnjunkCategorySelection -Rows $running.Targets -CloseApps
    Assert $running.Targets[0].Selected 'Autodesk Access alone does not prevent Select all when app closing is enabled'
    $preview=Invoke-UnjunkCleanup -Scan $running -TargetIds $target.Id -CloseApps -WhatIf
    Assert ((& $module {$script:closeCalls}) -eq 0 -and (Test-Path $installer)) 'Preview never stops the Access service or deletes installer files'
    $keep=Invoke-UnjunkCleanup -Scan $running -TargetIds $target.Id -Confirm:$false
    Assert ($keep.DeletedFiles -eq 0 -and (& $module {$script:closeCalls}) -eq 0) 'Disabling Close apps leaves Autodesk Access and its files alone'
    $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $target.Id -CloseApps -Confirm:$false
    Assert ($result.DeletedFiles -eq 1 -and -not (Test-Path $installer) -and (Test-Path $license) -and (Test-Path $excluded)) 'Finished installer leftovers are removed while license-manager files and exclusions survive'
    Assert ($result.StoppedServices -eq 1 -and $result.ClosedProcesses -eq 1) 'Autodesk service stop and UI termination are recorded separately'
    Assert (-not (Test-Path (Join-Path $root 'WI')) -and (Test-Path (Join-Path $root 'Network License Manager'))) 'Empty extraction directories are removed while license-manager directories remain'
    Assert ((Test-Path (Join-Path $outside 'keep.txt')) -and (Test-Path $junction)) 'Autodesk cleanup and empty-folder pruning never follow or remove junctions'
    Write-Host "All $passed Autodesk checks passed."
} finally {
    $env:SystemDrive=$previousDrive
    $path=[IO.Path]::GetFullPath($fixture);$parent=[IO.Path]::GetFullPath($PSScriptRoot)+'\'
    if(-not $path.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $path -Leaf) -notlike 'AutodeskFixture-*'){throw 'Unsafe fixture path'}
    if($junction -and (Test-Path -LiteralPath $junction) -and [IO.Path]::GetFullPath($junction).StartsWith($path+'\',[StringComparison]::OrdinalIgnoreCase)){[IO.Directory]::Delete($junction)}
    if(Test-Path -LiteralPath $path){Remove-Item -LiteralPath $path -Recurse -Force}
}
