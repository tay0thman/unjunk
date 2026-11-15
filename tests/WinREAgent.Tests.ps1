#requires -Version 5.1
# All deletions stay in a disposable fixture. Windows status providers are mocked.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Core.psm1') -Force
$module=Get-Module Unjunk.Core
$fixture=Join-Path $PSScriptRoot ('WinREAgentFixture-'+[Guid]::NewGuid().ToString('N'))
$root=Join-Path $fixture '$WinREAgent';$outside=Join-Path $fixture 'outside'
$passed=0;$link=$null;$lock=$null
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
function FixtureFile([string]$Relative){
    $path=Join-Path $root $Relative
    $null=New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($path)) -Force
    [IO.File]::WriteAllText($path,'fixture data')
    (Get-Item -LiteralPath $path).LastWriteTimeUtc=[DateTime]::UtcNow.AddMinutes(-1)
    return $path
}
function Reason { & $module {Get-UnjunkWinREAgentBlockReason} }
function Scan([string[]]$Exclusions=@()) { Invoke-UnjunkSystemFileScan -TempAgeDays 365 -Exclusions $Exclusions }
try {
    $live=@(Get-UnjunkTargets | Where-Object Name -eq 'Windows recovery-update leftovers ($WinREAgent)')
    Assert ($live.Count -eq 1 -and $live[0].Root -eq ($env:SystemDrive+'\$WinREAgent')) 'Catalog uses the literal system-drive $WinREAgent root'
    Assert ($live[0].Risk -eq 'Review' -and -not $live[0].Recommended -and $live[0].Processes.Count -eq 0) 'Recovery leftovers require explicit selection and never map processes for termination'
    $first=FixtureFile 'Backup\winre.wim';$second=FixtureFile 'Scratch\update.wim';$third=FixtureFile 'RollbackInfo.ini'
    $null=New-Item -ItemType Directory -Path $outside
    $outsideFile=Join-Path $outside 'keep.wim';[IO.File]::WriteAllText($outsideFile,'keep outside')
    $outsideHash=(Get-FileHash -LiteralPath $outsideFile).Hash
    $link=Join-Path $root 'Scratch\linked';$null=New-Item -ItemType Junction -Path $link -Target $outside
    & $module {param($fixtureRoot)
        $script:winREFixture=$fixtureRoot;$script:admin=$true;$script:pending='';$script:setupActive=0;$script:busy=@();$script:mounts=@()
        $script:config='<WindowsRE><WinreLocation path="\Recovery\WindowsRE" /></WindowsRE>'
        $script:mountError=$false;$script:mountQueries=0;$script:pendingOnQuery=0;$script:registryError=$false
        function script:Get-UnjunkWinREAgentRoot {$script:winREFixture}
        function script:Get-UnjunkTargets {New-UnjunkTarget 'Windows recovery-update leftovers ($WinREAgent)' 'Windows maintenance' (Get-UnjunkWinREAgentRoot) -Risk Review}
        function script:Test-UnjunkAdministrator {$script:admin}
        function script:Get-Process {param($ErrorAction) foreach($name in $script:busy){[pscustomobject]@{ProcessName=$name}}}
        function script:Test-Path {param($LiteralPath,$ErrorAction)
            if($LiteralPath -like 'HKLM:*'){if($script:registryError){throw 'Fixture registry unavailable'};return ($script:pending -and $LiteralPath.EndsWith($script:pending))}
            Microsoft.PowerShell.Management\Test-Path -LiteralPath $LiteralPath -ErrorAction Stop
        }
        function script:Get-ItemProperty {param($LiteralPath,$ErrorAction)
            if($LiteralPath -ne 'HKLM:\SYSTEM\Setup'){throw 'Unexpected registry access'}
            [pscustomobject]@{SystemSetupInProgress=$script:setupActive;UpgradeInProgress=0}
        }
        function script:Get-Content {param($LiteralPath,[switch]$Raw,$ErrorAction)
            if($LiteralPath -eq (Join-Path $env:SystemRoot 'System32\Recovery\ReAgent.xml')) {return $script:config}
            Microsoft.PowerShell.Management\Get-Content -LiteralPath $LiteralPath -Raw:$Raw -ErrorAction Stop
        }
        function script:Get-WindowsImage {param([switch]$Mounted,$ErrorAction)
            if(-not $Mounted){throw 'Only mounted-image inventory is allowed'}
            $script:mountQueries++
            if($script:mountError){throw 'Fixture image inventory unavailable'}
            if($script:pendingOnQuery -and $script:mountQueries -ge $script:pendingOnQuery){$script:pending='RebootPending'}
            $script:mounts
        }
        function script:Close-UnjunkTargetApps {throw 'Recovery cleanup must never stop processes'}
    } $root
    Assert (-not (Reason)) 'Idle, fully checked recovery-update staging is eligible'
    & $module {$script:admin=$false};Assert ((Reason) -match 'administrator') 'Non-elevated status checks preserve recovery files'; & $module {$script:admin=$true}
    foreach($marker in @('RebootPending','RebootInProgress','PackagesPending','RebootRequired')) {
        & $module {param($name)$script:pending=$name} $marker
        Assert ((Reason) -match 'restart') ('Pending servicing marker blocks cleanup: '+$marker)
    }
    & $module {$script:pending='';$script:setupActive=1};Assert ((Reason) -match 'setup is in progress') 'In-progress Windows setup blocks cleanup'
    & $module {$script:setupActive=0;$script:busy=@('TiWorker')};Assert ((Reason) -match 'TiWorker') 'Active Windows servicing blocks cleanup'
    & $module {$script:busy=@();$script:mounts=@([pscustomobject]@{Path=(Join-Path $script:winREFixture 'Scratch\Mount');ImagePath='C:\other\image.wim'})}
    Assert ((Reason) -match 'mounted Windows image') 'A mounted image inside the cleanup tree is preserved'
    & $module {$script:mounts=@([pscustomobject]@{Path='C:\other\Mount';ImagePath=(Join-Path $script:winREFixture 'Backup\winre.wim')})}
    Assert ((Reason) -match 'mounted Windows image') 'A backup image mounted elsewhere is preserved'
    & $module {$script:mounts=@([pscustomobject]@{Path=($script:winREFixture+'-other\Mount');ImagePath='C:\other\image.wim'})}
    Assert (-not (Reason)) 'Unrelated mounted images and similar folder names do not block this target'
    & $module {$script:mounts=@();$script:config='<WindowsRE><WinreLocation path="\$WinREAgent\Backup" /></WindowsRE>'}
    Assert ((Reason) -match 'recovery still references') 'Registered recovery locations inside $WinREAgent are preserved'
    & $module {$script:config='<WindowsRE />'};Assert ((Reason) -match 'Could not verify') 'Incomplete recovery configuration fails closed'
    & $module {$script:config='<not valid xml'};Assert ((Reason) -match 'Could not verify') 'Unreadable recovery configuration fails closed'
    & $module {$script:config='<WindowsRE><WinreLocation path="\Recovery\WindowsRE" /></WindowsRE>';$script:mountError=$true}
    Assert ((Reason) -match 'Could not verify') 'Failed mounted-image inventory cannot authorize cleanup'
    & $module {$script:mountError=$false;$script:registryError=$true};Assert ((Reason) -match 'Could not verify') 'Failed servicing-state inventory cannot authorize cleanup'
    & $module {$script:registryError=$false;$script:mountQueries=0}
    $safe=Invoke-UnjunkScan -CrowdStrikeSafe $true
    Assert ($safe.Targets[0].PolicyBlocked -and $safe.Targets[0].FileCount -eq 0 -and (& $module {$script:mountQueries}) -eq 0) 'Safe mode blocks recovery cleanup without querying mounted images'
    $scan=Scan
    Assert ($scan.Targets.Count -eq 1 -and $scan.Targets[0].SystemOption -eq 'winre-agent' -and $scan.Targets[0].FileCount -eq 3 -and $scan.Targets[0].MinimumAgeDays -eq 0) 'Windows scan includes all ages as a separate one-click stage'
    Assert ($scan.Targets[0].SkippedLinks -eq 1 -and $outsideFile -notin $scan.Targets[0].Files.Path) 'Scanning never follows a linked recovery subfolder'
    $null=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets[0].Id -CloseApps -WhatIf
    Assert ((Test-Path -LiteralPath $first) -and (Test-Path -LiteralPath $second) -and (Test-Path -LiteralPath $third)) 'WhatIf preserves every scanned file'
    $rejected=$false;try{$null=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets[0].Id -QueueOnReboot -WhatIf}catch{$rejected=$true}
    Assert $rejected 'Recovery-update files can never be queued for reboot deletion'
    & $module {$script:pending='RebootPending'}
    $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets[0].Id -CloseApps -Confirm:$false
    Assert ($result.DeletedFiles -eq 0 -and $result.SkippedFiles -eq 1 -and $result.Results[0].Reason -match 'restart') 'Cleanup rechecks a restart that became pending after scanning'
    $blocked=Scan
    Assert ($blocked.Targets[0].MaintenanceBlock -match 'restart' -and -not $blocked.Targets[0].CanSelect -and $blocked.Targets[0].FileCount -eq 0) 'A blocked scan shows its reason and never enumerates recovery files'
    & $module {$script:pending=''}
    $result=Invoke-UnjunkCleanup -Scan $blocked -TargetIds $blocked.Targets[0].Id -Confirm:$false
    Assert ($result.SkippedFiles -eq 1 -and $result.DeletedFiles -eq 0) 'Blocked inventories remain reported as skipped until a fresh scan'
    & $module {$script:mountQueries=0;$script:pendingOnQuery=4}
    $scan=Scan
    $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets[0].Id -CloseApps -Confirm:$false
    Assert ($result.DeletedFiles -eq 1 -and $result.SkippedFiles -eq 1) 'Servicing starting between files stops the remaining deletions'
    & $module {$script:pending='';$script:pendingOnQuery=0}
    $first=FixtureFile 'Backup\winre.wim';$second=FixtureFile 'Scratch\update.wim';$third=FixtureFile 'RollbackInfo.ini'
    $changed=FixtureFile 'changed.log';$locked=FixtureFile 'locked.log';$excluded=FixtureFile 'excluded.log'
    $scan=Scan @($excluded)
    [IO.File]::AppendAllText($changed,'new contents');$new=FixtureFile 'new-after-scan.log'
    $lock=[IO.File]::Open($locked,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets[0].Id -CloseApps -Confirm:$false
    Assert ($result.DeletedFiles -eq 3 -and $result.DeletedBytes -eq 36 -and $result.SkippedFiles -eq 2) 'Eligible backup, scratch and rollback files are deleted and accurately counted'
    Assert ((Test-Path -LiteralPath $changed) -and (Test-Path -LiteralPath $locked) -and (Test-Path -LiteralPath $excluded) -and (Test-Path -LiteralPath $new)) 'Changed, locked, excluded and newly created files are preserved'
    Assert ((Get-FileHash -LiteralPath $outsideFile).Hash -eq $outsideHash -and $result.ClosedProcesses -eq 0 -and $result.QueuedFiles -eq 0) 'Outside data, running processes and the reboot queue remain untouched'
    Assert ($result.RemovedDirectories -gt 0 -and (Test-Path -LiteralPath $root) -and (Test-Path -LiteralPath $link)) 'Only empty scanned folders are removed; linked and retained contents keep the root intact'
    $lock.Dispose();$lock=$null
    if(-not ([IO.Path]::GetFullPath($link)).StartsWith([IO.Path]::GetFullPath($fixture)+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe fixture link'}
    [IO.Directory]::Delete($link,$false);$link=$null
    $scan=Scan
    $directories=$scan.Targets[0].Directories
    $scan.Targets[0].Directories+=@([pscustomobject]@{Path=$outside;CreationTicks=(Get-Item -LiteralPath $outside).CreationTimeUtc.Ticks})
    $rejected=$false;try{$null=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets[0].Id -Confirm:$false}catch{$rejected=$true}
    Assert ($rejected -and (Test-Path -LiteralPath $changed)) 'Out-of-root directory records are rejected before deleting any files'
    $scan.Targets[0].Directories=$directories
    $newDirectory=Join-Path $root 'NewAfterScan';$null=New-Item -ItemType Directory -Path $newDirectory
    $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets[0].Id -Confirm:$false
    Assert ((Test-Path -LiteralPath $newDirectory) -and (Test-Path -LiteralPath $root)) 'Directories created after scanning are preserved even when empty'
    $scan=Scan
    $null=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets[0].Id -WhatIf
    Assert (Test-Path -LiteralPath $root) 'Preview never removes empty directory shells'
    $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $scan.Targets[0].Id -Confirm:$false
    Assert (-not (Test-Path -LiteralPath $root) -and $result.RemovedDirectories -eq 2) 'A fresh eligible cleanup removes the empty WinREAgent root itself'
    Write-Host "All $passed WinREAgent checks passed."
} finally {
    if($null -ne $lock){$lock.Dispose()}
    $path=[IO.Path]::GetFullPath($fixture);$parent=[IO.Path]::GetFullPath($PSScriptRoot)+'\'
    if(-not $path.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $path -Leaf) -notlike 'WinREAgentFixture-*'){throw 'Unsafe fixture path'}
    if($link -and (Test-Path -LiteralPath $link)) {
        if(-not ([IO.Path]::GetFullPath($link)).StartsWith($path+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe fixture link'}
        [IO.Directory]::Delete($link,$false)
    }
    if(Test-Path -LiteralPath $path){Remove-Item -LiteralPath $path -Recurse -Force}
    Remove-Module Unjunk.Core -Force
}
