#requires -Version 5.1
# Self-contained regression tests; no Pester installation required.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Core.psm1') -Force
$module=Get-Module Unjunk.Core
$testRoot=Join-Path $env:TEMP ('Unjunk-tests-'+[Guid]::NewGuid().ToString('N'))
$null=New-Item -Path $testRoot -ItemType Directory
$script:passed=0
function Assert([bool]$Condition,[string]$Message) {
    if(-not $Condition) { throw "FAIL: $Message" }
    $script:passed++; Write-Host "PASS: $Message"
}
function Candidate([string]$Path) {
    $item=Get-Item -LiteralPath $Path -Force
    [pscustomobject]@{Path=$item.FullName;Length=[long]$item.Length;LastWriteTicks=$item.LastWriteTimeUtc.Ticks;CreationTicks=$item.CreationTimeUtc.Ticks}
}
function DeleteCandidate($File,[string]$Root,[string[]]$Excluded=@(),[switch]$Preview) {
    & $module {param($f,$r,$e,$p) Remove-UnjunkCandidate -File $f -Root $r -Exclusions $e -WhatIf:$p -Confirm:$false} $File $Root $Excluded ([bool]$Preview)
}
try {
    $root=Join-Path $testRoot 'cache'; $outside=Join-Path $testRoot 'outside'
    $null=New-Item -Path $root,$outside -ItemType Directory
    $old=Join-Path $root 'old.tmp'; $recent=Join-Path $root 'recent.tmp'; $excluded=Join-Path $root 'keep.tmp'
    Set-Content -LiteralPath $old -Value 'old'; Set-Content -LiteralPath $recent -Value 'recent'; Set-Content -LiteralPath $excluded -Value 'keep'
    (Get-Item $old).LastWriteTimeUtc=[DateTime]::UtcNow.AddDays(-20)
    (Get-Item $excluded).LastWriteTimeUtc=[DateTime]::UtcNow.AddDays(-20)
    $target=& $module {param($r) New-UnjunkTarget 'Test temp' 'Tests' $r -Risk Temporary -MinimumAgeDays 7} $root
    $scan=& $module {param($t,$e) Get-UnjunkTargetScan $t 7 @($e) @{Cancel=$false}} $target $excluded
    Assert ($scan.FileCount -eq 1 -and $scan.Files[0].Path -eq $old) 'Age filter and exclusions limit the scan'
    Assert ((Test-Path $old) -and (Test-Path $recent)) 'Scanning does not delete files'
    $deepScan=& $module {param($t,$e) Get-UnjunkTargetScan $t 0 @($e) @{Cancel=$false} $false} $target $excluded
    Assert ($deepScan.MinimumAgeDays -eq 0 -and @($deepScan.Files | Where-Object Path -eq $recent).Count -eq 1) 'One-click zero-age temporary scan includes recent eligible files'
    $preview=DeleteCandidate (Candidate $old) $root -Preview
    Assert ($preview.Status -eq 'Preview' -and (Test-Path $old)) 'WhatIf preserves scanned files'
    $blocked=DeleteCandidate (Candidate $excluded) $root @($excluded)
    Assert ($blocked.Status -eq 'Skipped' -and (Test-Path $excluded)) 'Cleanup rechecks exclusions'
    $other=Join-Path $outside 'outside.txt'; Set-Content -LiteralPath $other -Value 'outside'
    $blocked=DeleteCandidate (Candidate $other) $root
    Assert ($blocked.Status -eq 'Skipped' -and (Test-Path $other)) 'Sibling paths cannot escape the target root'
    $changed=Candidate $old; Add-Content -LiteralPath $old -Value 'changed'
    $blocked=DeleteCandidate $changed $root
    Assert ($blocked.Status -eq 'Skipped' -and (Test-Path $old)) 'Changed files are preserved'
    $locked=Join-Path $root 'locked.tmp'; Set-Content -LiteralPath $locked -Value 'locked'
    $lockedCandidate=Candidate $locked
    $stream=[IO.File]::Open($locked,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    try { $blocked=DeleteCandidate $lockedCandidate $root; Assert ($blocked.Status -eq 'Skipped' -and (Test-Path $locked)) 'Locked files are skipped without reboot scheduling' }
    finally { $stream.Dispose() }
    $link=Join-Path $root 'junction'
    $null=New-Item -ItemType Junction -Path $link -Target $outside
    $linkTarget=& $module {param($r) New-UnjunkTarget 'Junction test' 'Tests' $r} $root
    $scan=& $module {param($t) Get-UnjunkTargetScan $t 7 @() @{Cancel=$false}} $linkTarget
    Assert ($scan.SkippedLinks -eq 1 -and @($scan.Files | Where-Object Path -like '*junction*').Count -eq 0) 'Scan does not follow junctions'
    $linked=Candidate (Join-Path $link 'outside.txt')
    $blocked=DeleteCandidate $linked $root
    Assert ($blocked.Status -eq 'Skipped' -and (Test-Path $other)) 'Cleanup rejects a path through a junction'
    # Remove just the test junction itself, never its target or descendants.
    [IO.Directory]::Delete($link)
    $oldCandidate=Candidate $old
    $newFile=Join-Path $root 'created-after-scan.tmp'; Set-Content -LiteralPath $newFile -Value 'keep'
    $deleted=DeleteCandidate $oldCandidate $root
    Assert ($deleted.Status -eq 'Deleted' -and -not (Test-Path $old)) 'Unchanged selected file is deleted and counted'
    Assert ((Test-Path $newFile) -and (Test-Path $recent)) 'New and unselected files remain in the directory'
    foreach($protected in @("$env:SystemRoot\Installer\sample.msi","$env:SystemRoot\WinSxS\sample","$env:ProgramData\Package Cache\sample",$env:USERPROFILE)) {
        $caught=$false; try { & $module {param($p) Assert-UnjunkRoot $p} $protected | Out-Null } catch {$caught=$true}
        Assert $caught "Protected root rejected: $protected"
    }
    $cancelled=& $module {param($t) Get-UnjunkTargetScan $t 7 @() @{Cancel=$true}} $target
    Assert ($cancelled.FileCount -eq 0) 'Cancellation stops enumeration'
    $catalog=@(Get-UnjunkTargets)
    Assert ($catalog.Count -eq @($catalog.Id | Sort-Object -Unique).Count) 'Catalog roots are deduplicated'
    Assert (@($catalog | Where-Object {$_.Risk -in @('Review','Diagnostics') -and $_.Recommended}).Count -eq 0) 'Diagnostics and installer review items are never recommended automatically'
    $running=& $module {Get-UnjunkRunningApps ([pscustomobject]@{Processes=@((Get-Process -Id $PID).ProcessName)})}
    Assert (@($running).Count -gt 0) 'Active app detection works'
    $fake=[pscustomobject]@{Cancelled=$true}
    $caught=$false; try { Invoke-UnjunkCleanup -Scan $fake -TargetIds @('unknown') -Confirm:$false } catch {$caught=$true}
    Assert $caught 'A cancelled scan cannot clean'
    $valid=[pscustomobject]@{Cancelled=$false;UserSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value;Computer=$env:COMPUTERNAME;ScannedAt=[DateTime]::UtcNow.AddHours(-1).ToString('o');Targets=@();Exclusions=@()}
    $caught=$false; try { Invoke-UnjunkCleanup -Scan $valid -TargetIds @('unknown') -Confirm:$false } catch {$caught=$true}
    Assert $caught 'An expired scan cannot clean'
    $valid.ScannedAt=[DateTime]::UtcNow.ToString('o')
    $caught=$false; try { Invoke-UnjunkCleanup -Scan $valid -TargetIds @('unknown') -Confirm:$false } catch {$caught=$true}
    Assert $caught 'Unknown selection is rejected'
    $json=Join-Path $testRoot 'report.json'; $null=Save-UnjunkReport -Report $scan -Path $json
    $roundTrip=Get-Content -LiteralPath $json -Raw | ConvertFrom-Json
    Assert ($roundTrip.Root -eq $root) 'Reports preserve paths and scan metadata'
    $safeTarget=& $module {param($r) New-UnjunkTarget 'Safe cache fixture' 'Tests' $r} $root
    $oldCache=Join-Path $root 'stale-cache.tmp'; Set-Content -LiteralPath $oldCache -Value 'stale cache'
    (Get-Item $oldCache).LastWriteTimeUtc=[DateTime]::UtcNow.AddDays(-10)
    $safeScan=& $module {param($t) Get-UnjunkTargetScan $t 1 @() @{Cancel=$false} $true} $safeTarget
    Assert ($safeScan.MinimumAgeDays -eq 7 -and @($safeScan.Files | Where-Object Path -eq $recent).Count -eq 0) 'CrowdStrike Safe enforces seven days even with a lower configured age'
    Assert (@($safeScan.Files | Where-Object Path -eq $oldCache).Count -eq 1) 'CrowdStrike Safe allows old user-cache files'
    foreach($risk in @('Review','Diagnostics')) {
        $blockedTarget=& $module {param($r,$risk) New-UnjunkTarget 'Blocked fixture' 'Tests' $r -Risk $risk} $root $risk
        $blockedScan=& $module {param($t) Get-UnjunkTargetScan $t 1 @() @{Cancel=$false} $true} $blockedTarget
        Assert ($blockedScan.PolicyBlocked -and $blockedScan.FileCount -eq 0) "CrowdStrike Safe blocks $risk targets without enumerating them"
    }
    $systemTarget=@(Get-UnjunkTargets | Where-Object Root -eq "$env:SystemRoot\Temp")[0]
    $blockedScan=& $module {param($t) Get-UnjunkTargetScan $t 1 @() @{Cancel=$false} $true} $systemTarget
    Assert $blockedScan.PolicyBlocked 'CrowdStrike Safe excludes system-level temp cleanup'
    # Exercise the public cleanup workflow against a fixture-only catalog.
    $originalCatalog=& $module { (Get-Item Function:Get-UnjunkTargets).ScriptBlock }
    try {
        & $module {param($target) $script:fixtureTarget=$target; function script:Get-UnjunkTargets { @($script:fixtureTarget) }} $safeTarget
        $fixtureScan=[pscustomobject]@{Cancelled=$false;UserSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value;Computer=$env:COMPUTERNAME;ScannedAt=[DateTime]::UtcNow.ToString('o');Targets=@($safeScan);Exclusions=@();CrowdStrikeSafe=$true}
        $previewResult=Invoke-UnjunkCleanup -Scan $fixtureScan -TargetIds @($safeTarget.Id) -WhatIf
        Assert ($previewResult.DeletedFiles -eq 0 -and (Test-Path $oldCache)) 'Public WhatIf does not delete files'
        $cancelResult=Invoke-UnjunkCleanup -Scan $fixtureScan -TargetIds @($safeTarget.Id) -State @{Cancel=$true;Progress=0;Message=''} -Confirm:$false
        Assert ($cancelResult.Cancelled -and (Test-Path $oldCache)) 'Public cleanup honors cancellation before deletion'
        $safeTarget.Processes=@((Get-Process -Id $PID).ProcessName)
        $runningResult=Invoke-UnjunkCleanup -Scan $fixtureScan -TargetIds @($safeTarget.Id) -Confirm:$false
        Assert ($runningResult.SkippedFiles -eq 1 -and (Test-Path $oldCache)) 'Public cleanup skips a running application target'
        $safeTarget.Processes=@()
        $cleanupResult=Invoke-UnjunkCleanup -Scan $fixtureScan -TargetIds @($safeTarget.Id) -Confirm:$false
        Assert ($cleanupResult.DeletedFiles -gt 0 -and -not (Test-Path $oldCache) -and (Test-Path $recent)) 'Public cleanup removes only eligible selected fixture files'
        Assert $cleanupResult.CrowdStrikeSafe 'Cleanup report records the CrowdStrike Safe profile'
    } finally {
        & $module {param($original) Set-Item Function:script:Get-UnjunkTargets -Value $original; Remove-Variable fixtureTarget -Scope Script} $originalCatalog
    }
    # Mock only the OS/elevation boundary. Never write to PendingFileRenameOperations.
    $originalNative=& $module { (Get-Item Function:Invoke-UnjunkNativeRebootDelete).ScriptBlock }
    $originalAdmin=& $module { (Get-Item Function:Test-UnjunkAdministrator).ScriptBlock }
    $originalCatalog=& $module { (Get-Item Function:Get-UnjunkTargets).ScriptBlock }
    try {
        & $module {
            param($t)
            Initialize-UnjunkRebootApi
            $script:fixtureTarget=$t; $script:queuedTestPaths=New-Object 'System.Collections.Generic.List[string]'
            function script:Get-UnjunkTargets { @($script:fixtureTarget) }
            function script:Test-UnjunkAdministrator { $true }
            function script:Invoke-UnjunkNativeRebootDelete([string]$Path) { $script:queuedTestPaths.Add($Path) }
        } $safeTarget
        $queueFile=Join-Path $root 'queue.tmp'; Set-Content -LiteralPath $queueFile -Value 'queue fixture'
        $queueCandidate=Candidate $queueFile
        $queueRow=& $module {param($t) Get-UnjunkTargetScan $t 1 @() @{Cancel=$false} $false} $safeTarget
        $queueScan=[pscustomobject]@{Cancelled=$false;UserSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value;Computer=$env:COMPUTERNAME;ScannedAt=[DateTime]::UtcNow.ToString('o');Targets=@($queueRow);Exclusions=@();CrowdStrikeSafe=$true}
        $caught=$false; try { Invoke-UnjunkCleanup -Scan $queueScan -TargetIds @($safeTarget.Id) -QueueOnReboot -Confirm:$false } catch {$caught=$true}
        Assert $caught 'CrowdStrike Safe blocks reboot queuing in the engine'
        $queueScan.CrowdStrikeSafe=$false
        & $module { function script:Test-UnjunkAdministrator { $false } }
        $caught=$false; try { Invoke-UnjunkCleanup -Scan $queueScan -TargetIds @($safeTarget.Id) -QueueOnReboot -Confirm:$false } catch {$caught=$true}
        Assert $caught 'Reboot queuing requires administrator access'
        $preview=Invoke-UnjunkCleanup -Scan $queueScan -TargetIds @($safeTarget.Id) -QueueOnReboot -WhatIf
        $queuedCount=& $module { $script:queuedTestPaths.Count }
        Assert ($queuedCount -eq 0 -and $preview.QueuedFiles -eq 0) 'Reboot WhatIf never calls the Windows scheduling boundary'
        & $module { function script:Test-UnjunkAdministrator { $true } }
        $safeTarget.Risk='Diagnostics'
        $caught=$false; try { Invoke-UnjunkCleanup -Scan $queueScan -TargetIds @($safeTarget.Id) -QueueOnReboot -Confirm:$false } catch {$caught=$true}
        Assert $caught 'Diagnostics cannot be queued for deletion at reboot'
        $safeTarget.Risk='Cache'; $safeTarget.Processes=@((Get-Process -Id $PID).ProcessName)
        $lockStream=[IO.File]::Open($queueFile,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        try { $queued=Invoke-UnjunkCleanup -Scan $queueScan -TargetIds @($safeTarget.Id) -QueueOnReboot -Confirm:$false }
        finally { $lockStream.Dispose() }
        Assert ($queued.QueuedFiles -gt 0 -and $queued.DeletedFiles -eq 0 -and $queued.DeletedBytes -eq 0 -and (Test-Path $queueFile)) 'Reboot mode queues locked/running-app cache files without deleting them now'
        Assert ($queued.QueuedBytes -gt 0 -and $queued.QueueOnReboot) 'Queued bytes and mode are reported separately from reclaimed space'
        $beforeCount=& $module { $script:queuedTestPaths.Count }
        Add-Content -LiteralPath $queueFile -Value 'changed'
        $changedQueue=& $module {param($f,$r) Remove-UnjunkCandidate $f $r -QueueOnReboot -Confirm:$false} $queueCandidate $root
        $afterCount=& $module { $script:queuedTestPaths.Count }
        Assert ($changedQueue.Status -eq 'Skipped' -and $afterCount -eq $beforeCount) 'Changed files cannot be queued'
        & $module { function script:Invoke-UnjunkNativeRebootDelete([string]$Path) { throw 'Simulated native scheduling failure' } }
        $failedQueue=& $module {param($f,$r) Remove-UnjunkCandidate $f $r -QueueOnReboot -Confirm:$false} (Candidate $queueFile) $root
        Assert ($failedQueue.Status -eq 'Skipped' -and $failedQueue.QueuedBytes -eq 0 -and (Test-Path $queueFile)) 'Native scheduling failure is not reported as queued or deleted'
    } finally {
        & $module {
            param($n,$a,$c)
            Set-Item Function:script:Invoke-UnjunkNativeRebootDelete -Value $n
            Set-Item Function:script:Test-UnjunkAdministrator -Value $a
            Set-Item Function:script:Get-UnjunkTargets -Value $c
            Remove-Variable fixtureTarget,queuedTestPaths -Scope Script
        } $originalNative $originalAdmin $originalCatalog
    }
    $rhinoRoot=Join-Path $testRoot 'Rhino9'
    $null=New-Item -ItemType Directory -Path $rhinoRoot,(Join-Path $rhinoRoot 'settings'),(Join-Path $rhinoRoot 'Plug-ins')
    $backup=Join-Path $rhinoRoot 'settings-backup.zip'; $backupUpper=Join-Path $rhinoRoot 'Settings-old.ZIP'
    $active=Join-Path $rhinoRoot 'settings.xml'; $nested=Join-Path $rhinoRoot 'settings\nested-backup.zip'; $plugin=Join-Path $rhinoRoot 'Plug-ins\plugin.zip'
    foreach($path in @($backup,$backupUpper,$active,$nested,$plugin)) { Set-Content -LiteralPath $path -Value 'disposable fixture' }
    $rhinoTarget=& $module {param($r) New-UnjunkTarget 'Rhino fixture' 'Creative apps' $r -Risk Review -FilePattern '*.zip' -Recurse $false} $rhinoRoot
    $rhinoScan=& $module {param($t) Get-UnjunkTargetScan $t 7 @() @{Cancel=$false} $false} $rhinoTarget
    Assert ($rhinoScan.FileCount -eq 2 -and @($rhinoScan.Files | Where-Object Path -in @($backup,$backupUpper)).Count -eq 2) 'Rhino ZIP target matches only immediate ZIP files, case-insensitively'
    Assert (-not $rhinoScan.Recommended) 'Rhino settings backups are not automatically recommended'
    $safeRhinoScan=& $module {param($t) Get-UnjunkTargetScan $t 7 @() @{Cancel=$false} $true} $rhinoTarget
    Assert ($safeRhinoScan.PolicyBlocked -and $safeRhinoScan.FileCount -eq 0) 'Rhino backup deletion is blocked by CrowdStrike Safe cleaning'
    $originalCatalog=& $module { (Get-Item Function:Get-UnjunkTargets).ScriptBlock }
    try {
        & $module {param($t) $script:fixtureTarget=$t; function script:Get-UnjunkTargets { @($script:fixtureTarget) }} $rhinoTarget
        $rhinoEnvelope=[pscustomobject]@{Cancelled=$false;UserSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value;Computer=$env:COMPUTERNAME;ScannedAt=[DateTime]::UtcNow.ToString('o');Targets=@($rhinoScan);Exclusions=@();CrowdStrikeSafe=$false}
        $originalFiles=$rhinoScan.Files
        $rhinoScan.Files=@($originalFiles)+(Candidate $active)
        $caught=$false; try { Invoke-UnjunkCleanup -Scan $rhinoEnvelope -TargetIds @($rhinoTarget.Id) -Confirm:$false } catch {$caught=$true}
        Assert ($caught -and (Test-Path $backup)) 'Cleanup rejects a non-ZIP injected into the backup selection before any deletion'
        $rhinoScan.Files=@($originalFiles)+(Candidate $nested)
        $caught=$false; try { Invoke-UnjunkCleanup -Scan $rhinoEnvelope -TargetIds @($rhinoTarget.Id) -Confirm:$false } catch {$caught=$true}
        Assert ($caught -and (Test-Path $backup)) 'Cleanup rechecks nonrecursive depth before any deletion'
        $rhinoScan.Files=$originalFiles
        $rhinoResult=Invoke-UnjunkCleanup -Scan $rhinoEnvelope -TargetIds @($rhinoTarget.Id) -Confirm:$false
        Assert ($rhinoResult.DeletedFiles -eq 2 -and (Test-Path $active) -and (Test-Path $nested) -and (Test-Path $plugin)) 'Rhino ZIP cleanup preserves active settings and all nested files'
    } finally {
        & $module {param($original) Set-Item Function:script:Get-UnjunkTargets -Value $original; Remove-Variable fixtureTarget -Scope Script} $originalCatalog
    }
    Write-Host "All $script:passed checks passed."
} finally {
    $resolved=[IO.Path]::GetFullPath($testRoot)
    $expected=[IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')+'\Unjunk-tests-'
    if(-not $resolved.StartsWith($expected,[StringComparison]::OrdinalIgnoreCase)) { throw 'Test cleanup path validation failed.' }
    $remainingLink=Join-Path $resolved 'cache\junction'
    if(Test-Path -LiteralPath $remainingLink) { [IO.Directory]::Delete($remainingLink) }
    Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction Stop
}
