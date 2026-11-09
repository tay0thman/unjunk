#requires -Version 5.1
# In-memory Windows providers and disposable reports only. Never runs real cleanup.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.OneClick.psm1') -Force
$module=Get-Module Unjunk.OneClick
& $module {$script:oneClickMutexName='Local\Unjunk-OneClickFixture-'+[Guid]::NewGuid().ToString('N')}
$passed=0
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
function App($Name,$Version,$Architecture='x64',$Kind='Installed',$Publisher='publisher') {
    $package=$Name+'_'+$Version+'_'+$Architecture+'__'+$Publisher
    [pscustomobject]@{Name=$Name;Version=$Version;PackageName=$package;Kind=$Kind;Id=$Kind+':'+$package}
}
$apps=@((App 'Microsoft.BingNews' '1.0.0.0'),(App 'Microsoft.WindowsTerminal' '1.0.0.0'),(App 'OpenAI.ChatGPT-Desktop' '1.0.0.0'),(App 'Contoso.Important' '1.0.0.0'),(App 'Microsoft.MicrosoftStickyNotes' '1.0.0.0' 'x64' 'DataLeftover'),(App 'Microsoft.MicrosoftStickyNotes' '1.0.0.0' 'x64' 'ProvisionedLeftover'))
$ids=@(Get-UnjunkOriginalAppIds -Apps $apps)
Assert ($ids.Count -eq 2 -and $apps[4].Id -in $ids -and $apps[5].Id -in $ids) 'One-click selects only already-uninstalled app data and unused provisioning'
Assert (@($apps[0..3] | Where-Object {$_.Id -in $ids}).Count -eq 0) 'One-click never selects installed apps, including former bloatware targets'
Assert (@(Get-UnjunkOriginalAppIds -Apps @((App 'Microsoft.WindowsTerminal' '1.0.0.0' 'x64' 'DataLeftover'),(App 'OpenAI.ChatGPT-Desktop' '1.0.0.0' 'x64' 'ProvisionedLeftover'))).Count -eq 0) 'Terminal and ChatGPT are protected even when supplied as leftovers'
$previousLocal=$env:LOCALAPPDATA
$fixtureRoot=Join-Path $PSScriptRoot ('OneClickFixture-'+[Guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $fixtureRoot
$env:LOCALAPPDATA=$fixtureRoot
try {
    $null=New-Item -ItemType Directory -Path (Join-Path $fixtureRoot 'Unjunk')
    $settingsPath=Join-Path $fixtureRoot 'Unjunk\settings.json'
    '{"CrowdStrikeSafe":true,"TempAgeDays":7}' | Set-Content -LiteralPath $settingsPath
    $beforeHash=(Get-FileHash -LiteralPath $settingsPath).Hash
    & $module {
        $script:scanCalls=0; $script:cleanCalls=0; $script:spaceCalls=0; $script:admin=$true; $script:busy=$false; $script:inventoryError=$false; $script:fail=$false
        function script:Test-UnjunkOneClickAdmin {$script:admin}
        function script:Test-UnjunkWindowsCleanupBusy {$false}
        function script:Get-Process {param($Name,$ErrorAction) if($script:busy){[pscustomobject]@{Id=123}}}
        function script:Get-UnjunkDiskSpace {
            $script:spaceCalls++
            [pscustomobject]@{Drive=($env:SystemDrive+'\');FreeBytes=if($script:spaceCalls%2){1000L}else{700L};TotalBytes=100000L}
        }
        function script:Invoke-UnjunkWindowsScan {
            param($CrowdStrikeSafe,$TempAgeDays,$Exclusions,$State)
            $script:scanCalls++; $script:scanArgs=@{Safe=$CrowdStrikeSafe;Days=$TempAgeDays;Exclusions=$Exclusions}
            [pscustomobject]@{Apps=@();PrivacyError='';EventLogError='';AppError=if($script:inventoryError){'Fixture inventory failure'}else{''};ShadowError=if($script:inventoryError){'Fixture shadow failure'}else{''};FileScan=[pscustomobject]@{Targets=@()}}
        }
        function script:Invoke-UnjunkWindowsCleanup {
            param($Scan,$OptionIds,$AppIds,$DiskCleanupProfile,$State,[switch]$CloseApps,[switch]$Confirm)
            $script:cleanCalls++; $script:cleanArgs=@{Options=$OptionIds;Profile=$DiskCleanupProfile;CloseApps=[bool]$CloseApps;Confirm=$Confirm.IsPresent}
            if($script:fail){throw 'Fixture cleanup failure'}
            $time=[DateTime]::UtcNow.ToString('o')
            [pscustomobject]@{Completed=4;Partial=0;Failed=0;Cancelled=$false;RestartRequired=$true;Results=@(
                [pscustomobject]@{Name='Temp';Kind='Files';ItemsCleaned=12L;FileResult=[pscustomobject]@{DeletedFiles=12L;DeletedBytes=456L};StartedAt=$time;CompletedAt=$time;Status='Completed';Output='Fixture'},
                [pscustomobject]@{Name='Unused provisioned copy';Kind='AppLeftovers';ItemsCleaned=1L;RemovedProvisionedPackages=1L;FileResult=$null;StartedAt=$time;CompletedAt=$time;Status='Completed';Output='Fixture'},
                [pscustomobject]@{Name='Shadow';Kind='Shadows';ItemsCleaned=2L;FileResult=$null;StartedAt=$time;CompletedAt=$time;Status='Completed';Output='Fixture'},
                [pscustomobject]@{Name='Privacy';Kind='Privacy';ItemsCleaned=2L;RemovedPrivacyRecords=2L;PrivacyBreakdown=@([pscustomobject]@{Name='IDM history';RemovedRecords=2});FileResult=$null;StartedAt=$time;CompletedAt=$time;Status='Completed';Output='Fixture'},
                [pscustomobject]@{Name='Explorer Home';Kind='ExplorerHistory';ItemsCleaned=1L;RemovedPrivacyRecords=1L;FileResult=$null;StartedAt=$time;CompletedAt=$time;Status='Completed';Output='Fixture'},
                [pscustomobject]@{Name='Event Viewer';Kind='EventLogs';ItemsCleaned=3L;FileResult=$null;StartedAt=$time;CompletedAt=$time;Status='Completed';Output='Fixture'},
                [pscustomobject]@{Name='Native';Kind='DiskCleanup';ItemsCleaned=$null;FileResult=$null;StartedAt=$time;CompletedAt=$time;Status='Completed';Output='Fixture'}
            )}
        }
    }
    $run=Invoke-UnjunkOneClick -Exclusions @('C:\Keep')
    $argsUsed=& $module {$script:scanArgs}
    Assert (-not $argsUsed.Safe -and $argsUsed.Days -eq 0 -and 'C:\Keep' -in $argsUsed.Exclusions) 'Full preset overrides Safe only for the run and keeps file exclusions'
    Assert ((Get-FileHash -LiteralPath $settingsPath).Hash -eq $beforeHash) 'Saved CrowdStrike Safe preference is unchanged'
    $cleanupArgs=& $module {$script:cleanArgs}
    Assert ($cleanupArgs.Profile -eq 'Full' -and -not $cleanupArgs.Confirm -and $cleanupArgs.CloseApps -and $cleanupArgs.Options.Count -eq @(& $module {Get-UnjunkWindowsOptions}).Count -and 'privacy' -in $cleanupArgs.Options -and 'autodesk-installers' -in $cleanupArgs.Options -and 'event-viewer' -in $cleanupArgs.Options) 'One-click selects every live cleanup category, full Disk Cleanup and app closing'
    Assert ($run.CountedItems -eq 21 -and $run.ClearedEventLogs -eq 3 -and $run.RemovedPrivacyRecords -eq 3 -and $run.DeletedFiles -eq 12 -and $run.RemovedApps -eq 0 -and $run.RemovedProvisionedPackages -eq 1 -and $run.RemovedShadows -eq 2) 'History counts files, privacy, Explorer recents, cleared event channels and leftovers without claiming app uninstallation'
    Assert ('explorer-recents' -in $cleanupArgs.Options) 'One-click includes the authorized Explorer Home exception'
    Assert ($run.UnreportedItemCounts -and $null -eq $run.Periods[-1].ItemsCleaned) 'Native cleanup does not invent an item count'
    Assert ($run.Disks[0].NetFreedBytes -eq -300 -and $run.Disks[0].FinalFreeBytes -eq 700) 'History preserves negative disk changes and actual final free space'
    Assert ($run.Status -eq 'Completed' -and $run.RestartRequired -and (Test-Path -LiteralPath $run.ReportPath)) 'Completion preserves restart requirements and saves the detailed report'
    Assert (@(Get-UnjunkHistory).Count -eq 1) 'Completed run can be reloaded from persisted history'
    $privacyPeriod=@((Get-UnjunkHistory).Periods | Where-Object Kind -eq 'Privacy')[0]
    Assert ($privacyPeriod.PrivacyBreakdown[0].Name -eq 'IDM history' -and $privacyPeriod.PrivacyBreakdown[0].RemovedRecords -eq 2) 'One-click timeline persists the explicit privacy group counts'
    $historyDirectory=Join-Path $fixtureRoot 'Unjunk\History'
    '{broken json' | Set-Content -LiteralPath (Join-Path $historyDirectory 'run-broken.json')
    Assert (@(Get-UnjunkHistory).Count -eq 1) 'A damaged history file does not hide valid runs'
    & $module {$script:inventoryError=$true}
    $partial=Invoke-UnjunkOneClick
    Assert ($partial.Status -eq 'Partial' -and $partial.Warnings.Count -eq 2) 'Inventory failures produce a partial result while independent work completes'
    Assert ('shadow-copies' -notin (& $module {$script:cleanArgs.Options})) 'Failed shadow inventory never becomes an empty success'
    & $module {$script:fail=$true}
    $failed=Invoke-UnjunkOneClick
    Assert ($failed.Status -eq 'Failed' -and $failed.Error -match 'Fixture cleanup failure' -and @(Get-UnjunkHistory).Count -eq 3) 'Failed cleanup is recorded for the timeline without a false success'
    & $module {$script:busy=$true}
    $rejected=$false;try{$null=Invoke-UnjunkOneClick}catch{$rejected=$_.Exception.Message -match 'already running'}
    Assert $rejected 'An existing Disk Cleanup blocks a second one-click launch before cleanup starts'
    Assert ((& $module {$script:cleanCalls}) -eq 3) 'Busy check performs no cleanup'
    & $module {$script:busy=$false;$script:admin=$false}
    $rejected=$false;try{$null=Invoke-UnjunkOneClick}catch{$rejected=$true}
    Assert $rejected 'One-click engine requires administrator access'
    & $module {
        function script:Start-Process {param($FilePath,$ArgumentList,$WindowStyle,[switch]$PassThru,$ErrorAction,$Verb)
            [pscustomobject]@{Path=$FilePath;Arguments=$ArgumentList;Style=$WindowStyle;Verb=$Verb}
        }
    }
    $launch=Start-UnjunkOneClick
    Assert ($launch.Style -eq 'Hidden' -and $launch.Arguments -match '-STA' -and $launch.Arguments -match '-ExpectedUserSid') 'Detached worker launch uses STA, hidden console and the originating user identity'
    $history=Start-UnjunkOneClick -History
    Assert ($history.Arguments -match '-History' -and -not $history.Verb) 'History opens independently without elevation or cleanup'
    & $module {
        $script:baseOptions=@(Get-UnjunkWindowsOptions)
        function script:Get-UnjunkWindowsOptions {$script:baseOptions;[pscustomobject]@{Id='future-cleanup'}}
        $script:admin=$true;$script:inventoryError=$false;$script:fail=$false
    }
    $future=Invoke-UnjunkOneClick -CloseApps $false
    Assert ('future-cleanup' -in (& $module {$script:cleanArgs.Options}) -and -not (& $module {$script:cleanArgs.CloseApps})) 'New categories join one-click automatically and the app-closing preference is honored'
} finally {
    $env:LOCALAPPDATA=$previousLocal
    $full=[IO.Path]::GetFullPath($fixtureRoot); $testRoot=[IO.Path]::GetFullPath($PSScriptRoot)+'\'
    if(-not $full.StartsWith($testRoot,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $full -Leaf) -notlike 'OneClickFixture-*'){throw 'Unsafe fixture path'}
    Remove-Item -LiteralPath $full -Recurse -Force
}
Write-Host "All $passed one-click checks passed."
