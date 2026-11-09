#requires -Version 5.1
# Disposable in-memory fixtures only: no Windows cleanup, registry writes or app removal.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Windows.psm1') -Force
$module=Get-Module Unjunk.Windows
& $module {$script:maintenanceMutexName='Local\Unjunk-WindowsFixture-'+[Guid]::NewGuid().ToString('N')}
$passed=0
function Assert([bool]$Condition,[string]$Message) {if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
function Reject([scriptblock]$Action,[string]$Message) {$failed=$false;try{& $Action | Out-Null}catch{$failed=$true};Assert $failed $Message}
$apps=@([pscustomobject]@{Id='leftover-data:fixture';Name='Contoso.Removed';Version='Residual data';Scope='Profile: Fixture';Kind='DataLeftover';PackageName='Contoso.Removed_8wekyb3d8bbwe';Root='C:\Fixture\Packages\Contoso.Removed_8wekyb3d8bbwe'})
& $module { param($apps) $script:fixtureApps=$apps; function script:Get-UnjunkWindowsApps {param($Exclusions,$State) $script:fixtureApps} } $apps
& $module {
    $script:handlers=@('Temporary Files','Previous Installations','Temporary Setup Files','Update Cleanup','Delivery Optimization Files','DownloadsFolder','Recycle Bin')
    $script:nativeCalls=New-Object 'System.Collections.Generic.List[object]'
    $script:removals=New-Object 'System.Collections.Generic.List[string]'
    $script:registry=@{}
    foreach($name in $script:handlers){$script:registry[$name]=@{StateFlags0001=2}}
    $script:registry['Temporary Files'].StateFlags9000=2
    $script:failNative=$false; $script:writeCount=0; $script:failWriteAt=0; $script:admin=$true
    $script:shadows=@([pscustomobject]@{Id='{11111111-1111-1111-1111-111111111111}';Volume='\\?\Volume{22222222-2222-2222-2222-222222222222}\';Created='2026-11-01'})
    function script:Test-UnjunkWindowsAdmin {$script:admin}
    function script:Get-UnjunkDiskHandlers {$script:handlers}
    function script:Get-UnjunkShadowCopies {$script:shadows}
    function script:Get-Process {param($Name,$ErrorAction) @()}
    function script:Get-Item {param($LiteralPath,$ErrorAction)
        $item=[pscustomobject]@{Values=$script:registry[(Split-Path $LiteralPath -Leaf)]}
        $item | Add-Member ScriptMethod GetValueNames {return @($this.Values.Keys)} -PassThru
    }
    function script:New-ItemProperty {param($LiteralPath,$Name,$Value,$PropertyType,$ErrorAction)
        $script:writeCount++;if($script:failWriteAt -eq $script:writeCount){throw 'Simulated registry write failure'}
        $script:registry[(Split-Path $LiteralPath -Leaf)][$Name]=$Value
    }
    function script:Remove-ItemProperty {param($LiteralPath,$Name,$ErrorAction) $script:registry[(Split-Path $LiteralPath -Leaf)].Remove($Name)}
    function script:Invoke-UnjunkWindowsNative {param($Executable,$Arguments)
        $snapshot=@{};foreach($key in $script:registry.Keys){$snapshot[$key]=$script:registry[$key].Clone()}
        $script:nativeCalls.Add([pscustomobject]@{Executable=$Executable;Arguments=@($Arguments);Registry=$snapshot})
        if($script:failNative){throw 'Simulated native failure'}
        [pscustomobject]@{ExitCode=0;Output='Simulated Windows result';RestartRequired=$false}
    }
    function script:Invoke-UnjunkAppLeftoverCleanup {param($App,$State,[switch]$Confirm) $script:removals.Add($App.Id);[pscustomobject]@{ExitCode=0;Output='Fixture leftovers cleaned';RestartRequired=$false;ItemsCleaned=1L;FileResult=$null;RemovedProvisionedPackages=1L}}
    function script:Remove-AppxProvisionedPackage {param([switch]$Online,$PackageName,$ErrorAction) $script:removals.Add('provisioned:'+$PackageName);[pscustomobject]@{RestartNeeded=$false}}
    function script:Invoke-UnjunkEventLogCleanup {param($Logs,$State,[switch]$Confirm) [pscustomobject]@{ExitCode=0;Output='Fixture event logs';RestartRequired=$false;ItemsCleaned=1L;SkippedLogs=1}}
    function script:Invoke-UnjunkPrivacyCleanup {param($Exclusions,$State,[switch]$CloseApps,[switch]$Confirm) $script:privacyCloseApps=[bool]$CloseApps;[pscustomobject]@{ExitCode=0;Output='Fixture privacy';RestartRequired=$false;ItemsCleaned=3L;RemovedPrivacyRecords=2L;PrivacyBreakdown=@([pscustomobject]@{Name='IDM history';RemovedRecords=2});FileResult=[pscustomobject]@{DeletedFiles=1;DeletedBytes=10;SkippedFiles=0}}}
    function script:Invoke-UnjunkExplorerHistoryCleanup {param($Scan,$State,[switch]$Confirm) [pscustomobject]@{ExitCode=0;Output='Fixture Explorer recents';RestartRequired=$false;ItemsCleaned=2L;RemovedPrivacyRecords=2L;SkippedItems=1;FileResult=$null}}
    function script:Invoke-UnjunkCleanup {param($Scan,$TargetIds,[switch]$Confirm,[switch]$CloseApps,$State) [pscustomobject]@{DeletedBytes=123;DeletedFiles=1;SkippedFiles=1;Results=@();QueuedFiles=0;QueuedBytes=0}}
}
function Scan {
    [pscustomobject]@{ScannedAt=[DateTime]::UtcNow.ToString('o');Computer=$env:COMPUTERNAME;UserSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value;
        CrowdStrikeSafe=$false;Cancelled=$false;Handlers=@(& $module {Get-UnjunkDiskHandlers});Apps=$apps;AppError='';Shadows=@(& $module {Get-UnjunkShadowCopies});ShadowError='';PrivacyError='';ExplorerHistoryError='';ExplorerHistory=[pscustomobject]@{RecentFiles=2;FoldersKept=4;FavoritesKept=1};EventLogError='';EventLogs=@([pscustomobject]@{Name='Security';RecordCount=10});Privacy=[pscustomobject]@{RecentLinks=1;RegistryRecords=2};
        FileScan=[pscustomobject]@{Exclusions=@();Targets=@([pscustomobject]@{Id='temp';SystemOption='system-temp';FileCount=1;Bytes=123},[pscustomobject]@{Id='kernel';SystemOption='live-kernel';FileCount=1;Bytes=456},[pscustomobject]@{Id='memory';SystemOption='memory-dumps';FileCount=1;Bytes=789},[pscustomobject]@{Id='cache';SystemOption='app-caches';FileCount=1;Bytes=100},[pscustomobject]@{Id='logs';SystemOption='app-logs';FileCount=1;Bytes=50},[pscustomobject]@{Id='autodesk';SystemOption='autodesk-installers';FileCount=1;Bytes=75})}}
}
$scan=Scan
$options=@(Get-UnjunkWindowsOptions)
Assert (@($options | Where-Object Id -in @('live-kernel','memory-dumps','shadow-copies')).Count -eq 3) 'Diagnostics and shadow copies have separate selectable options'
$plan=@(Get-UnjunkWindowsPlan -Scan $scan -OptionIds $options.Id)
$ordered=@(Get-UnjunkWindowsPlan -Scan $scan -OptionIds $options.Id -AppIds $apps.Id)
Assert ($ordered[0].Id -eq 'app-caches' -and $ordered[1].Kind -eq 'AppLeftovers' -and $ordered[-3].Kind -eq 'Shadows' -and $ordered[-2].Kind -eq 'DISM' -and $ordered[-1].Kind -eq 'DiskCleanup') 'Sequence is junk, app leftovers, temp/logs/dumps, shadow copies, DISM, then Disk Cleanup'
Assert ($ordered[-4].Kind -eq 'ExplorerHistory' -and $ordered[-5].Kind -eq 'Privacy') 'Explorer Home is cleared after private history and before native maintenance'
$oldScan=Scan;$oldScan.Apps=@([pscustomobject]@{Id='installed:old';Name='Contoso.Active';Kind='Installed'})
Reject {Get-UnjunkWindowsPlan -Scan $oldScan -AppIds 'installed:old'} 'A saved scan from the old uninstall UI cannot remove an installed app'
Assert (@($plan | Where-Object Id -in @('old-windows','windows-update')).Count -eq 0) 'Full cleanup absorbs overlapping old-installation and update tasks'
Assert (@($plan | Where-Object Id -eq 'shadow-copies').Count -eq 1) 'Clean all includes the explicit shadow-copy operation'
Assert (@($plan | Where-Object Kind -eq 'AppLeftovers').Count -eq 0) 'Clean all does not implicitly select app leftovers'
$full=@($plan | Where-Object Id -eq 'disk-full')[0]
Assert ('DownloadsFolder' -in $full.Handlers -and $full.Detail -match 'all drives' -and $full.Detail -match 'exclusions do not apply') 'Full cleanup previews every category, drive scope and native eligibility rules'
$original=@(Get-UnjunkWindowsPlan -Scan $scan -OptionIds 'disk-full' -DiskCleanupProfile Original)[0]
Assert ('DownloadsFolder' -notin $original.Handlers -and 'Previous Installations' -in $original.Handlers -and 'Update Cleanup' -in $original.Handlers) 'One-click original profile includes deep cleanup while preserving Downloads'
Reject {Get-UnjunkWindowsPlan -Scan $scan -OptionIds 'made-up'} 'Unknown cleanup options are rejected'
Reject {Get-UnjunkWindowsPlan -Scan $scan -AppIds 'installed:invented'} 'Unknown app identities are rejected'
$scan.ShadowError='Access denied';Reject {Get-UnjunkWindowsPlan -Scan $scan -OptionIds 'shadow-copies'} 'Failed shadow inventory cannot be treated as an empty successful cleanup';$scan.ShadowError=''
$scan.CrowdStrikeSafe=$true;Reject {Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'disk-full' -Confirm:$false} 'Engine blocks Windows cleanup in CrowdStrike Safe mode';$scan.CrowdStrikeSafe=$false
& $module {$script:admin=$false};Reject {Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'winsxs' -Confirm:$false} 'Engine requires administrator access'; & $module {$script:admin=$true}
$scan.ScannedAt=[DateTime]::UtcNow.AddMinutes(-31).ToString('o');Reject {Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'winsxs' -Confirm:$false} 'Expired scans cannot be cleaned';$scan=Scan
$scan.UserSid='another-account';Reject {Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'winsxs' -Confirm:$false} 'Cross-account scans cannot be cleaned';$scan=Scan
$result=Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds $options.Id -AppIds $apps.Id -WhatIf
Assert (@($result.Results | Where-Object Status -ne 'Preview').Count -eq 0) 'WhatIf returns a complete preview plan'
Assert ((& $module {$script:nativeCalls.Count+$script:removals.Count+$script:writeCount}) -eq 0) 'WhatIf does not invoke Windows, remove apps or change the registry'
& $module {$script:handlers+=@('New handler')};Reject {Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'disk-full' -Confirm:$false} 'Changed full-cleanup categories require a new reviewed scan'; & $module {$script:handlers=@($script:handlers | Where-Object {$_ -ne 'New handler'})}
$result=Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'winsxs','live-kernel','memory-dumps' -Confirm:$false
Assert ($result.Completed -eq 1 -and $result.Partial -eq 2) 'Partial file cleanup is reported separately from successful maintenance'
Assert ($result.Results[0].Kind -eq 'Files' -and $result.Results[0].ItemsCleaned -eq 1 -and $result.Results[-1].Kind -eq 'DISM' -and $null -eq $result.Results[-1].ItemsCleaned) 'File counts and native unknown counts are recorded, with files preceding long servicing'
Assert ([DateTime]::Parse($result.Results[0].CompletedAt) -ge [DateTime]::Parse($result.Results[0].StartedAt)) 'Each cleanup period records real start and finish timestamps'
$dism=@(& $module {$script:nativeCalls | Where-Object Executable -eq 'Dism.exe'})[0]
Assert ('/StartComponentCleanup' -in $dism.Arguments -and '/NoRestart' -in $dism.Arguments -and '/ResetBase' -notin $dism.Arguments) 'WinSxS performs component cleanup without ResetBase or an automatic restart'
$result=Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'windows-update' -Confirm:$false
$native=@(& $module {$script:nativeCalls | Where-Object Executable -eq 'cleanmgr.exe'})[-1]
Assert ($native.Arguments[0] -eq '/sagerun:9001') 'Unattended cleanup reserves an unused profile'
Assert ($native.Registry['Update Cleanup'].StateFlags9001 -eq 2 -and $native.Registry['DownloadsFolder'].StateFlags9001 -eq 0) 'Targeted update cleanup does not include Downloads or unrelated categories'
Assert ((& $module {$script:registry['Temporary Files'].StateFlags9000}) -eq 2 -and (& $module {$script:registry['Update Cleanup'].StateFlags0001}) -eq 2) 'Existing Disk Cleanup profiles remain unchanged'
Assert ((& $module {@($script:registry.Values | Where-Object {$_.ContainsKey('StateFlags9001')}).Count}) -eq 0) 'Temporary profile values are removed after completion'
$result=Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'disk-full' -DiskCleanupProfile Original -Confirm:$false
$native=@(& $module {$script:nativeCalls | Where-Object Executable -eq 'cleanmgr.exe'})[-1]
Assert ($result.Completed -eq 1 -and $native.Registry['DownloadsFolder'].StateFlags9001 -eq 0 -and $native.Registry['Previous Installations'].StateFlags9001 -eq 2) 'Original profile is honored by the executing registry profile'
& $module {$script:failNative=$true}
$result=Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'disk-full' -Confirm:$false
Assert ($result.Failed -eq 1 -and (& $module {@($script:registry.Values | Where-Object {$_.ContainsKey('StateFlags9001')}).Count}) -eq 0) 'Native failures are reported and temporary profiles are restored'
& $module {$script:failNative=$false;$script:failWriteAt=$script:writeCount+3}
$before=& $module {$script:nativeCalls.Count}
$result=Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'disk-full' -Confirm:$false
Assert ($result.Failed -eq 1 -and (& $module {$script:nativeCalls.Count}) -eq $before) 'A partial profile write never starts cleanup'
Assert ((& $module {@($script:registry.Values | Where-Object {$_.ContainsKey('StateFlags9001')}).Count}) -eq 0) 'Partial profile writes are rolled back'; & $module {$script:failWriteAt=0}
$result=Invoke-UnjunkWindowsCleanup -Scan $scan -AppIds $apps.Id -Confirm:$false
Assert ($result.Completed -eq 1 -and (& $module {$script:removals.Count}) -eq 1) 'App cleanup dispatches only the selected leftover identity'
$result=Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'privacy' -CloseApps -Confirm:$false
Assert ($result.Completed -eq 1 -and $result.Results[0].ItemsCleaned -eq 3 -and $result.Results[0].RemovedPrivacyRecords -eq 2 -and $result.Results[0].FileResult.DeletedFiles -eq 1) 'Privacy stage reports removed links and registry records through the Windows engine'
Assert ((& $module {$script:privacyCloseApps}) -and $result.Results[0].PrivacyBreakdown[0].Name -eq 'IDM history') 'Windows engine forwards app closing and retains the privacy breakdown'
& $module {function script:Invoke-UnjunkPrivacyCleanup {param($Exclusions,$State,[switch]$CloseApps,[switch]$Confirm) [pscustomobject]@{ExitCode=0;Output='Fixture locked registry store';RestartRequired=$false;ItemsCleaned=0L;RemovedPrivacyRecords=0L;SkippedItems=1;FileResult=[pscustomobject]@{DeletedFiles=0;DeletedBytes=0;SkippedFiles=0}}}}
$result=Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'privacy' -Confirm:$false
Assert ($result.Partial -eq 1 -and $result.Failed -eq 0 -and $result.Results[0].FileResult.SkippedFiles -eq 0) 'Locked registry or database stores report partial without inventing skipped files'
$result=Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'explorer-recents' -Confirm:$false
Assert ($result.Partial -eq 1 -and $result.Results[0].ItemsCleaned -eq 2 -and $result.Results[0].RemovedPrivacyRecords -eq 2 -and $null -eq $result.Results[0].FileResult) 'Explorer removals count as history records and partial verification stays visible'
$result=Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'event-viewer' -Confirm:$false
Assert ($result.Partial -eq 1 -and $result.Results[0].ItemsCleaned -eq 1 -and $result.Results[0].Kind -eq 'EventLogs') 'Denied event channels produce a partial result while confirmed clears remain counted'
& $module {$script:shadows+=@([pscustomobject]@{Id='{33333333-3333-3333-3333-333333333333}';Volume='\\?\Volume{22222222-2222-2222-2222-222222222222}\';Created='2026-11-06'})}
$result=Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'shadow-copies' -Confirm:$false
$vss=@(& $module {$script:nativeCalls | Where-Object Executable -eq 'vssadmin.exe'})
Assert ($result.Completed -eq 1 -and $vss.Count -eq 1 -and $vss[0].Arguments[2] -eq '/Shadow={11111111-1111-1111-1111-111111111111}') 'Shadow cleanup deletes only scanned snapshot IDs, preserving snapshots created later'
Assert (@($vss[0].Arguments | Where-Object {$_ -like '/For=*'}).Count -eq 0) 'Shadow-ID deletion does not combine the incompatible volume selector'
Assert ('/Quiet' -in $vss[0].Arguments -and '/all' -notin $vss[0].Arguments) 'Shadow cleanup runs quietly without a broad delete-all switch'
$state=@{Cancel=$true;Progress=0;Message=''};$before=& $module {$script:nativeCalls.Count}
$result=Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds 'winsxs' -State $state -Confirm:$false
Assert ($result.Cancelled -and $result.Results.Count -eq 0 -and (& $module {$script:nativeCalls.Count}) -eq $before) 'Stop prevents subsequent operations without terminating Windows servicing'
Write-Host "All $passed Windows maintenance checks passed."
