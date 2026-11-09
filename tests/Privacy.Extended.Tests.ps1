#requires -Version 5.1
# File fixtures and an in-memory registry only. Never modifies real history or stops an app.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Privacy.psm1') -Force
$module=Get-Module Unjunk.Privacy
$passed=0
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using Microsoft.Win32;
public class UnjunkPrivacyFixtureKey {
    public Dictionary<string,object> Values=new Dictionary<string,object>(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string,RegistryValueKind> Kinds=new Dictionary<string,RegistryValueKind>(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string,UnjunkPrivacyFixtureKey> Children=new Dictionary<string,UnjunkPrivacyFixtureKey>(StringComparer.OrdinalIgnoreCase);
    public bool Denied=false;
    public int ValueCount {get{return Values.Count;}}
    public int SubKeyCount {get{return Children.Count;}}
    public string[] GetValueNames(){if(Denied)throw new UnauthorizedAccessException("fixture");return new List<string>(Values.Keys).ToArray();}
    public string[] GetSubKeyNames(){if(Denied)throw new UnauthorizedAccessException("fixture");return new List<string>(Children.Keys).ToArray();}
    public object GetValue(string name){object value;return Values.TryGetValue(name,out value)?value:null;}
    public object GetValue(string name,object fallback,RegistryValueOptions options){return GetValue(name)??fallback;}
    public RegistryValueKind GetValueKind(string name){return Kinds[name];}
    public void SetValue(string name,object value,RegistryValueKind kind){Values[name]=value;Kinds[name]=kind;}
    public void DeleteValue(string name,bool missing){Values.Remove(name);Kinds.Remove(name);}
    public UnjunkPrivacyFixtureKey OpenSubKey(string name,bool writable){UnjunkPrivacyFixtureKey key;return Children.TryGetValue(name,out key)?key:null;}
    public void DeleteSubKeyTree(string name,bool missing){Children.Remove(name);}
    public void Dispose(){}
}
'@
function New-Key($Values){
    $key=New-Object UnjunkPrivacyFixtureKey
    foreach($name in $Values.Keys){$value=$Values[$name];$kind=if($value -is [byte[]]){'Binary'}elseif($value -is [int]){'DWord'}else{'String'};$key.SetValue($name,$value,$kind)}
    return $key
}
function New-Download([int]$Status=3,[int]$Saved=1,[long]$Downloaded=100){
    $key=New-Key @{Status=$Status;bSaved=$Saved;FileSize=[BitConverter]::GetBytes([long]100);Downloaded=[BitConverter]::GetBytes($Downloaded);Url0='private-download-url';LocalFileName='private-download-name'}
    $key.Children['ChList']=New-Key @{Cookie='private-cookie';EncPassword='private-credential'}
    return $key
}
$fixture=Join-Path $PSScriptRoot ('PrivacyExtendedFixture-'+[Guid]::NewGuid().ToString('N'))
$previousAppData=$env:APPDATA
$env:APPDATA=Join-Path $fixture 'Roaming'
$recent=Join-Path $env:APPDATA 'Microsoft\Windows\Recent';$prefetch=Join-Path $fixture 'Windows\Prefetch';$idm=Join-Path $env:APPDATA 'IDM'
$null=New-Item -ItemType Directory -Path $recent,$prefetch,$idm,(Join-Path $recent 'AutomaticDestinations'),(Join-Path $recent 'CustomDestinations'),(Join-Path $prefetch 'ReadyBoot'),(Join-Path $idm 'DwnlData') -Force
$files=@{Recent=(Join-Path $recent 'private.lnk');Auto=(Join-Path $recent 'AutomaticDestinations\pinned.automaticDestinations-ms');Custom=(Join-Path $recent 'CustomDestinations\pinned.customDestinations-ms');Prefetch=(Join-Path $prefetch 'APP.EXE-1234.pf');Excluded=(Join-Path $prefetch 'KEEP.EXE-1234.pf');Layout=(Join-Path $prefetch 'Layout.ini');Boot=(Join-Path $prefetch 'ReadyBoot\keep.pf');Url=(Join-Path $idm 'UrlHistory.txt');Folder=(Join-Path $idm 'foldresHistory.txt');Temp=(Join-Path $idm 'DwnlData\unfinished.part');Settings=(Join-Path $idm 'defextmap.dat');Downloaded=(Join-Path $fixture 'downloaded.zip')}
foreach($path in $files.Values){Set-Content -LiteralPath $path -Value 'private-fixture-content'}
$hashes=@{};foreach($name in @('Auto','Custom','Layout','Boot','Temp','Settings','Downloaded')){$hashes[$name]=(Get-FileHash -LiteralPath $files[$name]).Hash}
$registry=@{}
$base='Software\Microsoft\Windows\CurrentVersion\Explorer'
$registry[$base+'\RecentDocs']=New-Key @{MRUListEx=[byte[]](0,0,0,0);'0'='private-document'}
$registry[$base+'\WordWheelQuery']=New-Key @{MRUListEx=[byte[]](0,0,0,0);'0'='private-search'}
$userRoot=New-Key @{};$guid='{11111111-1111-1111-1111-111111111111}';$userGuid=New-Key @{};$count=New-Key @{'private-app'=1}
$userRoot.Children[$guid]=$userGuid;$userGuid.Children['Count']=$count
$registry[$base+'\UserAssist']=$userRoot;$registry[$base+'\UserAssist\'+$guid]=$userGuid;$registry[$base+'\UserAssist\'+$guid+'\Count']=$count
$recentApps='Software\Microsoft\Windows\CurrentVersion\Search\RecentApps';$apps=New-Key @{TrackingPreference=1};$app=New-Key @{AppId='private-app';LastAccessedTime=123};$items=New-Key @{};$item=New-Key @{Path='private-recent-file'}
$apps.Children[$guid]=$app;$app.Children['RecentItems']=$items;$items.Children[$guid]=$item
$registry[$recentApps]=$apps;$registry[$recentApps+'\'+$guid]=$app;$registry[$recentApps+'\'+$guid+'\RecentItems']=$items;$registry[$recentApps+'\'+$guid+'\RecentItems\'+$guid]=$item
$muiPath='Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache'
$mui=New-Key @{LangID='language';'private.exe.FriendlyAppName'='private-program';'private.exe.ApplicationCompany'='private-company'};$registry[$muiPath]=$mui
$pcaPath='Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Store';$pca=New-Key @{'private.exe'=[byte[]](1,2,3)};$registry[$pcaPath]=$pca
$downloads=New-Key @{Registration='keep-license';AppDataIDMFolder=$idm;LaunchOnStart=1}
$downloads.Children['100']=New-Download;$downloads.Children['101']=New-Download 1 0 50;$downloads.Children['102']=New-Download 3 1 50
$downloads.Children['103']=New-Download 3 0 100;$downloads.Children['Settings']=New-Key @{Important='keep'}
$registry['Software\DownloadManager']=$downloads
$bagsPath='Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags'
$bagMruPath='Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\BagMRU'
$bagMru=New-Key @{MRUListEx=[byte[]](0,0,0,0);NodeSlots=[byte[]](1,2);'0'=[byte[]](3,4)}
$bagChild=New-Key @{MRUListEx=[byte[]](0,0,0,0);NodeSlot=0;'0'=[byte[]](5,6)};$bagMru.Children['0']=$bagChild
$registry[$bagMruPath]=$bagMru;$registry[$bagMruPath+'\0']=$bagChild
$bags=New-Key @{};$bagZero=New-Key @{};$bagView=New-Key @{ViewMode=1;Sort=[byte[]](1,2)};$defaults=New-Key @{KeepDefaults=1}
$bags.Children['0']=$bagZero;$bags.Children['AllFolders']=$defaults;$bagZero.Children['Shell']=$bagView
$registry[$bagsPath]=$bags;$registry[$bagsPath+'\0']=$bagZero;$registry[$bagsPath+'\0\Shell']=$bagView;$registry[$bagsPath+'\AllFolders']=$defaults
$sid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$bamPath='HKLM\SYSTEM\CurrentControlSet\Services\bam\State\UserSettings\'+$sid
$damPath='HKLM\SYSTEM\CurrentControlSet\Services\dam\State\UserSettings\'+$sid
$bam=New-Key @{'\Device\HarddiskVolume3\private.exe'=[byte[]](1,2,3);Version=1;SequenceNumber=9}
$dam=New-Key @{'C:\private.exe'=[byte[]](1,2,3);Version=1;SequenceNumber=9};$registry[$bamPath]=$bam;$registry[$damPath]=$dam
$shimPath='HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\AppCompatCache'
$shim=New-Key @{AppCompatCache=[byte[]](1,2,3);CacheMainSdb=[byte[]](9,9);SdbTime=123};$registry[$shimPath]=$shim
$feature=New-Key @{'private.exe'=2};$registry[$base+'\FeatureUsage\AppSwitched']=$feature
$mpc=New-Key @{File1='private-media-file'};$registry['Software\MPC-HC\MPC-HC\Recent File List']=$mpc
$pcaFolder=Join-Path $fixture 'Windows\appcompat\pca';$null=New-Item -ItemType Directory -Path $pcaFolder -Force
foreach($name in @('PcaAppLaunchDic.txt','PcaGeneralDb0.txt','PcaGeneralDb1.txt','keep.dll')){Set-Content -LiteralPath (Join-Path $pcaFolder $name) -Value 'private-compatibility-fixture'}
& $module {param($registry)
    $script:fixtureRegistry=$registry;$script:idmRunning=$false;$script:closeCount=0;$script:stayRunning=$false
    function script:Open-UnjunkPrivacyRegistryKey {param($Path,$Writable) return $script:fixtureRegistry[$Path]}
    function script:Get-UnjunkPrivacyWindowsRoot {Join-Path (Split-Path $env:APPDATA -Parent) 'Windows'}
    function script:Get-UnjunkActivityInventory {param($Exclusions) [pscustomobject]@{Databases=@();Issues=@()}}
    function script:Test-UnjunkIdmRunning {return $script:idmRunning}
    function script:Close-UnjunkIdm {
        param($State)
        $script:closeCount++
        if(-not $script:stayRunning){$script:idmRunning=$false}
        Set-Content -LiteralPath (Join-Path $env:APPDATA 'IDM\UrlHistory.txt') -Value 'private-saved-at-exit'
        [pscustomobject]@{Name='IDMan';ProcessId=123;Status=if($script:stayRunning){'Skipped'}else{'Closed'};Reason='fixture'}
    }
} $registry
try {
    $preview=Get-UnjunkPrivacyPreview -Exclusions @($files.Excluded)
    Assert ($preview.RecentLinks -eq 1 -and $preview.PrefetchFiles -eq 1 -and $preview.IdmDownloads -eq 1 -and $preview.IdmHistoryFiles -eq 2) 'Preview finds Windows history, Prefetch and only completed IDM downloads'
    Assert ($preview.RegistryRecords -eq 24 -and $preview.Warnings.Count -eq 0) 'Windows registry inventory includes visible and hidden history stores'
    Assert ($preview.HiddenRegistryRecords -eq 11 -and $preview.CompatibilityFiles -eq 3) 'Hidden-history preview counts ShellBags, BAM/DAM, compatibility cache and PCA logs'
    Assert (($preview|ConvertTo-Json -Depth 8) -notmatch 'private-|Cookie|credential|Hash|100') 'Preview contains counts without private data or record identifiers'
    & $module {$script:idmRunning=$true}
    $whatIf=Invoke-UnjunkPrivacyCleanup -CloseApps -WhatIf
    Assert ((& $module {$script:closeCount}) -eq 0 -and (Test-Path -LiteralPath $files.Url) -and $downloads.Children.ContainsKey('100')) 'WhatIf does not close IDM or clear any history'
    $cancel=Invoke-UnjunkPrivacyCleanup -CloseApps -State @{Cancel=$true;Progress=0;Message=''} -Confirm:$false
    Assert ((& $module {$script:closeCount}) -eq 0 -and $cancel.ItemsCleaned -eq 0) 'Cancellation performs no process or history changes'
    $blocked=Invoke-UnjunkPrivacyCleanup -Exclusions @($files.Excluded) -Confirm:$false
    Assert ($blocked.SkippedItems -gt 0 -and $downloads.Children.ContainsKey('100') -and (Test-Path -LiteralPath $files.Url)) 'Running IDM is skipped when app closing is disabled'
    Assert (-not (Test-Path -LiteralPath $files.Recent) -and -not (Test-Path -LiteralPath $files.Prefetch) -and $count.ValueCount -eq 0 -and $app.ValueCount -eq 0 -and $item.ValueCount -eq 0) 'Independent recent-file, recent-program and Prefetch cleanup still completes'
    Assert ($mui.ValueCount -eq 1 -and $mui.Values['LangID'] -eq 'language' -and $apps.Values['TrackingPreference'] -eq 1) 'MUI language and Windows tracking preferences remain unchanged'
    Assert ((& $module {$script:closeCount}) -eq 0) 'Close-apps opt-out is respected'
    Assert ($bagMru.ValueCount -eq 0 -and $bagChild.ValueCount -eq 0 -and $bagView.ValueCount -eq 0 -and $defaults.Values['KeepDefaults'] -eq 1) 'Folder browsing traces are cleared without changing global view defaults'
    Assert ($bam.ValueCount -eq 2 -and $dam.ValueCount -eq 2 -and $bam.Values['Version'] -eq 1 -and $dam.Values['SequenceNumber'] -eq 9) 'BAM/DAM execution records clear while service bookkeeping values remain'
    Assert (-not $shim.Values.ContainsKey('AppCompatCache') -and $shim.Values.ContainsKey('CacheMainSdb') -and $shim.Values.ContainsKey('SdbTime')) 'Only the compatibility history value is removed, not its database configuration'
    Assert ($feature.ValueCount -eq 0 -and $mpc.ValueCount -eq 0) 'Taskbar application usage and supported application MRU records clear'
    Assert (@(Get-ChildItem -LiteralPath $pcaFolder -Filter '*.txt').Count -eq 0 -and (Test-Path -LiteralPath (Join-Path $pcaFolder 'keep.dll'))) 'PCA cleanup only removes the three named history logs'
    Assert (-not (& $module {param($path) Test-UnjunkPrivacyKey ($path.Replace((Get-UnjunkPrivacyUserSid),'S-1-5-18'))} $bamPath)) 'Background app history cleanup cannot cross into another user SID'
    $result=Invoke-UnjunkPrivacyCleanup -Exclusions @($files.Excluded) -CloseApps -Confirm:$false
    Assert ((& $module {$script:closeCount}) -eq 1 -and $result.FileResult.ClosedProcesses -eq 1) 'Enabled app closing stops IDM once before clearing its history'
    Assert (-not (Test-Path -LiteralPath $files.Url) -and -not (Test-Path -LiteralPath $files.Folder) -and -not $downloads.Children.ContainsKey('100')) 'IDM histories saved during app exit and completed records are removed'
    Assert ($result.ItemsCleaned -eq 3 -and $result.FileResult.DeletedFiles -eq 2 -and $result.RemovedPrivacyRecords -eq 1) 'IDM record and file counts are accurate without double counting'
    Assert ($downloads.Children.ContainsKey('101') -and $downloads.Children.ContainsKey('102') -and $downloads.Children.ContainsKey('103')) 'Unfinished, mismatched-size and unsaved download jobs are preserved'
    Assert ($downloads.Values['Registration'] -eq 'keep-license' -and $downloads.Values['LaunchOnStart'] -eq 1 -and $downloads.Children.ContainsKey('Settings')) 'IDM settings, registration and non-download registry branches remain intact'
    Assert (Test-Path -LiteralPath $files.Excluded) 'Prefetch path exclusions are honored'
    foreach($name in $hashes.Keys){Assert ((Get-FileHash -LiteralPath $files[$name]).Hash -eq $hashes[$name]) ('Preserved byte-for-byte: '+$name)}
    Assert (($result|ConvertTo-Json -Depth 8) -notmatch 'private-|credential|Cookie|Hash') 'Cleanup report stores no cleared file names, URLs, credentials or hashes'
    Assert ($result.PrivacyBreakdown.Count -eq 10 -and $result.Output -match 'IDM history:' -and $result.Output -match 'Prefetch:' -and $result.Output -match 'Recent programs:') 'Results explicitly show all privacy groups'
    $downloads.Children['200']=New-Download
    $entry=@(& $module {Get-UnjunkIdmHistoryEntries})[0]
    $downloads.Children['200'].Children['ChList'].SetValue('Cookie','changed-private-cookie','String')
    Assert (-not (& $module {param($e) Remove-UnjunkIdmHistoryEntry $e} $entry) -and $downloads.Children.ContainsKey('200')) 'Changed nested IDM history is preserved after scanning'
    $entry=@(& $module {Get-UnjunkIdmHistoryEntries})[0];& $module {$script:idmRunning=$true}
    Assert (-not (& $module {param($e) Remove-UnjunkIdmHistoryEntry $e} $entry)) 'An IDM restart blocks registry deletion'
    $rejected=$false;try{& $module {Remove-UnjunkIdmHistoryEntry ([pscustomobject]@{Key='..\Settings';Hash='x'})}}catch{$rejected=$true}
    Assert $rejected 'Forged IDM paths cannot delete settings or parent branches'
    & $module {$script:stayRunning=$true}
    $stuck=Invoke-UnjunkPrivacyCleanup -CloseApps -Exclusions @($files.Excluded) -Confirm:$false
    Assert ($stuck.SkippedItems -gt 0 -and $downloads.Children.ContainsKey('200')) 'Failed app closure leaves IDM history intact and reports skips'
    & $module {$script:idmRunning=$false;$script:stayRunning=$false}
    $pca.Denied=$true
    $partialPreview=Get-UnjunkPrivacyPreview -Exclusions @($files.Excluded)
    Assert ('Recent programs' -in $partialPreview.Warnings -and $partialPreview.IdmDownloads -eq 1) 'An inaccessible registry branch does not suppress other privacy categories'
    $partial=Invoke-UnjunkPrivacyCleanup -Exclusions @($files.Excluded) -Confirm:$false
    Assert ($partial.SkippedItems -gt 0 -and -not $downloads.Children.ContainsKey('200')) 'Partial inventory produces a partial result while available history is cleaned'
    $pca.Denied=$false
    $registry[$base+'\RecentDocs'].SetValue('0','before','String')
    $entry=@(& $module {Get-UnjunkPrivacyEntries} | Where-Object Name -eq '0')[0]
    $registry[$base+'\RecentDocs'].SetValue('0','after','String')
    Assert (-not (& $module {param($e) Remove-UnjunkPrivacyEntry $e} $entry)) 'Changed Windows registry values survive the execution recheck'
    & $module {
        function script:Get-UnjunkActivityInventory {param($Exclusions) [pscustomobject]@{Databases=@([pscustomobject]@{Path='fixture.db';Count=4});Issues=@()}}
        function script:Clear-UnjunkActivityDatabase {param($Path,$Exclusions) [pscustomobject]@{RemovedRecords=4;JournalCleared=$false}}
    }
    $activityResult=Invoke-UnjunkPrivacyCleanup -Exclusions @($files.Excluded) -Confirm:$false
    $activityGroup=@($activityResult.PrivacyBreakdown | Where-Object Name -eq 'Windows activity history')[0]
    Assert ($activityGroup.RemovedRecords -eq 4 -and $activityGroup.Issues.Count -eq 1 -and $activityResult.SkippedItems -gt 0) 'Activity rows remain counted when a busy journal produces a partial privacy result'
} finally {
    $env:APPDATA=$previousAppData
    $path=[IO.Path]::GetFullPath($fixture);$parent=[IO.Path]::GetFullPath($PSScriptRoot)+'\'
    if(-not $path.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $path -Leaf) -notlike 'PrivacyExtendedFixture-*'){throw 'Unsafe fixture path'}
    Remove-Item -LiteralPath $path -Recurse -Force
}
Write-Host "All $passed extended privacy checks passed."
