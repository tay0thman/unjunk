#requires -Version 5.1
# In-memory Shell and disposable filesystem fixtures. Never calls the real Explorer or removes real history.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.ExplorerHistory.psm1') -Force
$module=Get-Module Unjunk.ExplorerHistory
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
public class UnjunkExplorerFixtureItem {
    public static List<UnjunkExplorerFixtureItem> Store=new List<UnjunkExplorerFixtureItem>();
    public static int Invocations=0,LegacyInvocations=0;
    public static int VisibleLimit=Int32.MaxValue,Refills=0;
    public string Path;
    bool isFolder=false;
    public bool IsFolder {get{return isFolder;} set{isFolder=value;FileAttributes=value?16:128;}}
    public bool IsFileSystem=true,Pinned=false,NoRemove=false,Throw=false,ChangeAction=false;
    public bool MetadataUnavailable=false,UseRealMetadata=false,Hidden=false;
    public int FileAttributes=128;
    public int VerbReads=0;
    public string Action="Re&move from list";
    public UnjunkExplorerFixtureItem(string path){Path=path;}
    public void ClearHistory(){
        Invocations++;
        if(Throw)throw new Exception("private-file-error");
        if(!NoRemove){
            Store.Remove(this);
            if(Refills>0){Store.Add(new UnjunkExplorerFixtureItem("C:\\ExplorerFixture\\refill-"+Refills+".txt"));Refills--;}
        }
    }
    public void OpenAgain(){if(!Hidden&&!Store.Contains(this))Store.Add(this);}
    public object ExtendedProperty(string name){return Pinned;}
    public UnjunkExplorerFixtureVerbs Verbs(){
        VerbReads++;
        return new UnjunkExplorerFixtureVerbs(this,new[]{"Open","Delete",Pinned?"Remove from Favorites":"Add to Favorites",ChangeAction&&VerbReads>1?"Delete":Action});
    }
}
public class UnjunkExplorerFixtureVerb {
    public string Name;
    UnjunkExplorerFixtureItem owner;
    public UnjunkExplorerFixtureVerb(UnjunkExplorerFixtureItem item,string name){owner=item;Name=name;}
    public void DoIt(){
        if(Name!="Re&move from list")throw new Exception("Unsafe action invoked: "+Name);
        UnjunkExplorerFixtureItem.LegacyInvocations++;
        owner.Hidden=true;owner.ClearHistory();
    }
}
public class UnjunkExplorerFixtureVerbs {
    string[] names;UnjunkExplorerFixtureItem owner;
    public UnjunkExplorerFixtureVerbs(UnjunkExplorerFixtureItem item,string[] actions){owner=item;names=actions;}
    public int Count {get{return names.Length;}}
    public UnjunkExplorerFixtureVerb Item(int index){return new UnjunkExplorerFixtureVerb(owner,names[index]);}
}
public class UnjunkExplorerFixtureItems {
    public int Count {get{return Math.Min(UnjunkExplorerFixtureItem.VisibleLimit,UnjunkExplorerFixtureItem.Store.Count);}}
    public UnjunkExplorerFixtureItem Item(int index){return UnjunkExplorerFixtureItem.Store[index];}
}
'@
& $module {
    $script:sessions=0;$script:closed=0
    $script:nativeFileAttributes=${function:Get-UnjunkExplorerFileAttributes}
    function script:New-UnjunkExplorerSession {$script:sessions++;[pscustomobject]@{Shell=$null;Folder=$null;Items=(New-Object UnjunkExplorerFixtureItems)}}
    function script:Close-UnjunkExplorerSession {param($Session) if($null -ne $Session){$script:closed++}}
    function script:Remove-UnjunkExplorerRecentFile {
        param($Path)
        $items=@([UnjunkExplorerFixtureItem]::Store | Where-Object Path -eq $Path)
        if($items.Count -ne 1){throw 'Fixture history identity unavailable'}
        $items[0].ClearHistory()
    }
    function script:Get-UnjunkExplorerFileAttributes {
        param($Path)
        $items=@([UnjunkExplorerFixtureItem]::Store | Where-Object Path -eq $Path)
        if($items.Count -ne 1 -or $items[0].MetadataUnavailable){throw 'Fixture metadata unavailable'}
        if($items[0].UseRealMetadata){return (& $script:nativeFileAttributes $Path)}
        return [IO.FileAttributes]$items[0].FileAttributes
    }
}
$passed=0
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
function Item([string]$Name){$item=New-Object UnjunkExplorerFixtureItem ('C:\ExplorerFixture\'+$Name);[UnjunkExplorerFixtureItem]::Store.Add($item);return $item}
function Reset {[UnjunkExplorerFixtureItem]::Store.Clear();[UnjunkExplorerFixtureItem]::Invocations=0;[UnjunkExplorerFixtureItem]::LegacyInvocations=0;[UnjunkExplorerFixtureItem]::VisibleLimit=[int]::MaxValue;[UnjunkExplorerFixtureItem]::Refills=0}
Reset
$a=Item 'private-a.txt';$b=Item 'private-b.txt';$folder=Item 'pinned-folder';$folder.IsFolder=$true
$favorite=Item 'favorite.txt';$favorite.Pinned=$true;$excluded=Item 'keep.txt';$unknown=Item 'unsupported.txt';$unknown.Action='Delete'
$scan=Get-UnjunkExplorerHistory -Exclusions @($excluded.Path)
Assert ($scan.RecentFiles -eq 2 -and $scan.FoldersKept -eq 1 -and $scan.FavoritesKept -eq 1 -and $scan.Excluded -eq 1 -and $scan.Unsupported -eq 1) 'Inventory separates removable recent files from folders, favorites, exclusions and unsupported items'
Assert (($scan|ConvertTo-Json -Depth 6) -notmatch 'private-a|private-b|favorite.txt') 'Preview uses temporary identities without exposing recent filenames'
$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -WhatIf
Assert ([UnjunkExplorerFixtureItem]::Invocations -eq 0 -and [UnjunkExplorerFixtureItem]::Store.Count -eq 6 -and $result.ItemsCleaned -eq 0) 'WhatIf does not invoke any Shell action'
$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -State @{Cancel=$true;Message='';Progress=0} -Confirm:$false
Assert ([UnjunkExplorerFixtureItem]::Invocations -eq 0) 'Cancellation does not invoke any Shell action'
$later=Item 'new-after-scan.txt';$a.Pinned=$true
$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($result.ItemsCleaned -eq 2 -and $result.RemovedPrivacyRecords -eq 2 -and $null -eq $result.FileResult) 'Confirmed Explorer list removals count as history, not deleted files'
Assert ([UnjunkExplorerFixtureItem]::Store.Contains($a) -and -not [UnjunkExplorerFixtureItem]::Store.Contains($later)) 'New pins survive while the current unpinned recent list is cleared'
Assert ([UnjunkExplorerFixtureItem]::Store.Contains($folder) -and [UnjunkExplorerFixtureItem]::Store.Contains($favorite) -and [UnjunkExplorerFixtureItem]::Store.Contains($excluded)) 'All folders, pinned favorites and exclusions remain'
Assert ($result.SkippedItems -eq 1 -and [UnjunkExplorerFixtureItem]::Invocations -eq 2) 'Unsupported items are reported without invoking Delete or Unpin'
Assert (($result|ConvertTo-Json -Depth 6) -notmatch 'private-|favorite.txt|[0-9A-F]{64}') 'Cleanup reports contain counts without identities or filenames'
Assert ((& $module {$script:sessions}) -eq (& $module {$script:closed})) 'Every Shell session is released'
Reset;$first=Item 'first.txt';$second=Item 'second.txt';$third=Item 'third.txt'
$scan=Get-UnjunkExplorerHistory;$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($result.ItemsCleaned -eq 3 -and [UnjunkExplorerFixtureItem]::Store.Count -eq 0) 'Reverse traversal does not skip adjacent items when Explorer removes synchronously'
Reset;$folder=Item 'folder';$folder.IsFolder=$true;$favorite=Item 'favorite';$favorite.Pinned=$true
for($i=0;$i -lt 6;$i++){$null=Item ('older-'+$i+'.txt')}
[UnjunkExplorerFixtureItem]::VisibleLimit=4
$scan=Get-UnjunkExplorerHistory;$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($scan.RecentFiles -eq 2 -and $result.ItemsCleaned -eq 6 -and $result.Batches -eq 3 -and $result.RemainingRecentFiles -eq 0) 'Older entries revealed behind the visible batch are also cleared'
Assert ([UnjunkExplorerFixtureItem]::Store.Contains($folder) -and [UnjunkExplorerFixtureItem]::Store.Contains($favorite)) 'All batches preserve folders and pinned favorites'
Reset;$null=Item 'refilling.txt';[UnjunkExplorerFixtureItem]::Refills=200
$scan=Get-UnjunkExplorerHistory;$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($result.Batches -eq 100 -and $result.RemainingRecentFiles -eq 1 -and $result.SkippedItems -eq 1) 'Continuously recreated history stops at a bound and reports the remaining entry'
Reset;$item=Item 'late-excluded.txt';$scan=Get-UnjunkExplorerHistory;$scan.Exclusions=@($item.Path)
$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($result.ItemsCleaned -eq 0 -and [UnjunkExplorerFixtureItem]::Invocations -eq 0) 'Execution checks exclusions again'
Reset;$item=Item 'changed-action.txt';$scan=Get-UnjunkExplorerHistory;$item.VerbReads=0;$item.ChangeAction=$true
$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($result.ItemsCleaned -eq 0 -and [UnjunkExplorerFixtureItem]::Invocations -eq 0 -and $result.SkippedItems -eq 1) 'A changed action index cannot become a Delete command'
Reset;$item=Item 'no-op.txt';$item.NoRemove=$true;$scan=Get-UnjunkExplorerHistory
$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($result.ItemsCleaned -eq 0 -and $result.SkippedItems -eq 1 -and [UnjunkExplorerFixtureItem]::Store.Contains($item)) 'An accepted but ineffective Shell action is not reported as cleaned'
Reset;$item=Item 'throw.txt';$item.Throw=$true;$scan=Get-UnjunkExplorerHistory
$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($result.ItemsCleaned -eq 0 -and $result.SkippedItems -eq 1 -and $result.Output -notmatch 'private-file-error') 'Shell failures are reported without logging private file details'
$scan.ScannedAt=[DateTime]::UtcNow.AddHours(-1).ToString('o');$rejected=$false
try{Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false}catch{$rejected=$true}
Assert $rejected 'Expired Explorer scans are rejected'
$scan=Get-UnjunkExplorerHistory;$scan.UserSid='S-1-5-18';$rejected=$false
try{Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false}catch{$rejected=$true}
Assert $rejected 'Cross-account Explorer scans are rejected'
$scan=Get-UnjunkExplorerHistory;$scan.Entries[0].Id='not-a-valid-identity';$rejected=$false
try{Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false}catch{$rejected=$true}
Assert $rejected 'Malformed history identities are rejected before Shell execution'
Assert (-not (& $module {Test-UnjunkExplorerRemoveAction ''}) -and -not (& $module {Test-UnjunkExplorerRemoveAction 'Delete'}) -and -not (& $module {Test-UnjunkExplorerRemoveAction 'Unpin from Quick access'})) 'Empty, file-deletion and unpin verbs are never eligible'
Reset
$fixtureParent=[IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\')
$fixtureRoot=Join-Path $fixtureParent ('ExplorerArchiveFixture-'+[Guid]::NewGuid().ToString('N'))
try {
    $null=New-Item -ItemType Directory -Path $fixtureRoot
    $source=Join-Path $fixtureRoot 'payload';$null=New-Item -ItemType Directory -Path $source
    [IO.File]::WriteAllText((Join-Path $source 'keep.txt'),'Archive contents must survive history removal.')
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archivePath=Join-Path $fixtureRoot 'recent.zip'
    [IO.Compression.ZipFile]::CreateFromDirectory($source,$archivePath)
    $originalHash=(Get-FileHash -LiteralPath $archivePath).Hash
    $directoryPath=Join-Path $fixtureRoot 'actual-directory.zip';$null=New-Item -ItemType Directory -Path $directoryPath
    $archive=Item 'archive';$archive.Path=$archivePath;$archive.IsFolder=$true;$archive.UseRealMetadata=$true
    $directory=Item 'directory';$directory.Path=$directoryPath;$directory.IsFolder=$true;$directory.UseRealMetadata=$true
    $virtual=Item 'virtual';$virtual.IsFolder=$true;$virtual.IsFileSystem=$false;$virtual.MetadataUnavailable=$true
    $pinnedArchive=Item 'pinned.zip';$pinnedArchive.IsFolder=$true;$pinnedArchive.FileAttributes=32;$pinnedArchive.Pinned=$true
    $excludedArchive=Item 'excluded.zip';$excludedArchive.IsFolder=$true;$excludedArchive.FileAttributes=32
    $otherArchive=Item 'other-container.7z';$otherArchive.IsFolder=$true;$otherArchive.FileAttributes=32
    $scan=Get-UnjunkExplorerHistory -Exclusions @($excludedArchive.Path)
    Assert ($scan.RecentFiles -eq 2 -and $scan.FoldersKept -eq 2 -and $scan.Issues -eq 0) 'Shell containers use physical directory attributes: archives are recent files while real and virtual folders stay protected'
    Assert ($scan.FavoritesKept -eq 1 -and $scan.Excluded -eq 1) 'Pinned and excluded archives remain protected'
    $result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($result.ItemsCleaned -eq 2 -and $result.RemainingRecentFiles -eq 0 -and $result.SkippedItems -eq 0) 'Archive usage history is cleared regardless of filename extension'
    Assert ((Get-FileHash -LiteralPath $archivePath).Hash -eq $originalHash) 'Removing archive history leaves the original archive bytes intact'
    Assert ([UnjunkExplorerFixtureItem]::Store.Contains($directory) -and [UnjunkExplorerFixtureItem]::Store.Contains($virtual) -and [UnjunkExplorerFixtureItem]::Store.Contains($pinnedArchive) -and [UnjunkExplorerFixtureItem]::Store.Contains($excludedArchive)) 'Directory names ending in .zip, virtual folders, pins and exclusions survive archive cleanup'
} finally {
    Reset
    $resolvedFixture=[IO.Path]::GetFullPath($fixtureRoot)
    if(-not $resolvedFixture.StartsWith($fixtureParent+'\',[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolvedFixture -Leaf) -notmatch '^ExplorerArchiveFixture-[0-9a-f]{32}$'){throw 'Invalid fixture cleanup path'}
    if(Test-Path -LiteralPath $resolvedFixture){Remove-Item -LiteralPath $resolvedFixture -Recurse -Force}
}
$archive=Item 'unreadable.zip';$archive.IsFolder=$true;$archive.MetadataUnavailable=$true
$scan=Get-UnjunkExplorerHistory;$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($scan.Issues -eq 1 -and $scan.FoldersKept -eq 0 -and $result.SkippedItems -eq 1 -and [UnjunkExplorerFixtureItem]::Invocations -eq 0) 'Unreadable container metadata is reported as incomplete inventory rather than a successful empty cleanup'
Reset;$archive=Item 'unreadable-pinned-container';$archive.IsFolder=$true;$archive.MetadataUnavailable=$true;$archive.Pinned=$true
$scan=Get-UnjunkExplorerHistory;$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($scan.Issues -eq 0 -and $scan.FavoritesKept -eq 1 -and $result.SkippedItems -eq 0 -and [UnjunkExplorerFixtureItem]::Invocations -eq 0) 'A confirmed pin remains protected even when its target metadata is inaccessible'
Reset;$archive=Item 'changed-to-directory.zip';$archive.IsFolder=$true;$archive.FileAttributes=32
$scan=Get-UnjunkExplorerHistory;$archive.FileAttributes=16
$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($scan.RecentFiles -eq 1 -and $result.ItemsCleaned -eq 0 -and [UnjunkExplorerFixtureItem]::Invocations -eq 0) 'An archive replaced by a directory after scanning is protected at execution'
Reset;$archive=Item 'newly-pinned.zip';$archive.IsFolder=$true;$archive.FileAttributes=32
$scan=Get-UnjunkExplorerHistory;$archive.Pinned=$true
$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($scan.RecentFiles -eq 1 -and $result.ItemsCleaned -eq 0 -and [UnjunkExplorerFixtureItem]::Invocations -eq 0) 'An archive pinned after scanning is protected at execution'
Reset
Assert ((& $module {$script:sessions}) -eq (& $module {$script:closed})) 'Archive and incomplete-inventory paths release every Shell session'
Reset;$document=Item 'reopened.pdf';$archive=Item 'reopened.zip';$archive.IsFolder=$true;$archive.FileAttributes=32
$scan=Get-UnjunkExplorerHistory;$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($result.ItemsCleaned -eq 2 -and [UnjunkExplorerFixtureItem]::LegacyInvocations -eq 0) 'Cleanup uses usage removal and never invokes the persistent-hide context-menu action'
$document.OpenAgain();$archive.OpenAgain();$scan=Get-UnjunkExplorerHistory
Assert ($scan.RecentFiles -eq 2 -and -not $document.Hidden -and -not $archive.Hidden) 'Reopening a cleaned document or archive records it again instead of suppressing it'
$result=Invoke-UnjunkExplorerHistoryCleanup -Scan $scan -Confirm:$false
Assert ($result.ItemsCleaned -eq 2 -and $result.RemainingRecentFiles -eq 0) 'History created by later opens can be cleared by the next cleanup'
Reset
Write-Host "All $passed Explorer history checks passed."
