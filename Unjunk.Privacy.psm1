#requires -Version 5.1
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Core.psm1')
Import-Module (Join-Path $PSScriptRoot 'Unjunk.ActivityHistory.psm1')
$script:effectivePrivacyRoots=$null
$script:privacyKeyChecks=@{}
$script:privacyRoots=@('Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs',
    'Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU','Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths',
    'Software\Microsoft\Windows\CurrentVersion\Explorer\WordWheelQuery',
    'Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSavePidlMRU',
    'Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\LastVisitedPidlMRU',
    'Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist',
    'Software\Microsoft\Windows\CurrentVersion\Search\RecentApps',
    'Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache',
    'Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Store',
    'Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Persisted',
    'Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\BagMRU',
    'Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags',
    'Software\Microsoft\Windows\Shell\BagMRU','Software\Microsoft\Windows\Shell\Bags',
    'Software\Microsoft\Windows\ShellNoRoam\BagMRU','Software\Microsoft\Windows\ShellNoRoam\Bags',
    'Software\Microsoft\Windows\CurrentVersion\Explorer\FeatureUsage\AppBadgeUpdated',
    'Software\Microsoft\Windows\CurrentVersion\Explorer\FeatureUsage\AppLaunch',
    'Software\Microsoft\Windows\CurrentVersion\Explorer\FeatureUsage\AppSwitched',
    'Software\Microsoft\Windows\CurrentVersion\Explorer\FeatureUsage\ShowJumpView',
    'Software\Microsoft\Windows\CurrentVersion\Explorer\FeatureUsage\TrayButtonClicked',
    'Software\MPC-HC\MPC-HC\Recent File List','Software\MPC-HC\MPC-HC\MediaHistory',
    'Software\MPC-BE\Recent File List',
    'Software\Microsoft\Windows\CurrentVersion\Applets\Paint\Recent File List',
    'Software\Microsoft\Windows\CurrentVersion\Applets\Wordpad\Recent File List')
$script:idmRegistryRoot='Software\DownloadManager'

function Open-UnjunkPrivacyRegistryKey([string]$Path,[bool]$Writable=$false) {
    if($Path.StartsWith('HKLM\',[StringComparison]::OrdinalIgnoreCase)){
        return [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($Path.Substring(5),$Writable)
    }
    [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($Path,$Writable)
}

function Get-UnjunkPrivacyUserSid {[Security.Principal.WindowsIdentity]::GetCurrent().User.Value}

function Get-UnjunkPrivacyRoots {
    if($null -ne $script:effectivePrivacyRoots){return $script:effectivePrivacyRoots}
    $sid=Get-UnjunkPrivacyUserSid
    $script:effectivePrivacyRoots=@($script:privacyRoots)+@(foreach($service in @('bam','dam')){foreach($branch in @('UserSettings','State\UserSettings')){
        'HKLM\SYSTEM\CurrentControlSet\Services\'+$service+'\'+$branch+'\'+$sid
    }})+@('HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\AppCompatCache')
    return $script:effectivePrivacyRoots
}

function Test-UnjunkPrivacyKey([string]$Path) {
    if(-not $script:privacyKeyChecks.ContainsKey($Path)){$script:privacyKeyChecks[$Path]=[bool](Test-UnjunkPrivacyKeyUncached $Path)}
    return $script:privacyKeyChecks[$Path]
}

function Test-UnjunkPrivacyKeyUncached([string]$Path) {
    if($Path -match '\\\.\.?($|\\)'){return $false}
    foreach($root in @(Get-UnjunkPrivacyRoots)){
        if($root.EndsWith('\UserAssist')){
            if($Path -match ('^'+[regex]::Escape($root)+'\\\{[0-9a-fA-F-]{36}\}\\Count$')){return $true}
        }elseif($root.EndsWith('\RecentApps')){
            if($Path -match ('^'+[regex]::Escape($root)+'\\\{[0-9a-fA-F-]{36}\}(\\RecentItems\\\{[0-9a-fA-F-]{36}\})?$')){return $true}
        }elseif($root -match '\\(Bags|BagMRU)$'){
            # Preserve global folder defaults under Bags\AllFolders; clear numbered history only.
            if($Path -eq $root -or ($Path.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase) -and $Path.Substring($root.Length) -match '^\\\d+(\\|$)')){return $true}
        }elseif($Path -eq $root -or (($root -match 'RecentDocs|PidlMRU$|\\(MediaHistory|Recent File List)$') -and $Path.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase))){return $true}
    }
    return $false
}

function Test-UnjunkPrivacyValue([string]$Path,[string]$Name) {
    if(-not $Name -or -not (Test-UnjunkPrivacyKey $Path)){return $false}
    if($Path -match '^HKLM\\SYSTEM\\CurrentControlSet\\Services\\(bam|dam)\\'){
        return $Name -match '^(\\Device\\|[a-zA-Z]:\\).+\.exe$'
    }
    if($Path -match '^HKLM\\.+\\AppCompatCache$'){return $Name -eq 'AppCompatCache'}
    # Preserve the language preference; only application name/company cache values are history.
    if($Path.EndsWith('\MuiCache',[StringComparison]::OrdinalIgnoreCase)){return $Name -match '\.(FriendlyAppName|ApplicationCompany)$'}
    return $true
}

function Get-UnjunkPrivacyGroup([string]$Path) {
    if($Path -match '\\(Bags|BagMRU)(\\|$)'){return 'Folder browsing history'}
    if($Path -match '^HKLM\\'){return 'Background and compatibility history'}
    if($Path -match '\\(MPC-HC|MPC-BE|Applets)\\'){return 'Application file history'}
    if($Path -match '\\(UserAssist|RecentApps|MuiCache|Compatibility Assistant|FeatureUsage)(\\|$)'){return 'Recent programs'}
    return 'Recent file history'
}

function Get-UnjunkPrivacyHash($Value,$Kind) {
    if($Value -is [byte[]]){
        $hash=[Security.Cryptography.SHA256]::Create()
        try{
            $prefix=[Text.Encoding]::UTF8.GetBytes([string]$Kind+'|')
            $null=$hash.TransformBlock($prefix,0,$prefix.Length,$prefix,0)
            $null=$hash.TransformFinalBlock($Value,0,$Value.Length)
            return [Convert]::ToBase64String($hash.Hash)
        }finally{$hash.Dispose()}
    }
    $bytes=[Text.Encoding]::UTF8.GetBytes(([string]$Kind+'|'+(ConvertTo-Json -InputObject $Value -Compress -Depth 8)))
    $hash=[Security.Cryptography.SHA256]::Create()
    try{[Convert]::ToBase64String($hash.ComputeHash($bytes))}finally{$hash.Dispose()}
}

function Set-UnjunkPrivacyProgress($State,[string]$Message) {
    if($null -eq $State){return}
    $State.Message=$Message
    if($State.ContainsKey('StatusPath') -and (-not $State.ContainsKey('PrivacyHeartbeat') -or ([DateTime]::UtcNow-$State.PrivacyHeartbeat).TotalSeconds -ge 2)){
        $State.PrivacyHeartbeat=[DateTime]::UtcNow
        try{@{Finished=$false;UpdatedAt=[DateTime]::UtcNow.ToString('o');Message=$Message}|ConvertTo-Json|Set-Content -LiteralPath $State.StatusPath -Encoding UTF8 -ErrorAction Stop}catch{}
    }
}

function Get-UnjunkPrivacyEntries {
    param($Issues,$State)
    $count=0
    foreach($root in @(Get-UnjunkPrivacyRoots)){
        if($null -ne $State -and $State.Cancel){break}
        Set-UnjunkPrivacyProgress $State ('Scanning privacy: '+(Get-UnjunkPrivacyGroup $root))
        $stack=New-Object 'System.Collections.Generic.Stack[string]';$stack.Push($root)
        while($stack.Count){
            if($null -ne $State -and $State.Cancel){break}
            $path=$stack.Pop();$key=$null
            try {
                $key=Open-UnjunkPrivacyRegistryKey $path
                if($null -eq $key){continue}
                if(Test-UnjunkPrivacyKey $path){
                    foreach($name in $key.GetValueNames()){
                        if(-not (Test-UnjunkPrivacyValue $path $name)){continue}
                        $value=$key.GetValue($name,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                        [pscustomobject]@{Key=$path;Name=$name;Hash=(Get-UnjunkPrivacyHash $value $key.GetValueKind($name))}
                        $count++;if($count%250 -eq 0){Set-UnjunkPrivacyProgress $State ("Scanning privacy: $count registry records counted")}
                    }
                }
                foreach($child in $key.GetSubKeyNames()){
                    $next=$path+'\'+$child
                    $branch=($path -eq $root -and $root.EndsWith('\UserAssist') -and $child -match '^\{[0-9a-fA-F-]{36}\}$') -or
                        ($root.EndsWith('\RecentApps') -and (Test-UnjunkPrivacyKey $path) -and $child -eq 'RecentItems')
                    if((Test-UnjunkPrivacyKey $next) -or $branch){$stack.Push($next)}
                }
            } catch {
                if($null -eq $Issues){throw}
                $Issues.Add((Get-UnjunkPrivacyGroup $root))
            } finally {if($null -ne $key){$key.Dispose()}}
        }
    }
}

function Remove-UnjunkPrivacyEntry($Entry) {
    if(-not (Test-UnjunkPrivacyValue $Entry.Key $Entry.Name)){throw 'Registry location is outside the privacy cleanup allowlist.'}
    $key=$null
    try {
        $key=Open-UnjunkPrivacyRegistryKey $Entry.Key $true
        if($null -eq $key -or $Entry.Name -notin $key.GetValueNames()){return $false}
        $value=$key.GetValue($Entry.Name,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        if((Get-UnjunkPrivacyHash $value $key.GetValueKind($Entry.Name)) -ne $Entry.Hash){return $false}
        $key.DeleteValue($Entry.Name,$false)
        return $true
    } finally {if($null -ne $key){$key.Dispose()}}
}

function Get-UnjunkIdmDataRoot {
    $key=$null
    try {
        $key=Open-UnjunkPrivacyRegistryKey $script:idmRegistryRoot
        if($null -ne $key){
            $path=$key.GetValue('AppDataIDMFolder',$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            if($path -is [string] -and $path){return $path}
        }
    } finally {if($null -ne $key){$key.Dispose()}}
    Join-Path $env:APPDATA 'IDM'
}

function ConvertTo-UnjunkIdmSize($Value) {
    if($Value -is [byte[]] -and $Value.Length -eq 8){return [decimal][BitConverter]::ToUInt64($Value,0)}
    if(($Value -is [long] -or $Value -is [int]) -and $Value -ge 0){return [decimal]$Value}
    return $null
}

function Test-UnjunkIdmCompleted($Key) {
    # Locally verified IDM schema: completed/saved status 3/1 plus equal 64-bit sizes.
    # Unknown schemas and unfinished transfers stay intact; no downloaded file is targeted.
    foreach($name in @('Status','bSaved','FileSize','Downloaded','Url0')){if($name -notin $Key.GetValueNames()){return $false}}
    if($Key.GetValueKind('Status') -ne [Microsoft.Win32.RegistryValueKind]::DWord -or
        $Key.GetValueKind('bSaved') -ne [Microsoft.Win32.RegistryValueKind]::DWord -or
        $Key.GetValue('Status') -ne 3 -or $Key.GetValue('bSaved') -ne 1){return $false}
    $size=ConvertTo-UnjunkIdmSize $Key.GetValue('FileSize');$downloaded=ConvertTo-UnjunkIdmSize $Key.GetValue('Downloaded')
    return $null -ne $size -and $null -ne $downloaded -and $size -eq $downloaded
}

function Get-UnjunkPrivacyTreeHash($Key,[int]$Depth=0) {
    if($Depth -gt 8 -or $Key.ValueCount -gt 10000 -or $Key.SubKeyCount -gt 1000){throw 'Unexpected history record shape.'}
    $parts=New-Object 'System.Collections.Generic.List[object]'
    foreach($name in @($Key.GetValueNames() | Sort-Object)){
        $value=$Key.GetValue($name,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        $parts.Add(@($name,(Get-UnjunkPrivacyHash $value $Key.GetValueKind($name))))
    }
    foreach($name in @($Key.GetSubKeyNames() | Sort-Object)){
        $child=$null
        try{$child=$Key.OpenSubKey($name,$false);if($null -eq $child){throw 'History changed during scan.'};$parts.Add(@($name,(Get-UnjunkPrivacyTreeHash $child ($Depth+1))))}
        finally{if($null -ne $child){$child.Dispose()}}
    }
    Get-UnjunkPrivacyHash $parts.ToArray() 'RegistryTree'
}

function Get-UnjunkIdmHistoryEntries {
    param($Issues)
    $root=$null
    try {
        $root=Open-UnjunkPrivacyRegistryKey $script:idmRegistryRoot
        if($null -eq $root){return}
        foreach($name in $root.GetSubKeyNames()){
            if($name -notmatch '^\d{1,10}$'){continue}
            $key=$null
            try {
                $key=$root.OpenSubKey($name,$false)
                if($null -ne $key -and (Test-UnjunkIdmCompleted $key)){
                    [pscustomobject]@{Key=$name;Hash=(Get-UnjunkPrivacyTreeHash $key)}
                }
            } catch {if($null -eq $Issues){throw};$Issues.Add('IDM history')}
            finally{if($null -ne $key){$key.Dispose()}}
        }
    } catch {if($null -eq $Issues){throw};$Issues.Add('IDM history')}
    finally{if($null -ne $root){$root.Dispose()}}
}

function Test-UnjunkIdmRunning {
    $processes=@(Get-Process -Name IDMan -ErrorAction SilentlyContinue)
    try{return $processes.Count -gt 0}finally{foreach($process in $processes){$process.Dispose()}}
}

function Close-UnjunkIdm($State) {
    $core=Get-Module Unjunk.Core
    & $core {param($state)
        # Reuse the bounded interactive-process closer, with an exact IDM process mapping.
        $target=[pscustomobject]@{Name='IDM history';Risk='Cache';Processes=@('IDMan')}
        Close-UnjunkTargetApps -Target $target -State $state
    } $State
}

function Remove-UnjunkIdmHistoryEntry($Entry) {
    if($Entry.Key -notmatch '^\d{1,10}$'){throw 'Not an IDM download record.'}
    if(Test-UnjunkIdmRunning){return $false}
    $root=$null;$key=$null
    try {
        $root=Open-UnjunkPrivacyRegistryKey $script:idmRegistryRoot $true
        if($null -eq $root){return $false}
        $key=$root.OpenSubKey($Entry.Key,$false)
        if($null -eq $key -or -not (Test-UnjunkIdmCompleted $key) -or (Get-UnjunkPrivacyTreeHash $key) -ne $Entry.Hash){return $false}
        $key.Dispose();$key=$null
        if(Test-UnjunkIdmRunning){return $false}
        $root.DeleteSubKeyTree($Entry.Key,$false)
        return $true
    } finally {if($null -ne $key){$key.Dispose()};if($null -ne $root){$root.Dispose()}}
}

function Get-UnjunkPrivacyFileScan {
    param([string[]]$Exclusions=@(),$State)
    $core=Get-Module Unjunk.Core
    & $core {param($exclusions,$state)
        $root=Join-Path $env:APPDATA 'Microsoft\Windows\Recent'
        # Nonrecursive matching never enters AutomaticDestinations or CustomDestinations.
        $target=New-UnjunkTarget 'Recent shortcuts' 'Privacy' $root -Risk Review -FilePattern '*.lnk' -Recurse $false
        Get-UnjunkTargetScan $target 0 $exclusions $state $false
    } $Exclusions $State
}

function Get-UnjunkPrivacyFileScans {
    param([string[]]$Exclusions=@(),$State,[switch]$IdmOnly)
    $core=Get-Module Unjunk.Core
    if(-not $IdmOnly){
        try{Get-UnjunkPrivacyFileScan -Exclusions $Exclusions -State $State}catch{New-UnjunkPrivacyFailedScan 'Recent shortcuts'}
        try{& $core {param($exclusions,$state,$windowsRoot)
            $target=New-UnjunkTarget 'Prefetch' 'Privacy' (Join-Path $windowsRoot 'Prefetch') -Risk Review -FilePattern '*.pf' -Recurse $false
            Get-UnjunkTargetScan $target 0 $exclusions $state $false
        } $Exclusions $State (Get-UnjunkPrivacyWindowsRoot)}catch{New-UnjunkPrivacyFailedScan 'Prefetch'}
        foreach($name in @('PcaAppLaunchDic.txt','PcaGeneralDb0.txt','PcaGeneralDb1.txt')){
            try{& $core {param($exclusions,$state,$windowsRoot,$name)
                $target=New-UnjunkTarget 'Compatibility Assistant logs' 'Privacy' (Join-Path $windowsRoot ('appcompat\pca\'+$name)) -Risk Review -FilePattern $name -Recurse $false
                Get-UnjunkTargetScan $target 0 $exclusions $state $false
            } $Exclusions $State (Get-UnjunkPrivacyWindowsRoot) $name}catch{New-UnjunkPrivacyFailedScan 'Compatibility Assistant logs'}
        }
    }
    try{$idmRoot=Get-UnjunkIdmDataRoot}catch{New-UnjunkPrivacyFailedScan 'IDM history';return}
    foreach($name in @('UrlHistory.txt','foldresHistory.txt')){
        try{& $core {param($root,$name,$exclusions,$state)
            $target=New-UnjunkTarget 'IDM history' 'Privacy' (Join-Path $root $name) -Risk Review -FilePattern $name -Recurse $false
            Get-UnjunkTargetScan $target 0 $exclusions $state $false
        } $idmRoot $name $Exclusions $State}catch{New-UnjunkPrivacyFailedScan 'IDM history'}
    }
}

function Get-UnjunkPrivacyWindowsRoot {return $env:SystemRoot}

function New-UnjunkPrivacyFailedScan([string]$Name) {
    [pscustomobject]@{Name=$Name;Root='';Files=@();Issues=@('History location unavailable.');SkippedLinks=0}
}

function Get-UnjunkPrivacyInventory {
    param([string[]]$Exclusions=@(),$State)
    $issues=New-Object 'System.Collections.Generic.List[string]'
    $scans=@(Get-UnjunkPrivacyFileScans -Exclusions $Exclusions -State $State)
    $entries=@(Get-UnjunkPrivacyEntries -Issues $issues -State $State)
    $idm=@(Get-UnjunkIdmHistoryEntries -Issues $issues)
    $activity=Get-UnjunkActivityInventory -Exclusions $Exclusions
    if($activity.Issues.Count){$issues.Add('Windows activity history')}
    foreach($scan in $scans){if($scan.Issues.Count -or $scan.SkippedLinks){$issues.Add($scan.Name)}}
    [pscustomobject]@{Scans=$scans;Entries=$entries;IdmEntries=$idm;Activity=$activity;Issues=$issues.ToArray()}
}

function Get-UnjunkPrivacyPreview {
    param([string[]]$Exclusions=@(),$State)
    if($null -eq $State){$State=@{Cancel=$false;Progress=0;Message=''}}
    $core=Get-Module Unjunk.Core
    $Exclusions=@(& $core {param($paths) foreach($path in $paths){Get-UnjunkFullPath $path}} $Exclusions)
    $inventory=Get-UnjunkPrivacyInventory -Exclusions $Exclusions -State $State
    $recent=@($inventory.Scans | Where-Object Name -eq 'Recent shortcuts' | ForEach-Object {$_.Files})
    $prefetch=@($inventory.Scans | Where-Object Name -eq 'Prefetch' | ForEach-Object {$_.Files})
    $idm=@($inventory.Scans | Where-Object Name -eq 'IDM history' | ForEach-Object {$_.Files})
    $activityRecords=0L;foreach($database in $inventory.Activity.Databases){$activityRecords+=$database.Count}
    # Counts only: never return private document names, URLs, usage values or hashes.
    [pscustomobject]@{RecentLinks=$recent.Count;RegistryRecords=$inventory.Entries.Count;PrefetchFiles=$prefetch.Count;
        IdmHistoryFiles=$idm.Count;IdmDownloads=$inventory.IdmEntries.Count;
        CompatibilityFiles=@($inventory.Scans | Where-Object Name -eq 'Compatibility Assistant logs' | ForEach-Object {$_.Files}).Count;
        HiddenRegistryRecords=@($inventory.Entries | Where-Object {(Get-UnjunkPrivacyGroup $_.Key) -in @('Folder browsing history','Background and compatibility history')}).Count;
        ActivityRecords=$activityRecords;
        Warnings=@($inventory.Issues | Sort-Object -Unique);
        Scope='Current Windows account plus system Prefetch/PCA/compatibility history';JumpListsPreserved=$true;
        NotCovered=@('Live Amcache/SRUM databases','Filesystem journals','Unrecognized app stores and cloud-synced history')}
}

function Invoke-UnjunkPrivacyCleanup {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param([string[]]$Exclusions=@(),$State,[switch]$CloseApps)
    if($null -eq $State){$State=@{Cancel=$false;Progress=0;Message=''}}
    $core=Get-Module Unjunk.Core
    $Exclusions=@(& $core {param($paths) foreach($path in $paths){Get-UnjunkFullPath $path}} $Exclusions)
    $inventory=Get-UnjunkPrivacyInventory -Exclusions $Exclusions -State $State
    $groups=[ordered]@{}
    foreach($name in @('Recent shortcuts','Recent file history','Recent programs','IDM history','Prefetch','Folder browsing history','Background and compatibility history','Application file history','Compatibility Assistant logs','Windows activity history')){
        $groups[$name]=[pscustomobject]@{Name=$name;DeletedFiles=0L;DeletedBytes=0L;RemovedRecords=0L;SkippedItems=0L;Issues=@()}
    }
    $closures=@();$skippedFiles=0L
    if(-not $State.Cancel -and $PSCmdlet.ShouldProcess('Recent files/programs, IDM completed-download and URL history, and system Prefetch','Clear privacy traces; preserve Jump Lists')){
        $idmFiles=@($inventory.Scans | Where-Object Name -eq 'IDM history' | ForEach-Object {$_.Files})
        if($CloseApps -and ($idmFiles.Count -or $inventory.IdmEntries.Count) -and (Test-UnjunkIdmRunning)){
            $State.Message='Closing Internet Download Manager before privacy cleanup...'
            $closures=@(Close-UnjunkIdm -State $State)
            # Normal exit can save IDM history. Inventory its persisted state after it exits.
            $inventory.Scans=@($inventory.Scans | Where-Object Name -ne 'IDM history')+@(Get-UnjunkPrivacyFileScans -Exclusions $Exclusions -State $State -IdmOnly)
            $issues=New-Object 'System.Collections.Generic.List[string]'
            $inventory.IdmEntries=@(Get-UnjunkIdmHistoryEntries -Issues $issues)
            $inventory.Issues+=@($issues.ToArray())
        }
        foreach($name in $inventory.Issues){$groups[$name].Issues+=@('Some history could not be read; unavailable records were skipped.')}
        foreach($scan in $inventory.Scans){
            $group=$groups[$scan.Name]
            if($scan.Issues.Count -or $scan.SkippedLinks){$group.Issues+=@('Some files were inaccessible or links; they were skipped.')}
            foreach($file in $scan.Files){
                if($State.Cancel){break}
                if($scan.Name -eq 'IDM history' -and (Test-UnjunkIdmRunning)){$group.SkippedItems++;$skippedFiles++;$group.Issues+=@('IDM is running; close it or enable Close apps.');continue}
                $result=& $core {param($file,$root,$exclusions) Remove-UnjunkCandidate -File $file -Root $root -Exclusions $exclusions -Confirm:$false} $file $scan.Root $Exclusions
                if($result.Status -eq 'Deleted'){$group.DeletedFiles++;$group.DeletedBytes+=$result.Bytes}else{$group.SkippedItems++;$skippedFiles++}
            }
        }
        $registryCount=0
        foreach($entry in $inventory.Entries){
            if($State.Cancel){break}
            $group=$groups[(Get-UnjunkPrivacyGroup $entry.Key)]
            try{if(Remove-UnjunkPrivacyEntry $entry){$group.RemovedRecords++}else{$group.SkippedItems++}}catch{$group.SkippedItems++}
            $registryCount++;if($registryCount%100 -eq 0){Set-UnjunkPrivacyProgress $State ("Cleaning privacy: $registryCount registry records checked")}
        }
        foreach($entry in $inventory.IdmEntries){
            if($State.Cancel){break}
            $group=$groups['IDM history']
            try{if(Remove-UnjunkIdmHistoryEntry $entry){$group.RemovedRecords++}else{$group.SkippedItems++}}catch{$group.SkippedItems++}
        }
        foreach($database in $inventory.Activity.Databases){
            if($State.Cancel){break}
            $group=$groups['Windows activity history']
            try{
                $activityResult=Clear-UnjunkActivityDatabase -Path $database.Path -Exclusions $Exclusions
                $group.RemovedRecords+=$activityResult.RemovedRecords
                if(-not $activityResult.JournalCleared){$group.Issues+=@('Activity records were cleared, but the database journal was in use and could not be truncated.')}
            }
            catch{$group.SkippedItems++;$group.Issues+=@('An activity database was locked, changed or inaccessible; its transaction was not committed.')}
        }
    }
    $deleted=0L;$bytes=0L;$removed=0L;$skipped=0L;$lines=@()
    foreach($group in $groups.Values){
        $group.Issues=@($group.Issues | Sort-Object -Unique)
        $deleted+=$group.DeletedFiles;$bytes+=$group.DeletedBytes;$removed+=$group.RemovedRecords;$skipped+=$group.SkippedItems+$group.Issues.Count
        $lines+=('{0}: {1} files, {2} history records removed; {3} skipped.' -f $group.Name,$group.DeletedFiles,$group.RemovedRecords,$group.SkippedItems)
        if($group.Issues.Count){$lines+=($group.Issues -join ' ')}
    }
    $lines+='General privacy cleanup preserves Jump Lists; Explorer Home recents have a separate cleanup stage. IDM downloaded files and unfinished transfers kept. Windows can recreate history as apps run.'
    $lines+='Not erased: live Amcache/SRUM databases, filesystem journals, unrecognized app stores and cloud-synced history.'
    $fileResult=[pscustomobject]@{DeletedFiles=$deleted;DeletedBytes=$bytes;SkippedFiles=$skippedFiles;QueuedFiles=0;QueuedBytes=0L;Results=@();
        AppClosures=$closures;ClosedProcesses=@($closures | Where-Object Status -in @('Closed','Terminated')).Count}
    [pscustomobject]@{ExitCode=0;ItemsCleaned=($deleted+$removed);RemovedPrivacyRecords=$removed;SkippedItems=$skipped;PrivacyBreakdown=@($groups.Values);
        FileResult=$fileResult;RestartRequired=$false;Output=($lines -join "`r`n")}
}

Export-ModuleMember -Function Get-UnjunkPrivacyPreview,Invoke-UnjunkPrivacyCleanup
