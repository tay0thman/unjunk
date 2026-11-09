#requires -Version 5.1
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Core.psm1')
$script:explorerHome='shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}'

function Release-UnjunkExplorerObject($Object) {
    if($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)){
        $null=[Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function New-UnjunkExplorerSession {
    $shell=$null;$folder=$null;$items=$null
    try {
        $shell=New-Object -ComObject Shell.Application
        $folder=$shell.Namespace($script:explorerHome)
        if($null -eq $folder){throw 'Explorer Home is unavailable.'}
        $items=$folder.Items()
        [pscustomobject]@{Shell=$shell;Folder=$folder;Items=$items}
    } catch {
        foreach($object in @($items,$folder,$shell)){Release-UnjunkExplorerObject $object}
        throw 'Explorer Home inventory is unavailable.'
    }
}

function Close-UnjunkExplorerSession($Session) {
    if($null -ne $Session){foreach($object in @($Session.Items,$Session.Folder,$Session.Shell)){Release-UnjunkExplorerObject $object}}
}

function Get-UnjunkExplorerItemId([string]$Path) {
    if([string]::IsNullOrWhiteSpace($Path)){throw 'Explorer item identity unavailable.'}
    $hash=[Security.Cryptography.SHA256]::Create()
    try{([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($Path.ToUpperInvariant())))).Replace('-','')}
    finally{$hash.Dispose()}
}

function Test-UnjunkExplorerRemoveAction([string]$Name) {
    # Eligibility hint only. No context-menu action is ever invoked by cleanup.
    return $Name.Replace('&','').Trim() -in @('Remove from list','Remove from Recent','Remove from Quick access','Remove from Home')
}

function Get-UnjunkExplorerFileAttributes([string]$Path) {
    [IO.File]::GetAttributes($Path)
}

function Get-UnjunkExplorerItemInfo($Item,[string[]]$Exclusions=@()) {
    $id=Get-UnjunkExplorerItemId ([string]$Item.Path)
    $pinned=$false
    try{$pinned=($Item.ExtendedProperty('System.Home.IsPinned') -eq $true)}catch{}
    if($Item.IsFolder){
        # Shell containers include ZIP archives. Only a physical directory (or a virtual
        # namespace folder) is protected as a folder; an archive file still has history.
        try {
            if(-not $Item.IsFileSystem -or ((Get-UnjunkExplorerFileAttributes ([string]$Item.Path)) -band [IO.FileAttributes]::Directory) -ne 0){
                return [pscustomobject]@{Id=$id;Kind='Folder';ActionIndex=-1}
            }
        } catch {
            # Confirmed pins remain protected even when their target cannot be read.
            if($pinned){return [pscustomobject]@{Id=$id;Kind='Pinned';ActionIndex=-1}}
            # Unknown metadata must be an inventory issue, not a false empty success.
            throw
        }
    }
    $core=Get-Module Unjunk.Core
    if([string]$Item.Path -match '^[a-zA-Z]:\\'){
        $excluded=& $core {param($path,$exclusions) Test-UnjunkExcluded (Get-UnjunkFullPath $path) $exclusions} ([string]$Item.Path) $Exclusions
        if($excluded){return [pscustomobject]@{Id=$id;Kind='Excluded';ActionIndex=-1}}
    }
    $verbs=$null;$actions=@()
    try {
        $verbs=$Item.Verbs()
        for($i=0;$i -lt $verbs.Count;$i++){
            $verb=$null
            try {
                $verb=$verbs.Item($i);$name=([string]$verb.Name).Replace('&','').Trim()
                if($name -match '^(Unpin from |Remove from Favorites$)'){$pinned=$true}
                if(Test-UnjunkExplorerRemoveAction $name){$actions+=@($i)}
            } finally {Release-UnjunkExplorerObject $verb}
        }
    } finally {Release-UnjunkExplorerObject $verbs}
    $kind=if($pinned){'Pinned'}elseif($actions.Count -eq 1){'Recent'}else{'Unsupported'}
    [pscustomobject]@{Id=$id;Kind=$kind;ActionIndex=$(if($kind -eq 'Recent'){$actions[0]}else{-1})}
}

function Get-UnjunkExplorerHistory {
    param([string[]]$Exclusions=@(),$State)
    $core=Get-Module Unjunk.Core
    $Exclusions=@(& $core {param($paths)foreach($path in $paths){Get-UnjunkFullPath $path}} $Exclusions)
    $session=$null;$entries=New-Object 'System.Collections.Generic.List[object]';$issues=0
    try {
        $session=New-UnjunkExplorerSession
        for($i=0;$i -lt $session.Items.Count;$i++){
            if($null -ne $State -and $State.Cancel){break}
            $item=$null
            try{$item=$session.Items.Item($i);$entries.Add((Get-UnjunkExplorerItemInfo $item $Exclusions))}
            catch{$issues++}finally{Release-UnjunkExplorerObject $item}
        }
    } finally {Close-UnjunkExplorerSession $session}
    # Ids are temporary, one-way identities, not filenames. Cleanup reports retain only counts.
    [pscustomobject]@{Schema=1;Computer=$env:COMPUTERNAME;UserSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value;
        ScannedAt=[DateTime]::UtcNow.ToString('o');Entries=$entries.ToArray();Issues=$issues;
        RecentFiles=@($entries | Where-Object Kind -eq 'Recent').Count;FoldersKept=@($entries | Where-Object Kind -eq 'Folder').Count;
        FavoritesKept=@($entries | Where-Object Kind -eq 'Pinned').Count;Excluded=@($entries | Where-Object Kind -eq 'Excluded').Count;
        Unsupported=@($entries | Where-Object Kind -eq 'Unsupported').Count;Exclusions=$Exclusions}
}

function Remove-UnjunkExplorerRecentFile([string]$Path) {
    if(-not ('Unjunk.ExplorerRecentNative' -as [type])){
        Add-Type -Path (Join-Path $PSScriptRoot 'Unjunk.ExplorerRecentNative.cs') -ErrorAction Stop
    }
    [Unjunk.ExplorerRecentNative]::RemoveFileHistory($Path)
}

function Invoke-UnjunkExplorerRemoveAction($Item,[string]$ExpectedId,[string[]]$Exclusions=@()) {
    $info=Get-UnjunkExplorerItemInfo $Item $Exclusions
    if($info.Id -ne $ExpectedId -or $info.Kind -ne 'Recent'){return $false}
    $verbs=$null;$verb=$null
    try {
        $verbs=$Item.Verbs();$verb=$verbs.Item($info.ActionIndex)
        if(-not (Test-UnjunkExplorerRemoveAction ([string]$verb.Name))){return $false}
        # Remove from list hides future opens as well. Clear only this file's Explorer
        # usage through IApplicationDestinations; no context-menu verb is executed.
        Remove-UnjunkExplorerRecentFile ([string]$Item.Path)
        return $true
    } finally {Release-UnjunkExplorerObject $verb;Release-UnjunkExplorerObject $verbs}
}

function Invoke-UnjunkExplorerHistoryCleanup {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param([Parameter(Mandatory)]$Scan,$State)
    if($null -eq $State){$State=@{Cancel=$false;Progress=0;Message=''}}
    if($Scan.Schema -ne 1 -or $Scan.Computer -ne $env:COMPUTERNAME -or $Scan.UserSid -ne [Security.Principal.WindowsIdentity]::GetCurrent().User.Value){throw 'Explorer history scan belongs to another account or computer.'}
    $age=([DateTime]::UtcNow-[DateTime]::Parse($Scan.ScannedAt).ToUniversalTime()).TotalMinutes
    if($age -lt -1 -or $age -gt 30){throw 'Explorer history scan expired. Scan again.'}
    $targets=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach($entry in $Scan.Entries){
        if($entry.Id -notmatch '^[0-9A-Fa-f]{64}$'){throw 'Invalid Explorer history identity.'}
        if($entry.Kind -eq 'Recent'){$null=$targets.Add($entry.Id)}
    }
    $attempted=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $confirmed=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $failed=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $skipped=0;$issues=@();$session=$null;$remainingCount=$Scan.RecentFiles;$batches=0
    if(-not $State.Cancel -and $PSCmdlet.ShouldProcess("$($targets.Count) Explorer Home recent files",'Remove unpinned file entries; retain all folders, favorites and other apps Jump Lists')){
        # Home reveals older history after its visible batch is removed. Refresh until empty;
        # continuing to rely on the initial visible identities would leave that history behind.
        $current=Get-UnjunkExplorerHistory -Exclusions $Scan.Exclusions
        $deadline=[DateTime]::UtcNow.AddMinutes(2)
        while(-not $State.Cancel -and $batches -lt 100 -and [DateTime]::UtcNow -lt $deadline){
            $skipped=[Math]::Max($skipped,[int]$current.Issues)
            $targets.Clear()
            foreach($entry in $current.Entries){
                if($entry.Kind -eq 'Unsupported'){$null=$failed.Add($entry.Id)}
                if($entry.Kind -eq 'Recent' -and -not $attempted.Contains($entry.Id)){$null=$targets.Add($entry.Id)}
            }
            if(-not $targets.Count){break}
            $batches++;$requested=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
            $State.Message="Clearing Explorer Home recent files: batch $batches; $($confirmed.Count) confirmed removed..."
            if($State.ContainsKey('StatusPath')){try{@{Finished=$false;UpdatedAt=[DateTime]::UtcNow.ToString('o');Message=$State.Message}|ConvertTo-Json|Set-Content -LiteralPath $State.StatusPath -Encoding UTF8 -ErrorAction Stop}catch{}}
            try {
                $session=New-UnjunkExplorerSession
                for($i=$session.Items.Count-1;$i -ge 0;$i--){
                    if($State.Cancel){break}
                    $item=$null;$id=$null
                    try {
                        $item=$session.Items.Item($i);$id=Get-UnjunkExplorerItemId ([string]$item.Path)
                        if(-not $targets.Contains($id) -or $attempted.Contains($id)){continue}
                        $null=$attempted.Add($id)
                        if(Invoke-UnjunkExplorerRemoveAction $item $id $Scan.Exclusions){$null=$requested.Add($id)}else{$null=$failed.Add($id)}
                    } catch {if($id){$null=$failed.Add($id)}else{$skipped++};$issues+=@('Explorer did not accept a recent-file removal.')}
                    finally{Release-UnjunkExplorerObject $item}
                }
            } finally {Close-UnjunkExplorerSession $session;$session=$null}
            $verified=$false;$verifyDeadline=[DateTime]::UtcNow.AddSeconds(3)
            do {
                try {
                    $current=Get-UnjunkExplorerHistory -Exclusions $Scan.Exclusions
                    if($current.Issues){throw 'Incomplete verification.'}
                    $remaining=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
                    foreach($entry in $current.Entries){$null=$remaining.Add($entry.Id)}
                    $verified=$true
                    if(@($requested | Where-Object {$remaining.Contains($_)}).Count -eq 0){break}
                }catch{$verified=$false}
                if([DateTime]::UtcNow -lt $verifyDeadline){Start-Sleep -Milliseconds 200}
            }while([DateTime]::UtcNow -lt $verifyDeadline)
            if(-not $verified){foreach($id in $requested){$null=$failed.Add($id)};$issues+=@('Explorer removals could not be verified; they were not counted as cleaned.');break}
            foreach($id in $requested){if($remaining.Contains($id)){$null=$failed.Add($id)}else{$null=$confirmed.Add($id)}}
        }
        $remainingCount=$current.RecentFiles
        foreach($entry in $current.Entries){if($entry.Kind -in @('Recent','Unsupported')){$null=$failed.Add($entry.Id)}}
        $skipped+=$failed.Count
        if($remainingCount){$issues+=@('Some recent entries remain or were recreated. Unsuccessful removals are not retried indefinitely.')}
        if($current.Issues -or $current.Unsupported){$issues+=@('Some Explorer items could not be identified or have no supported history-removal action.')}
    }
    $output="Explorer Home: removed $($confirmed.Count) recent-file entries in $batches batches; $remainingCount eligible entries remain; $skipped skipped. All folders and pinned favorites were kept. Files can appear again when reopened. Other apps' Jump Lists and original files were not targeted. Only Explorer's recent-history store is eligible for this action."
    if($issues.Count){$output+="`r`n"+(@($issues | Sort-Object -Unique) -join ' ')}
    [pscustomobject]@{ExitCode=0;ItemsCleaned=[long]$confirmed.Count;RemovedPrivacyRecords=[long]$confirmed.Count;SkippedItems=$skipped;RemainingRecentFiles=$remainingCount;Batches=$batches;
        RestartRequired=$false;FileResult=$null;Output=$output}
}

Export-ModuleMember -Function Get-UnjunkExplorerHistory,Invoke-UnjunkExplorerHistoryCleanup
