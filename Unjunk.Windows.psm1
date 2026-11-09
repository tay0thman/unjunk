#requires -Version 5.1
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Core.psm1')
Import-Module (Join-Path $PSScriptRoot 'Unjunk.AppLeftovers.psm1')
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Privacy.psm1')
Import-Module (Join-Path $PSScriptRoot 'Unjunk.ExplorerHistory.psm1')
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Events.psm1')
$script:volumeCaches='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches'
$script:maintenanceMutexName='Global\Unjunk-WindowsMaintenance'

function Test-UnjunkWindowsAdmin {
    $principal=[Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-UnjunkWindowsCleanupBusy {
    $mutex=$null; $held=$false
    try {
        try { $mutex=[Threading.Mutex]::OpenExisting($script:maintenanceMutexName) }
        catch [Threading.WaitHandleCannotBeOpenedException] { return $false }
        catch [UnauthorizedAccessException] { return $true }
        try { $held=$mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $held=$true }
        return -not $held
    } finally { if($held){$mutex.ReleaseMutex()}; if($null -ne $mutex){$mutex.Dispose()} }
}

function Get-UnjunkDiskHandlers {
    @(Get-ChildItem -LiteralPath $script:volumeCaches -ErrorAction Stop | Select-Object -ExpandProperty PSChildName | Sort-Object -Unique)
}

function Get-UnjunkOriginalDiskHandlers {
    # The v39 script's fixed deep-clean profile. Never adopt third-party handlers or DownloadsFolder silently.
    @('Active Setup Temp Folders','BranchCache','D3D Shader Cache','Delivery Optimization Files',
      'Device Driver Packages','Diagnostic Data Viewer database files','Downloaded Program Files',
      'Internet Cache Files','Language Pack','Old ChkDsk Files','Previous Installations','Recycle Bin',
      'RetailDemo Offline Content','Service Pack Cleanup','Setup Log Files','System error memory dump files',
      'System error minidump files','Temporary Files','Temporary Setup Files','Thumbnail Cache','Update Cleanup',
      'Upgrade Discarded Files','User file versions','Windows Defender','Windows Error Reporting Files',
      'Windows Error Reporting Archive Files','Windows Error Reporting Queue Files',
      'Windows Error Reporting System Archive Files','Windows Error Reporting System Queue Files',
      'Windows ESD installation files','Windows Upgrade Log Files')
}

function Get-UnjunkWindowsOptions {
    @(
        [pscustomobject]@{Id='app-caches';Name='Application caches and junk';Note='Known browser, application and graphics caches. Uses the main scan settings. Close-apps can stop mapped apps first; otherwise running caches are skipped. Locked files are skipped.'},
        [pscustomobject]@{Id='system-temp';Name='User and Windows temporary files';Note='Scanned user Temp and Windows\Temp files, using your temporary-file age and exclusions. Locked or changed files are skipped.'},
        [pscustomobject]@{Id='app-logs';Name='Application logs and error reports';Note='Known application diagnostic logs and crash reports of all ages. Includes Autodesk reports, Google Drive logs, Zoom logs and user crash dumps. Event Viewer channels are handled by their separate cleanup option.'},
        [pscustomobject]@{Id='privacy';Name='Privacy: IDM, recents and hidden history';Note='Includes Prefetch, recent files/programs, IDM completed-download/URL/folder history, ShellBags, UserAssist/MUI, BAM/DAM, compatibility cache/PCA logs and local Windows activity databases. Close apps also closes IDM first. Keeps Jump Lists, downloaded files and unfinished transfers. ShellBag cleanup resets per-folder views; Prefetch rebuilds. Locked stores are reported.'},
        [pscustomobject]@{Id='explorer-recents';Name='Explorer Home recent files (keep pins)';Note='Clears usage history for scanned, unpinned recent files; they can appear again when reopened. Keeps all folders, pinned favorites, original files and other apps'' Jump Lists. This explicit exception updates Explorer''s own recent-history store; it does not delete Jump List databases or hide future opens.'},
        [pscustomobject]@{Id='autodesk-installers';Name='C:\Autodesk extracted installers';Note='Deletes scanned extracted installer/download files of all ages, including WI and Access, then empty extraction folders. Close apps stops the Autodesk Access updater service and UI first; its startup setting stays unchanged. Active installers and Network License Manager are preserved.'},
        [pscustomobject]@{Id='event-viewer';Name='Event Viewer logs (including security/audit)';Note='Clears scanned Windows event channels, including Security, audit and security-product channels. Logging stays enabled. Windows may record the clear itself; denied or unsupported channels are reported. This removes troubleshooting and audit history.'},
        [pscustomobject]@{Id='old-windows';Name='Old Windows installations and upgrade files';Note='Windows.old, temporary setup files and upgrade leftovers. Removes the files needed to go back to the previous Windows installation.'},
        [pscustomobject]@{Id='windows-update';Name='Windows Update cleanup';Note='Windows-managed update cleanup and Delivery Optimization files.'},
        [pscustomobject]@{Id='winsxs';Name='WinSxS component cleanup';Note='Runs DISM StartComponentCleanup in a visible progress console to remove superseded components. Does not use ResetBase or restart Windows.'},
        [pscustomobject]@{Id='live-kernel';Name='LiveKernelReports';Note='Scanned diagnostic files under Windows\LiveKernelReports. Removes saved kernel crash evidence; locked files are skipped.'},
        [pscustomobject]@{Id='memory-dumps';Name='System memory dumps and minidumps';Note='Windows\MEMORY.DMP and files under Windows\Minidump. Removes saved crash evidence; locked files are skipped.'},
        [pscustomobject]@{Id='shadow-copies';Name='Shadow copies and restore points';Note='Deletes the scanned client-accessible shadow copies on local volumes, including System Restore snapshots. Those restore points and previous file versions cannot be recovered.'},
        [pscustomobject]@{Id='disk-full';Name='Full Windows Disk Cleanup (foreground)';Note='Runs every registered Disk Cleanup category with its normal progress dialog visible. Includes Recycle Bin, Downloads, old installations and diagnostics where supported. Windows scans all drives. Supersedes the separate old-installation and update options.'}
    )
}

function Get-UnjunkShadowCopies {
    @(Get-CimInstance -ClassName Win32_ShadowCopy -ErrorAction Stop | Where-Object ClientAccessible | ForEach-Object {
        [pscustomobject]@{Id=[string]$_.ID;Volume=[string]$_.VolumeName;Created=[string]$_.InstallDate}
    })
}

function Get-UnjunkWindowsApps {
    param([string[]]$Exclusions=@(),$State)
    Get-UnjunkAppLeftovers -Exclusions $Exclusions -State $State
}
function Invoke-UnjunkWindowsScan {
    param([int]$TempAgeDays=7,[string[]]$Exclusions=@(),[bool]$CrowdStrikeSafe=$true,$State)
    if($null -eq $State) { $State=@{Cancel=$false;Progress=0;Message=''} }
    if($CrowdStrikeSafe) { throw 'Turn off CrowdStrike Safe cleaning to scan Windows maintenance targets.' }
    if(-not (Test-UnjunkWindowsAdmin)) { throw 'Open Unjunk as administrator to scan Windows cleanup and all-user apps.' }
    $State.Message='Reading Windows Disk Cleanup categories...'
    $handlers=@(Get-UnjunkDiskHandlers)
    $State.Progress=15; $State.Message='Scanning system temporary files and crash dumps...'
    $fileScan=Invoke-UnjunkSystemFileScan -TempAgeDays $TempAgeDays -Exclusions $Exclusions -IncludeCaches -IncludeLogs -State $State
    $apps=@(); $appError=''
    if(-not $State.Cancel) {
        $State.Message='Finding machine-wide leftovers of removed apps...'
        try { $apps=@(Get-UnjunkWindowsApps -Exclusions $Exclusions -State $State | Sort-Object Name,Version,Kind) }
        catch { $appError=$_.Exception.Message }
    }
    $shadows=@(); $shadowError=''
    if(-not $State.Cancel) {
        $State.Message='Reading shadow copies and restore points...'
        try { $shadows=@(Get-UnjunkShadowCopies) } catch { $shadowError=$_.Exception.Message }
    }
    $privacy=$null;$privacyError=''
    if(-not $State.Cancel){
        $State.Message='Counting privacy traces while preserving Jump Lists...'
        try{$privacy=Get-UnjunkPrivacyPreview -Exclusions $Exclusions -State $State}catch{$privacyError=$_.Exception.Message}
    }
    $explorer=$null;$explorerError=''
    if(-not $State.Cancel){
        $State.Message='Reading Explorer Home recent files and protected pins...'
        try{$explorer=Get-UnjunkExplorerHistory -Exclusions $Exclusions -State $State}catch{$explorerError='Explorer Home inventory is unavailable.'}
    }
    $eventLogs=@();$eventLogError=''
    if(-not $State.Cancel){
        $State.Message='Reading Event Viewer channels, including security and audit logs...'
        try{$eventLogs=@(Get-UnjunkEventLogs -State $State)}catch{$eventLogError=$_.Exception.Message}
    }
    $State.Progress=100
    [pscustomobject]@{Version=41;ScannedAt=[DateTime]::UtcNow.ToString('o');Computer=$env:COMPUTERNAME;
        UserSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value;CrowdStrikeSafe=$false;Cancelled=[bool]$State.Cancel;
        Handlers=$handlers;Apps=$apps;AppError=$appError;Shadows=$shadows;ShadowError=$shadowError;Privacy=$privacy;PrivacyError=$privacyError;ExplorerHistory=$explorer;ExplorerHistoryError=$explorerError;EventLogs=$eventLogs;EventLogError=$eventLogError;FileScan=$fileScan;Options=@(Get-UnjunkWindowsOptions)}
}

function Get-UnjunkWindowsPlan {
    param([Parameter(Mandatory)]$Scan,[string[]]$OptionIds=@(),[string[]]$AppIds=@(),[ValidateSet('Full','Original')][string]$DiskCleanupProfile='Full')
    $options=@(Get-UnjunkWindowsOptions); $jobs=New-Object 'System.Collections.Generic.List[object]'
    foreach($id in $OptionIds) { if($id -notin $options.Id) { throw "Unknown Windows cleanup option: $id" } }
    $ids=@($options.Id | Where-Object {$_ -in $OptionIds}); $selectedApps=@($AppIds | Sort-Object -Unique)
    if($selectedApps.Count -and $Scan.AppError) { throw 'App inventory is incomplete. Scan again before pruning apps.' }
    foreach($id in $selectedApps) { if($id -notin @($Scan.Apps | Select-Object -ExpandProperty Id)) { throw 'An app is not in this scan. Scan again.' } }
    if('disk-full' -in $ids) { $ids=@($ids | Where-Object {$_ -notin @('old-windows','windows-update')}) }
    foreach($id in $ids) {
        $option=@($options | Where-Object Id -eq $id)[0]; $handlers=@(); $rows=@(); $kind='Files'
        switch($id) {
            'app-caches' { $rows=@($Scan.FileScan.Targets | Where-Object SystemOption -eq $id) }
            'system-temp' { $rows=@($Scan.FileScan.Targets | Where-Object SystemOption -eq $id) }
            'app-logs' { $rows=@($Scan.FileScan.Targets | Where-Object SystemOption -eq $id) }
            'autodesk-installers' { $rows=@($Scan.FileScan.Targets | Where-Object SystemOption -eq $id) }
            'event-viewer' {
                $kind='EventLogs'
                if($null -eq $Scan.PSObject.Properties['EventLogError'] -or $Scan.EventLogError){throw 'Event Viewer inventory is unavailable. Scan again before clearing logs.'}
            }
            'privacy' {
                $kind='Privacy'
                if($null -eq $Scan.PSObject.Properties['PrivacyError'] -or $Scan.PrivacyError){throw 'Privacy inventory is unavailable. Scan again before privacy cleanup.'}
            }
            'explorer-recents' {
                $kind='ExplorerHistory'
                if($null -eq $Scan.PSObject.Properties['ExplorerHistoryError'] -or $Scan.ExplorerHistoryError -or $null -eq $Scan.ExplorerHistory){throw 'Explorer Home inventory is unavailable. Scan again before clearing its recent list.'}
            }
            'live-kernel' { $rows=@($Scan.FileScan.Targets | Where-Object SystemOption -eq $id) }
            'memory-dumps' { $rows=@($Scan.FileScan.Targets | Where-Object SystemOption -eq $id) }
            'old-windows' { $kind='DiskCleanup'; $handlers=@($Scan.Handlers | Where-Object {$_ -in @('Previous Installations','Temporary Setup Files','Upgrade Discarded Files','Windows Upgrade Log Files','Setup Log Files')}) }
            'windows-update' { $kind='DiskCleanup'; $handlers=@($Scan.Handlers | Where-Object {$_ -in @('Update Cleanup','Delivery Optimization Files')}) }
            'disk-full' {
                $kind='DiskCleanup'; $handlers=@($Scan.Handlers)
                if($DiskCleanupProfile -eq 'Original') { $handlers=@($handlers | Where-Object {$_ -in @(Get-UnjunkOriginalDiskHandlers)}) }
            }
            'winsxs' { $kind='DISM' }
            'shadow-copies' { $kind='Shadows'; if($Scan.ShadowError) { throw ('Shadow-copy inventory failed: '+$Scan.ShadowError) } }
        }
        if($kind -eq 'DiskCleanup' -and -not $handlers.Count) { throw "No Windows Disk Cleanup categories are available for $($option.Name)." }
        $detail=$option.Note
        if($id -eq 'disk-full' -and $DiskCleanupProfile -eq 'Original') { $detail='Foreground Disk Cleanup using the legacy script categories, including Recycle Bin, previous installations, update files and diagnostics.' }
        if($kind -eq 'DiskCleanup') { $detail+="`r`nCategories: "+($handlers -join ', ')+".`r`nWindows chooses eligible files at execution time on all drives; Unjunk file age and exclusions do not apply." }
        if($kind -eq 'Files') {
            $count=[long](($rows | Measure-Object FileCount -Sum).Sum)
            $bytes=[long](($rows | Measure-Object Bytes -Sum).Sum)
            $detail+="`r`n$count scanned files / $(Format-UnjunkSize $bytes)."
        }
        if($kind -eq 'Shadows') {
            $detail+="`r`n$(@($Scan.Shadows).Count) snapshot(s):`r`n"+(($Scan.Shadows | ForEach-Object {$_.Id+' | '+$_.Volume+' | '+$_.Created}) -join "`r`n")
        }
        if($kind -eq 'Privacy'){
            $detail+="`r`n$($Scan.Privacy.RecentLinks) recent shortcuts and $($Scan.Privacy.RegistryRecords) Windows registry history records."
            if($null -ne $Scan.Privacy.PSObject.Properties['PrefetchFiles']){$detail+=" $($Scan.Privacy.PrefetchFiles) Prefetch files, $($Scan.Privacy.IdmDownloads) completed IDM downloads and $($Scan.Privacy.IdmHistoryFiles) IDM history files."}
            if($null -ne $Scan.Privacy.PSObject.Properties['HiddenRegistryRecords']){$detail+=" $($Scan.Privacy.HiddenRegistryRecords) hidden registry records, $($Scan.Privacy.CompatibilityFiles) PCA logs and $($Scan.Privacy.ActivityRecords) local activity records."}
            if($null -ne $Scan.Privacy.PSObject.Properties['Warnings'] -and $Scan.Privacy.Warnings.Count){$detail+=' Unavailable records: '+($Scan.Privacy.Warnings -join ', ')+'. Other privacy categories can still run.'}
            $detail+=' Eligible records are re-read at execution; reports store separate counts. Jump Lists are preserved.'
        }
        if($kind -eq 'EventLogs'){$detail+="`r`n$(@($Scan.EventLogs).Count) scanned channels, including security/audit channels. New channels created after this scan are preserved."}
        if($kind -eq 'ExplorerHistory'){$detail+="`r`n$($Scan.ExplorerHistory.RecentFiles) currently visible recent files; $($Scan.ExplorerHistory.FoldersKept) folders and $($Scan.ExplorerHistory.FavoritesKept) pinned favorites kept. Older entries revealed by Explorer are cleared in further batches. Pins and exclusions are rechecked throughout."}
        $jobs.Add([pscustomobject]@{Id=$id;Name=$option.Name;Kind=$kind;Handlers=$handlers;TargetIds=@($rows | Select-Object -ExpandProperty Id);App=$null;Detail=$detail})
    }
    foreach($id in $selectedApps) {
        $app=@($Scan.Apps | Where-Object Id -eq $id)[0]
        if($app.Kind -notin @('DataLeftover','ProvisionedLeftover') -or (Test-UnjunkProtectedAppName $app.Name)){throw 'Installed apps and protected apps cannot be pruned. Scan again for leftovers.'}
        $appDetail=if($app.Kind -eq 'DataLeftover'){'Delete scanned residual data under this local profile only if the app remains absent from every account.'}else{'Remove an unused machine-wide provisioned copy only if no account has the app installed. It will not be automatically installed for future users.'}
        $jobs.Add([pscustomobject]@{Id=$id;Name=($app.Name+' / '+$app.Scope);Kind='AppLeftovers';Handlers=@();TargetIds=@();App=$app;
            Detail=($appDetail+"`r`nPackage: "+$app.PackageName+"`r`nLocation: "+$app.Root)})
    }
    # One ordered sequence for preview and execution: junk, leftovers, temp/logs, snapshots, DISM, Disk Cleanup.
    return @($jobs | Sort-Object @{Expression={
        if($_.Id -eq 'app-caches'){0}elseif($_.Kind -eq 'AppLeftovers'){1}elseif($_.Id -eq 'autodesk-installers'){1.5}elseif($_.Id -eq 'system-temp'){2}
        elseif($_.Id -eq 'app-logs'){3}elseif($_.Id -eq 'live-kernel'){4}elseif($_.Id -eq 'memory-dumps'){5}
        elseif($_.Kind -eq 'EventLogs'){5.5}elseif($_.Kind -eq 'Privacy'){6}elseif($_.Kind -eq 'ExplorerHistory'){6.1}elseif($_.Kind -eq 'Shadows'){7}elseif($_.Kind -eq 'DISM'){8}else{9}
    }})
}

function Invoke-UnjunkWindowsNative {
    param([ValidateSet('cleanmgr.exe','Dism.exe','vssadmin.exe')][string]$Executable,[string[]]$Arguments,$State)
    $binary=Join-Path $env:SystemRoot ('System32\'+$Executable)
    $directory=Join-Path $env:LOCALAPPDATA 'Unjunk\Reports'
    $null=New-Item -ItemType Directory -Path $directory -Force -ErrorAction Stop
    $prefix=Join-Path $directory ('native-'+[Guid]::NewGuid().ToString('N'))
    $outPath=$prefix+'.out.txt'; $errPath=$prefix+'.err.txt'
    # Wait for the command itself; Start-Process -Wait also waits for lingering descendant hosts.
    # Keep reporting elapsed time without guessing a percentage or terminating Windows servicing.
    $foreground=$Executable -in @('Dism.exe','cleanmgr.exe')
    $resultPath=$prefix+'.result.json'
    $parameters=@{FilePath=$binary;ArgumentList=$Arguments;WindowStyle='Hidden';PassThru=$true;ErrorAction='Stop'}
    if($Executable -eq 'Dism.exe') {
        $operation=if('/StartComponentCleanup' -in $Arguments){'StartComponentCleanup'}elseif('/AnalyzeComponentStore' -in $Arguments){'AnalyzeComponentStore'}elseif($Arguments.Count -eq 1 -and $Arguments[0] -eq '/?'){'Help'}else{throw 'Unsupported DISM operation.'}
        $parameters.FilePath="$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $parameters.ArgumentList=@('-NoProfile','-ExecutionPolicy','Bypass','-File',('"'+(Join-Path $PSScriptRoot 'Unjunk.Dism.ps1')+'"'),'-Operation',$operation,'-OutputPath',('"'+$outPath+'"'),'-ResultPath',('"'+$resultPath+'"'))
        $parameters.WindowStyle='Normal'
    } elseif($Executable -eq 'cleanmgr.exe') {
        $parameters.WindowStyle='Normal'
    } else {
        $parameters.RedirectStandardOutput=$outPath; $parameters.RedirectStandardError=$errPath
    }
    $process=Start-Process @parameters
    # Windows PowerShell needs a retained process handle to retrieve the exit code without -Wait.
    try {
      $null=$process.Handle
      $started=[DateTime]::UtcNow
      do {
        if($null -ne $State) {
            $elapsed=[DateTime]::UtcNow-$started
            $mode=if($foreground){'foreground window'}else{'running'}
            $State.Message=('{0} / {1} / {2:00}:{3:00}:{4:00} elapsed / process {5}. Stop waits for this operation.' -f $Executable,$mode,[Math]::Floor($elapsed.TotalHours),$elapsed.Minutes,$elapsed.Seconds,$process.Id)
            if($State.ContainsKey('StatusPath')) {
                try { @{Finished=$false;UpdatedAt=[DateTime]::UtcNow.ToString('o');Message=$State.Message} | ConvertTo-Json | Set-Content -LiteralPath $State.StatusPath -Encoding UTF8 -ErrorAction Stop } catch { }
            }
        }
      } while(-not $process.WaitForExit(1000))
      $process.WaitForExit()
      $exitCode=$process.ExitCode
      if($null -eq $exitCode){throw 'Windows did not return a process exit code.'}
      if($Executable -eq 'Dism.exe') {
          if(-not (Test-Path -LiteralPath $resultPath)){throw 'The DISM window closed without a completion result. Check the DISM log before retrying.'}
          $completion=Get-Content -LiteralPath $resultPath -Raw -ErrorAction Stop | ConvertFrom-Json
          if($null -eq $completion.ExitCode -or [int]$completion.ExitCode -ne $exitCode){throw 'DISM completion could not be verified.'}
      }
      $output=if($Executable -eq 'cleanmgr.exe'){'Windows Disk Cleanup ran in its foreground progress window.'}else{((Get-Content -LiteralPath $outPath -Raw -ErrorAction SilentlyContinue),(Get-Content -LiteralPath $errPath -Raw -ErrorAction SilentlyContinue)) -join "`r`n"}
      if($Executable -eq 'cleanmgr.exe'){$output | Set-Content -LiteralPath $outPath -Encoding UTF8 -ErrorAction Stop}
      [pscustomobject]@{ExitCode=$exitCode;Output=$output;OutputPath=$outPath;ErrorPath=if(Test-Path -LiteralPath $errPath){$errPath}else{$null};Foreground=$foreground;RestartRequired=($exitCode -eq 3010)}
    } finally { $process.Dispose() }
}

function Invoke-UnjunkComponentAnalysis {
    param($State)
    if(-not (Test-UnjunkWindowsAdmin)){throw 'Administrator access is required for WinSxS analysis.'}
    $lock=New-Object Threading.Mutex($false,$script:maintenanceMutexName);$held=$false
    try {
        try{$held=$lock.WaitOne(0)}catch [Threading.AbandonedMutexException]{$held=$true}
        if(-not $held){throw 'Another Unjunk Windows cleanup is running. Wait for it to finish before analysis.'}
        $result=Invoke-UnjunkWindowsNative -Executable Dism.exe -Arguments @('/Online','/Cleanup-Image','/AnalyzeComponentStore','/English') -State $State
        $result | Add-Member -NotePropertyName CompletedAt -NotePropertyValue ([DateTime]::UtcNow.ToString('o')) -PassThru
    } finally {if($held){$lock.ReleaseMutex()};$lock.Dispose()}
}

function Remove-UnjunkShadowCopies([object[]]$Shadows,$State) {
    $output=New-Object 'System.Collections.Generic.List[string]'
    $deleted=0L; $failed=0
    foreach($shadow in $Shadows) {
        if($State.Cancel) { break }
        if($shadow.Id -notmatch '^\{[0-9a-fA-F-]{36}\}$' -or $shadow.Volume -notmatch '^\\\\\?\\Volume\{[0-9a-fA-F-]{36}\}\\$') { throw 'Invalid shadow-copy identity. Scan again.' }
        $current=@(Get-UnjunkShadowCopies | Where-Object {$_.Id -eq $shadow.Id -and $_.Volume -eq $shadow.Volume})
        if(-not $current.Count) { $output.Add('Already absent: '+$shadow.Id); continue }
        try {
            $result=Invoke-UnjunkWindowsNative -Executable vssadmin.exe -Arguments @('delete','shadows',('/Shadow='+$shadow.Id),'/Quiet') -State $State
            if($result.ExitCode -ne 0) { throw "Shadow-copy removal failed for $($shadow.Id): $($result.Output)" }
            $deleted++; $output.Add('Deleted '+$shadow.Id+': '+$result.Output)
        } catch { $failed++; $output.Add($_.Exception.Message) }
    }
    [pscustomobject]@{ExitCode=if($failed){1}else{0};ItemsCleaned=$deleted;Output=if($output.Count){$output -join "`r`n"}else{'No shadow copies to remove.'};RestartRequired=$false}
}

function Invoke-UnjunkDiskCleanup {
    param([string[]]$Handlers,$State)
    $current=@(Get-UnjunkDiskHandlers)
    foreach($handler in $Handlers) { if($handler -notin $current) { throw 'Disk Cleanup categories changed. Scan again.' } }
    if(-not $Handlers.Count) { throw 'No Disk Cleanup categories selected.' }
    if(@(Get-Process -Name cleanmgr -ErrorAction SilentlyContinue).Count) { throw 'Another Disk Cleanup is running. Wait for it to finish.' }
    # Reserve an unused profile rather than overwrite the user's saved cleanup profiles.
    $profile=0
    foreach($candidate in 9000..9999) {
        $valueName='StateFlags{0:D4}' -f $candidate; $used=$false
        foreach($handler in $current) {
            $key=Get-Item -LiteralPath (Join-Path $script:volumeCaches $handler) -ErrorAction Stop
            if($valueName -in $key.GetValueNames()) { $used=$true; break }
        }
        if(-not $used) { $profile=$candidate; break }
    }
    if(-not $profile) { throw 'No free Disk Cleanup profile is available.' }
    $valueName='StateFlags{0:D4}' -f $profile; $written=New-Object 'System.Collections.Generic.List[string]'
    try {
        foreach($handler in $current) {
            $path=Join-Path $script:volumeCaches $handler
            $value=if($handler -in $Handlers){2}else{0}
            $null=New-ItemProperty -LiteralPath $path -Name $valueName -Value $value -PropertyType DWord -ErrorAction Stop
            $written.Add($path)
        }
        Invoke-UnjunkWindowsNative -Executable cleanmgr.exe -Arguments @("/sagerun:$profile") -State $State
    } finally {
        foreach($path in $written) { Remove-ItemProperty -LiteralPath $path -Name $valueName -ErrorAction Continue }
    }
}

function Invoke-UnjunkWindowsCleanup {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param([Parameter(Mandatory)]$Scan,[string[]]$OptionIds=@(),[string[]]$AppIds=@(),[ValidateSet('Full','Original')][string]$DiskCleanupProfile='Full',[switch]$CloseApps,$State)
    if($null -eq $State) { $State=@{Cancel=$false;Progress=0;Message=''} }
    if($Scan.CrowdStrikeSafe) { throw 'Windows maintenance is blocked by CrowdStrike Safe cleaning.' }
    if($Scan.Cancelled) { throw 'Complete a new Windows scan first.' }
    if($Scan.UserSid -ne [Security.Principal.WindowsIdentity]::GetCurrent().User.Value -or $Scan.Computer -ne $env:COMPUTERNAME) { throw 'The Windows scan belongs to another account or computer.' }
    $age=[DateTime]::UtcNow-[DateTime]::Parse($Scan.ScannedAt).ToUniversalTime()
    if($age.TotalMinutes -gt 30 -or $age.TotalMinutes -lt 0) { throw 'The Windows scan expired. Scan again.' }
    if(-not (Test-UnjunkWindowsAdmin)) { throw 'Administrator access is required for Windows maintenance.' }
    $jobs=@(Get-UnjunkWindowsPlan -Scan $Scan -OptionIds $OptionIds -AppIds $AppIds -DiskCleanupProfile $DiskCleanupProfile)
    if(-not $jobs.Count) { throw 'Select at least one cleanup option or app.' }
    # Validate all identities/categories before any cleanup; a changed full profile must be reviewed again.
    $nativeJobs=@($jobs | Where-Object Kind -eq 'DiskCleanup')
    if($nativeJobs.Count) {
        $handlers=@(Get-UnjunkDiskHandlers)
        foreach($job in $nativeJobs) {
            if(@($job.Handlers | Where-Object {$_ -notin $handlers}).Count) { throw 'Disk Cleanup categories changed. Scan again.' }
            $expected=$handlers
            if($DiskCleanupProfile -eq 'Original') { $expected=@($handlers | Where-Object {$_ -in @(Get-UnjunkOriginalDiskHandlers)}) }
            if($job.Id -eq 'disk-full' -and @($expected | Where-Object {$_ -notin $job.Handlers}).Count) { throw 'New Disk Cleanup categories appeared. Scan again to review them.' }
        }
    }
    $appJobs=@($jobs | Where-Object Kind -eq 'AppLeftovers')
    if($appJobs.Count) {
        $current=@(Get-UnjunkWindowsApps -Exclusions $Scan.FileScan.Exclusions)
        foreach($job in $appJobs) { if($job.Id -notin @($current | Select-Object -ExpandProperty Id)) { throw 'A selected app changed or is no longer removable. Scan again.' } }
    }
    $results=New-Object 'System.Collections.Generic.List[object]'; $lock=$null; $held=$false
    try {
        if(-not $WhatIfPreference) {
            $lock=New-Object Threading.Mutex($false,$script:maintenanceMutexName)
            try { $held=$lock.WaitOne(0) } catch [Threading.AbandonedMutexException] { $held=$true }
            if(-not $held) { throw 'Another Unjunk Windows cleanup is running.' }
        }
        foreach($job in $jobs) {
            if($State.Cancel) { break }
            $startedAt=[DateTime]::UtcNow.ToString('o'); $itemsCleaned=$null
            $State.Message='Running '+$job.Name+'. Stop waits for the current Windows operation to finish.'
            $State.Progress=[int](100*$results.Count/[Math]::Max(1,$jobs.Count))
            if(-not $PSCmdlet.ShouldProcess($job.Name,$job.Detail)) {
                $results.Add([pscustomobject]@{Id=$job.Id;Name=$job.Name;Status='Preview';Output=$job.Detail;RestartRequired=$false;FileResult=$null;ExitCode=$null}); continue
            }
            try {
                $fileResult=$null
                switch($job.Kind) {
                    'Files' {
                        if(-not $job.TargetIds.Count) { throw 'System-file targets are missing. Scan again.' }
                        $fileResult=Invoke-UnjunkCleanup -Scan $Scan.FileScan -TargetIds $job.TargetIds -CloseApps:$CloseApps -Confirm:$false -State $State
                        $native=[pscustomobject]@{ExitCode=0;Output="Deleted $(Format-UnjunkSize $fileResult.DeletedBytes) / $($fileResult.DeletedFiles) files; $($fileResult.SkippedFiles) skipped.";RestartRequired=$false}
                    }
                    'DiskCleanup' { $native=Invoke-UnjunkDiskCleanup -Handlers $job.Handlers -State $State }
                    'DISM' { $native=Invoke-UnjunkWindowsNative -Executable Dism.exe -Arguments @('/Online','/Cleanup-Image','/StartComponentCleanup','/NoRestart','/English') -State $State }
                    'AppLeftovers' { $native=Invoke-UnjunkAppLeftoverCleanup -App $job.App -State $State -Confirm:$false; $fileResult=$native.FileResult }
                    'Privacy' { $native=Invoke-UnjunkPrivacyCleanup -Exclusions $Scan.FileScan.Exclusions -State $State -CloseApps:$CloseApps -Confirm:$false; $fileResult=$native.FileResult }
                    'ExplorerHistory' { $native=Invoke-UnjunkExplorerHistoryCleanup -Scan $Scan.ExplorerHistory -State $State -Confirm:$false }
                    'EventLogs' { $native=Invoke-UnjunkEventLogCleanup -Logs $Scan.EventLogs -State $State -Confirm:$false }
                    'Shadows' { $native=Remove-UnjunkShadowCopies -Shadows @($Scan.Shadows) -State $State }
                }
                if($job.Kind -eq 'Files') { $itemsCleaned=[long]$fileResult.DeletedFiles }
                if($job.Kind -in @('AppLeftovers','Privacy','ExplorerHistory','EventLogs')) { $itemsCleaned=[long]$native.ItemsCleaned }
                if($job.Kind -eq 'Shadows' -and $null -ne $native.PSObject.Properties['ItemsCleaned']) { $itemsCleaned=[long]$native.ItemsCleaned }
                if($native.ExitCode -notin @(0,3010)) { throw "Windows returned exit code $($native.ExitCode). $($native.Output)" }
                $status=if($null -ne $fileResult -and $fileResult.SkippedFiles -gt 0){'Partial'}else{'Completed'}
                if($job.Kind -eq 'EventLogs' -and $native.SkippedLogs -gt 0){$status='Partial'}
                if($job.Kind -eq 'Privacy' -and $null -ne $native.PSObject.Properties['SkippedItems'] -and $native.SkippedItems -gt 0){$status='Partial'}
                if($job.Kind -eq 'ExplorerHistory' -and $native.SkippedItems -gt 0){$status='Partial'}
                $output=$native.Output
                if($job.Kind -eq 'DiskCleanup') { $output="Disk Cleanup finished. Windows does not report a per-category reclaimed-byte total.`r`n"+$output }
                $provisionedCount=if($job.Kind -eq 'AppLeftovers'){[long]$native.RemovedProvisionedPackages}else{0L}
                $privacyCount=if($job.Kind -in @('Privacy','ExplorerHistory')){[long]$native.RemovedPrivacyRecords}else{0L}
                $privacyBreakdown=if($job.Kind -eq 'Privacy' -and $null -ne $native.PSObject.Properties['PrivacyBreakdown']){@($native.PrivacyBreakdown)}else{@()}
                $results.Add([pscustomobject]@{Id=$job.Id;Name=$job.Name;Kind=$job.Kind;StartedAt=$startedAt;CompletedAt=[DateTime]::UtcNow.ToString('o');ItemsCleaned=$itemsCleaned;RemovedProvisionedPackages=$provisionedCount;RemovedPrivacyRecords=$privacyCount;PrivacyBreakdown=@($privacyBreakdown);Status=$status;Output=$output;RestartRequired=$native.RestartRequired;FileResult=$fileResult;ExitCode=$native.ExitCode})
            } catch {
                $results.Add([pscustomobject]@{Id=$job.Id;Name=$job.Name;Kind=$job.Kind;StartedAt=$startedAt;CompletedAt=[DateTime]::UtcNow.ToString('o');ItemsCleaned=$itemsCleaned;Status='Failed';Output=$_.Exception.Message;RestartRequired=$false;FileResult=$fileResult;ExitCode=$null})
            }
        }
    } finally { if($held) { $lock.ReleaseMutex() }; if($null -ne $lock) { $lock.Dispose() } }
    $State.Progress=100
    [pscustomobject]@{CompletedAt=[DateTime]::UtcNow.ToString('o');Cancelled=[bool]$State.Cancel;Results=$results.ToArray();
        Completed=@($results | Where-Object Status -eq 'Completed').Count;Failed=@($results | Where-Object Status -eq 'Failed').Count;
        Partial=@($results | Where-Object Status -eq 'Partial').Count;RestartRequired=(@($results | Where-Object RestartRequired).Count -gt 0)}
}

Export-ModuleMember -Function Get-UnjunkWindowsOptions,Invoke-UnjunkWindowsScan,Get-UnjunkWindowsPlan,Invoke-UnjunkWindowsCleanup,Test-UnjunkWindowsCleanupBusy,Invoke-UnjunkComponentAnalysis
