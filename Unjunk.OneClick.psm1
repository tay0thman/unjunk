#requires -Version 5.1
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Core.psm1')
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Windows.psm1')
Import-Module (Join-Path $PSScriptRoot 'Unjunk.AppLeftovers.psm1')
$script:oneClickMutexName='Global\Unjunk-OneClick'

function Get-UnjunkOriginalAppIds {
    param([object[]]$Apps=@())
    # Compatibility name: automatic selection is now verified leftovers only.
    @($Apps | Where-Object {$_.Kind -in @('DataLeftover','ProvisionedLeftover') -and -not (Test-UnjunkProtectedAppName $_.Name)} | Select-Object -ExpandProperty Id -Unique)
}
function Get-UnjunkDiskSpace {
    @(foreach($drive in [IO.DriveInfo]::GetDrives()) {
        try {
            if($drive.DriveType -eq [IO.DriveType]::Fixed -and $drive.IsReady) {
                [pscustomobject]@{Drive=$drive.Name;FreeBytes=[long]$drive.AvailableFreeSpace;TotalBytes=[long]$drive.TotalSize}
            }
        } catch { } # An unavailable drive is unknown, never recorded as zero free bytes.
    })
}

function Test-UnjunkOneClickAdmin {
    $principal=[Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-UnjunkSpaceChange {
    param([object[]]$Before=@(),[object[]]$After=@())
    @(foreach($prior in $Before) {
        $current=@($After | Where-Object Drive -eq $prior.Drive)
        if($current.Count -ne 1) { continue }
        [pscustomobject]@{Drive=$prior.Drive;BeforeFreeBytes=[long]$prior.FreeBytes;FinalFreeBytes=[long]$current[0].FreeBytes;
            NetFreedBytes=([long]$current[0].FreeBytes-[long]$prior.FreeBytes);TotalBytes=[long]$current[0].TotalBytes}
    })
}

function Save-UnjunkHistory {
    param([Parameter(Mandatory)]$Run,[string]$Directory=(Join-Path $env:LOCALAPPDATA 'Unjunk\History'))
    $null=New-Item -ItemType Directory -Path $Directory -Force -ErrorAction Stop
    $name='run-'+[Guid]::NewGuid().ToString('N')
    $temporary=Join-Path $Directory ($name+'.tmp')
    $path=Join-Path $Directory ($name+'.json')
    # Readers only see complete records; an interrupted write never replaces an older run.
    $Run | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $temporary -Encoding UTF8 -ErrorAction Stop
    [IO.File]::Move($temporary,$path)
    $path
}

function Get-UnjunkHistory {
    param([string]$Directory=(Join-Path $env:LOCALAPPDATA 'Unjunk\History'))
    $sid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $records=@(foreach($file in @(Get-ChildItem -LiteralPath $Directory -Filter 'run-*.json' -File -ErrorAction SilentlyContinue)) {
        try {
            $run=Get-Content -LiteralPath $file.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if($run.Schema -ne 1 -or $run.Source -ne 'Unjunk.OneClick' -or $run.Computer -ne $env:COMPUTERNAME -or $run.UserSid -ne $sid) { continue }
            $null=[DateTimeOffset]::Parse($run.StartedAt); $null=[DateTimeOffset]::Parse($run.CompletedAt)
            if($run.CountedItems -lt 0 -or $run.Status -notin @('Completed','Partial','Failed')) { continue }
            foreach($required in @('Id','DeletedFiles','RemovedApps','RemovedShadows','RestartRequired','Disks','Periods','Warnings','ReportPath','Error')) {
                if($null -eq $run.PSObject.Properties[$required]) { throw 'Incomplete history record.' }
            }
            foreach($period in $run.Periods) {
                $null=[DateTimeOffset]::Parse($period.StartedAt); $null=[DateTimeOffset]::Parse($period.CompletedAt)
                if($null -eq $period.PSObject.Properties['ItemsCleaned'] -or $null -eq $period.PSObject.Properties['Output']) { throw 'Incomplete cleanup period.' }
            }
            foreach($disk in $run.Disks) { $null=[long]$disk.FinalFreeBytes; $null=[long]$disk.NetFreedBytes; $null=$disk.Drive }
            $run | Add-Member -NotePropertyName HistoryPath -NotePropertyValue $file.FullName -Force
            $run
        } catch { } # One damaged record must not hide other completed runs.
    })
    @($records | Sort-Object StartedAt)
}

function Invoke-UnjunkOneClick {
    [CmdletBinding()]
    param([string[]]$Exclusions=@(),[bool]$CloseApps=$true)
    if(-not (Test-UnjunkOneClickAdmin)) { throw 'One-click cleanup requires administrator access.' }
    if(Test-UnjunkWindowsCleanupBusy){throw 'Another Unjunk Windows cleanup is still running. It must finish before a new run can start.'}
    if(@(Get-Process -Name cleanmgr -ErrorAction SilentlyContinue).Count) { throw 'Windows Disk Cleanup is already running. Let that operation finish before starting the one-click cleaner.' }
    $lock=New-Object Threading.Mutex($false,$script:oneClickMutexName)
    $held=$false
    try {
        try { $held=$lock.WaitOne(0) } catch [Threading.AbandonedMutexException] { $held=$true }
        if(-not $held) { throw 'A one-click cleanup is already running.' }
        $run=[pscustomobject]@{Schema=1;Source='Unjunk.OneClick';Id=[Guid]::NewGuid().ToString('N');Computer=$env:COMPUTERNAME;
            UserSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value;StartedAt=[DateTime]::UtcNow.ToString('o');CompletedAt=$null;
            Profile='All cleanup categories with machine-wide app leftovers';CrowdStrikeSafeOverride=$true;CloseApps=$CloseApps;Status='Failed';CountedItems=0L;DeletedFiles=0L;DeletedFileBytes=0L;
            RemovedApps=0L;RemovedProvisionedPackages=0L;RemovedPrivacyRecords=0L;RemovedShadows=0L;ClearedEventLogs=0L;UnreportedItemCounts=$false;RestartRequired=$false;Disks=@();Periods=@();Warnings=@();ReportPath='';Error=''}
        $before=@(Get-UnjunkDiskSpace)
        $statusDirectory=Join-Path $env:LOCALAPPDATA 'Unjunk'
        $null=New-Item -ItemType Directory -Path $statusDirectory -Force -ErrorAction Stop
        $statusPath=Join-Path $statusDirectory 'one-click-status.json'
        $state=@{Cancel=$false;Progress=0;Message='Scanning full cleanup targets...';StatusPath=$statusPath}
        @{Finished=$false;UpdatedAt=$run.StartedAt;Message=$state.Message} | ConvertTo-Json | Set-Content -LiteralPath $statusPath -Encoding UTF8 -ErrorAction Stop
        try {
            # The override belongs to this invocation only. Never write settings.json or Falcon configuration.
            $scan=Invoke-UnjunkWindowsScan -CrowdStrikeSafe $false -TempAgeDays 0 -Exclusions $Exclusions -State $state
            # Derive the preset from the live catalog so new categories cannot be silently omitted.
            $options=@(Get-UnjunkWindowsOptions | Select-Object -ExpandProperty Id)
            $apps=@()
            if($scan.AppError) { $run.Warnings+=@('App pruning skipped: '+$scan.AppError) }
            else { $apps=@(Get-UnjunkOriginalAppIds -Apps @($scan.Apps)) }
            if($scan.PrivacyError){$run.Warnings+=@('Privacy cleanup skipped: '+$scan.PrivacyError);$options=@($options | Where-Object {$_ -ne 'privacy'})}
            if($null -ne $scan.PSObject.Properties['ExplorerHistoryError'] -and $scan.ExplorerHistoryError){$run.Warnings+=@('Explorer Home cleanup skipped: '+$scan.ExplorerHistoryError);$options=@($options | Where-Object {$_ -ne 'explorer-recents'})}
            if($scan.EventLogError){$run.Warnings+=@('Event Viewer cleanup skipped: '+$scan.EventLogError);$options=@($options | Where-Object {$_ -ne 'event-viewer'})}
            if($scan.ShadowError) {
                $run.Warnings+=@('Shadow-copy cleanup skipped: '+$scan.ShadowError)
                $options=@($options | Where-Object {$_ -ne 'shadow-copies'})
            }
            foreach($row in $scan.FileScan.Targets) {
                if(@($row.Issues).Count) { $run.Warnings+=@($row.Name+': '+($row.Issues -join '; ')) }
            }
            $report=Invoke-UnjunkWindowsCleanup -Scan $scan -OptionIds $options -AppIds $apps -DiskCleanupProfile Full -CloseApps:$CloseApps -State $state -Confirm:$false
            $run.RestartRequired=[bool]$report.RestartRequired
            $run.Periods=@(foreach($result in $report.Results) {
                if($null -ne $result.ItemsCleaned) { $run.CountedItems+=[long]$result.ItemsCleaned }
                elseif($result.Kind -in @('DiskCleanup','DISM')) { $run.UnreportedItemCounts=$true }
                if($null -ne $result.FileResult) {
                    $run.DeletedFiles+=[long]$result.FileResult.DeletedFiles; $run.DeletedFileBytes+=[long]$result.FileResult.DeletedBytes
                }
                if($result.Kind -eq 'AppLeftovers' -and $null -ne $result.PSObject.Properties['RemovedProvisionedPackages']) { $run.RemovedProvisionedPackages+=[long]$result.RemovedProvisionedPackages }
                if($result.Kind -in @('Privacy','ExplorerHistory') -and $null -ne $result.PSObject.Properties['RemovedPrivacyRecords']) { $run.RemovedPrivacyRecords+=[long]$result.RemovedPrivacyRecords }
                if($result.Kind -eq 'Shadows' -and $null -ne $result.ItemsCleaned) { $run.RemovedShadows+=[long]$result.ItemsCleaned }
                if($result.Kind -eq 'EventLogs' -and $null -ne $result.ItemsCleaned) { $run.ClearedEventLogs+=[long]$result.ItemsCleaned }
                [pscustomobject]@{Name=$result.Name;Kind=$result.Kind;StartedAt=$result.StartedAt;CompletedAt=$result.CompletedAt;
                    ItemsCleaned=$result.ItemsCleaned;Status=$result.Status;Output=$result.Output;
                    PrivacyBreakdown=$(if($null -ne $result.PSObject.Properties['PrivacyBreakdown']){@($result.PrivacyBreakdown)}else{@()})}
            })
            $run.Status=if($report.Failed -and -not ($report.Completed+$report.Partial)){'Failed'}elseif($report.Failed -or $report.Partial -or $report.Cancelled -or $run.Warnings.Count){'Partial'}else{'Completed'}
            try { $run.ReportPath=Save-UnjunkReport -Report $report -Kind 'one-click' }
            catch { $run.Warnings+=@('Detailed report could not be saved: '+$_.Exception.Message); if($run.Status -eq 'Completed'){$run.Status='Partial'} }
        } catch { $run.Error=$_.Exception.Message; $run.Status='Failed' }
        $run.CompletedAt=[DateTime]::UtcNow.ToString('o')
        $run.Disks=@(Get-UnjunkSpaceChange -Before $before -After @(Get-UnjunkDiskSpace))
        $path=Save-UnjunkHistory -Run $run
        @{Finished=$true;UpdatedAt=$run.CompletedAt;Message=('One-click cleanup '+$run.Status.ToLowerInvariant()+'. Open Cleanup timeline for details.')} | ConvertTo-Json | Set-Content -LiteralPath $statusPath -Encoding UTF8 -ErrorAction SilentlyContinue
        $run | Add-Member -NotePropertyName HistoryPath -NotePropertyValue $path
        $run
    } finally { if($held){$lock.ReleaseMutex()}; $lock.Dispose() }
}

function Start-UnjunkOneClick {
    param([switch]$History)
    $entry=Join-Path $PSScriptRoot 'Unjunk.OneClick.ps1'
    $sid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $arguments='-NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "'+$entry+'" -ExpectedUserSid "'+$sid+'"'
    $parameters=@{FilePath="$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe";ArgumentList=$arguments;WindowStyle='Hidden';PassThru=$true;ErrorAction='Stop'}
    if($History) { $parameters.ArgumentList+=' -History' }
    else {
        $principal=[Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
        if(-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { $parameters.Verb='RunAs' }
    }
    Start-Process @parameters
}

Export-ModuleMember -Function Get-UnjunkOriginalAppIds,Get-UnjunkHistory,Invoke-UnjunkOneClick,Start-UnjunkOneClick
