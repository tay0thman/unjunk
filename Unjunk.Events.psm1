#requires -Version 5.1
Set-StrictMode -Version 2.0

function Get-UnjunkEventLogNames {
    @([Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession.GetLogNames())
}

function Get-UnjunkEventLogRecordCount([string]$Name) {
    $info=[Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession.GetLogInformation($Name,[Diagnostics.Eventing.Reader.PathType]::LogName)
    return $info.RecordCount
}

function Clear-UnjunkEventLog([string]$Name) {
    # Let Windows' event-log utility request its normal privileges (including Security).
    # No channel is disabled, no policy is changed, and access failures remain failures.
    if($Name.Contains('"') -or $Name.Contains([char]0)){throw 'Unsupported event channel name.'}
    $process=New-Object Diagnostics.Process
    $process.StartInfo=New-Object Diagnostics.ProcessStartInfo
    $process.StartInfo.FileName=Join-Path $env:SystemRoot 'System32\wevtutil.exe'
    $process.StartInfo.Arguments='cl "'+$Name+'"'
    $process.StartInfo.UseShellExecute=$false;$process.StartInfo.CreateNoWindow=$true
    $process.StartInfo.RedirectStandardOutput=$true;$process.StartInfo.RedirectStandardError=$true
    try {
        if(-not $process.Start()){throw 'Windows event-log utility did not start.'}
        $stdout=$process.StandardOutput.ReadToEndAsync();$stderr=$process.StandardError.ReadToEndAsync()
        $process.WaitForExit()
        $output=$stdout.GetAwaiter().GetResult()+$stderr.GetAwaiter().GetResult()
        if($process.ExitCode -ne 0){throw ('Windows refused to clear the channel (exit '+$process.ExitCode+'): '+$output.Trim())}
    } finally {$process.Dispose()}
}

function Get-UnjunkEventLogs {
    param($State)
    foreach($name in @(Get-UnjunkEventLogNames | Sort-Object -Unique)) {
        if($null -ne $State -and $State.Cancel){break}
        if($null -ne $State){$State.Message='Reading Event Viewer channel: '+$name}
        $count=$null;$issue=''
        try{$count=Get-UnjunkEventLogRecordCount $name}catch{$issue=$_.Exception.Message}
        [pscustomobject]@{Name=$name;RecordCount=$count;ReadError=$issue}
    }
}

function Invoke-UnjunkEventLogCleanup {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param([object[]]$Logs=@(),$State)
    if($null -eq $State){$State=@{Cancel=$false;Message=''}}
    $current=@(Get-UnjunkEventLogNames)
    $results=New-Object 'System.Collections.Generic.List[object]'
    foreach($log in @($Logs | Sort-Object Name -Unique)) {
        if($State.Cancel){break}
        $name=[string]$log.Name
        $State.Message='Clearing Event Viewer channel: '+$name
        $count=$null
        try {
            if([string]::IsNullOrWhiteSpace($name) -or $name -notin $current){throw 'Channel is no longer registered; not cleared.'}
            try{$count=Get-UnjunkEventLogRecordCount $name}catch{ } # A read failure does not widen scope or disable the channel.
            if($null -ne $count -and $count -eq 0){
                $results.Add([pscustomobject]@{Name=$name;Status='Empty';RecordsBefore=0;Reason='Already empty.'});continue
            }
            if(-not $PSCmdlet.ShouldProcess($name,'Clear this Event Viewer channel, including any security/audit history')){
                $results.Add([pscustomobject]@{Name=$name;Status='Preview';RecordsBefore=$count;Reason='No log was cleared.'});continue
            }
            Clear-UnjunkEventLog $name
            $results.Add([pscustomobject]@{Name=$name;Status='Cleared';RecordsBefore=$count;Reason='Windows accepted the clear. New events, including the clear itself, may be recorded.'})
        } catch {
            $results.Add([pscustomobject]@{Name=$name;Status='Skipped';RecordsBefore=$count;Reason=$_.Exception.Message})
        }
    }
    $cleared=@($results | Where-Object Status -eq 'Cleared').Count
    $skipped=@($results | Where-Object Status -eq 'Skipped').Count
    $details=($results | ForEach-Object {$_.Status+' | '+$_.Name+' | '+$_.Reason}) -join "`r`n"
    # Count successfully cleared channels, not an invented count of records removed during concurrent logging.
    [pscustomobject]@{ExitCode=0;RestartRequired=$false;ItemsCleaned=[long]$cleared;ClearedLogs=$cleared;SkippedLogs=$skipped;
        Output=("Cleared $cleared Event Viewer channels; $skipped skipped. Logging stays enabled. Counts represent channels, not individual events.`r`n"+$details);Results=$results.ToArray()}
}

Export-ModuleMember -Function Get-UnjunkEventLogs,Invoke-UnjunkEventLogCleanup
