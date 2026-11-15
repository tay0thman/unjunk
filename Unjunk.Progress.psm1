#requires -Version 5.1
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Core.psm1')

function Write-UnjunkProgressStatus($State) {
    if(-not $State.ContainsKey('StatusPath')) { return }
    $snapshot=if($State.ContainsKey('StepSnapshot')){$State.StepSnapshot}else{$null}
    $finished=$State.ContainsKey('Finished') -and $State.Finished
    $data=[pscustomobject]@{Finished=[bool]$finished;UpdatedAt=[DateTime]::UtcNow.ToString('o');Message=[string]$State.Message;Snapshot=$snapshot}
    $temporary=$State.StatusPath+'.tmp'
    try {
        [IO.File]::WriteAllText($temporary,($data|ConvertTo-Json -Depth 7),(New-Object Text.UTF8Encoding($false)))
        if([IO.File]::Exists($State.StatusPath)){[IO.File]::Replace($temporary,$State.StatusPath,[NullString]::Value)}else{[IO.File]::Move($temporary,$State.StatusPath)}
        $State.StatusWriteError=''
    } catch { $State.StatusWriteError=$_.Exception.Message } # Shared in-memory progress still updates.
}

function Update-UnjunkCleanupProgress {
    param($State,[object[]]$Jobs,[object[]]$Results,[string]$CurrentId='',[switch]$Finished)
    if($PSBoundParameters.ContainsKey('Jobs')) { $State.ProgressJobs=@($Jobs) }
    if($PSBoundParameters.ContainsKey('Results')) { $State.ProgressResults=@($Results) }
    $jobsNow=if($State.ContainsKey('ProgressJobs')){@($State.ProgressJobs)}else{@()}
    $resultsNow=if($State.ContainsKey('ProgressResults')){@($State.ProgressResults)}else{@()}
    $byId=@{};foreach($result in $resultsNow){$byId[$result.Id]=$result}
    $items=0L;$bytes=0L;$finishedCount=0
    $steps=@(foreach($job in $jobsNow) {
        $status='Queued';$detail='Waiting for earlier stages.';$count=$null;$size=0L;$started='';$ended='';$hasFileResult=$false
        if($job.Id -eq $CurrentId){$status='Running';$detail='Working on this stage.'}
        elseif($Finished){$status='Not run';$detail='Stopped before this stage.'}
        if($byId.ContainsKey($job.Id)) {
            $result=$byId[$job.Id];$status=$result.Status;$detail=[string]$result.Output;$finishedCount++
            if($null -ne $result.PSObject.Properties['ItemsCleaned']){$count=$result.ItemsCleaned}
            if($null -ne $count){$items+=[long]$count}
            if($null -ne $result.PSObject.Properties['FileResult'] -and $null -ne $result.FileResult){$hasFileResult=$true;$size=[long]$result.FileResult.DeletedBytes;$bytes+=$size}
            if($null -ne $result.PSObject.Properties['StartedAt']){$started=$result.StartedAt}
            if($null -ne $result.PSObject.Properties['CompletedAt']){$ended=$result.CompletedAt}
            $removedFolders=$hasFileResult -and $null -ne $result.FileResult.PSObject.Properties['RemovedDirectories'] -and $result.FileResult.RemovedDirectories -gt 0
            if($job.Id -eq 'winre-agent' -and $status -eq 'Partial' -and $count -eq 0 -and $size -eq 0 -and -not $removedFolders){
                $status='Skipped'
                if($hasFileResult -and $null -ne $result.FileResult.PSObject.Properties['Results']) {
                    $reasons=@($result.FileResult.Results | Where-Object Status -eq 'Skipped' | Select-Object -ExpandProperty Reason -Unique)
                    if($reasons.Count){$detail=($reasons -join ' ')+"`r`n"+$detail}
                }
            }
        }
        $summary=($detail -split '\r?\n')[0]
        if($summary.Length -gt 360){$summary=$summary.Substring(0,357)+'...'}
        if($detail.Length -gt 4000){$detail=$detail.Substring(0,4000)+"`r`nFull details are saved in the cleanup report."}
        $countText=if($null -ne $count){[string]$count}elseif($status -eq 'Completed' -and $job.Kind -in @('DISM','DiskCleanup')){'Windows managed'}else{'-'}
        [pscustomobject]@{Id=$job.Id;Name=$job.Name;Status=$status;ItemsCleaned=$count;ItemsText=$countText;DeletedBytes=$size;
            SizeText=$(if($hasFileResult){Format-UnjunkSize $size}else{'-'});Summary=$summary;Detail=$detail;StartedAt=$started;CompletedAt=$ended}
    })
    # Replace an immutable snapshot instead of mutating a collection while the UI enumerates it.
    $State.StepSnapshot=[pscustomobject]@{Revision=[Guid]::NewGuid().ToString('N');Steps=$steps;Total=$steps.Count;FinishedCount=$finishedCount;
        CountedItems=$items;DeletedBytes=$bytes;Finished=[bool]$Finished;
        Summary=("$finishedCount / $($steps.Count) stages finished | $items items confirmed | $(Format-UnjunkSize $bytes) in files removed")}
    Write-UnjunkProgressStatus $State
}

function Format-UnjunkProgressText($Snapshot) {
    if($null -eq $Snapshot){return ''}
    $Snapshot.Summary+"`r`n`r`n"+(($Snapshot.Steps | ForEach-Object {
        $_.Status+' | '+$_.Name+"`r`n"+$_.Summary
    }) -join "`r`n`r`n")
}

Export-ModuleMember -Function Update-UnjunkCleanupProgress,Write-UnjunkProgressStatus,Format-UnjunkProgressText
