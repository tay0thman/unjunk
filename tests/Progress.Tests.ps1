#requires -Version 5.1
param([string]$RenderDirectory,[string]$ProjectDirectory=(Join-Path $PSScriptRoot '..'))
$ErrorActionPreference='Stop'
Import-Module (Join-Path $ProjectDirectory 'Unjunk.Progress.psm1') -Force
. (Join-Path $ProjectDirectory 'Unjunk.Progress.Gui.ps1')
$fixture=Join-Path $PSScriptRoot ('ProgressFixture-'+[Guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $fixture
$passed=0;$ui=$null;$observer=$null
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
function Result($Id,$Status,$Items,$Bytes,$Output) {
    [pscustomobject]@{Id=$Id;Status=$Status;ItemsCleaned=$Items;Output=$Output;StartedAt='2026-11-14T12:00:00Z';CompletedAt='2026-11-14T12:00:01Z';
        FileResult=$(if($null -ne $Bytes){[pscustomobject]@{DeletedBytes=$Bytes;Results=@([pscustomobject]@{Status='Skipped';Reason='Windows update is awaiting a restart.'})}}else{$null})}
}
try {
    $jobs=@([pscustomobject]@{Id='cache';Name='Application caches';Kind='Files'},[pscustomobject]@{Id='winre-agent';Name='Windows recovery-update leftovers ($WinREAgent)';Kind='Files'},
        [pscustomobject]@{Id='winsxs';Name='WinSxS component cleanup';Kind='DISM'},[pscustomobject]@{Id='disk-full';Name='Full Windows Disk Cleanup';Kind='DiskCleanup'})
    $state=[hashtable]::Synchronized(@{Cancel=$false;Message='Scanning';StatusPath=(Join-Path $fixture 'status.json')})
    Update-UnjunkCleanupProgress -State $state -Jobs $jobs -Results @()
    Assert ($state.StepSnapshot.Total -eq 4 -and $state.StepSnapshot.FinishedCount -eq 0 -and @($state.StepSnapshot.Steps|Where-Object Status -ne 'Queued').Count -eq 0) 'The full ordered plan is queued before work starts'
    Update-UnjunkCleanupProgress -State $state -CurrentId 'cache'
    $before=$state.StepSnapshot
    Assert ($before.Steps[0].Status -eq 'Running' -and $before.Steps[1].Status -eq 'Queued') 'Only the current stage is marked running'
    $results=@((Result 'cache' 'Completed' 120 10485760 'Deleted 10 MiB / 120 files.'))
    Update-UnjunkCleanupProgress -State $state -Results $results -CurrentId 'winre-agent'
    Assert ($before.Steps[0].Status -eq 'Running' -and $state.StepSnapshot.Steps[0].Status -eq 'Completed') 'Published snapshots cannot change underneath the UI'
    Assert ($state.StepSnapshot.FinishedCount -eq 1 -and $state.StepSnapshot.CountedItems -eq 120 -and $state.StepSnapshot.DeletedBytes -eq 10485760) 'Completed item and byte totals appear while later stages still run'
    $results+=Result 'winre-agent' 'Partial' 0 0 'Deleted 0 B / 0 files; 1 skipped.'
    Update-UnjunkCleanupProgress -State $state -Results $results -CurrentId 'winsxs'
    Assert ($state.StepSnapshot.Steps[1].Status -eq 'Skipped' -and $state.StepSnapshot.Steps[1].Summary -match '^Windows update is awaiting a restart') 'Recovery skip reasons appear before generic zero-deletion text'
    $results[1].FileResult|Add-Member -NotePropertyName RemovedDirectories -NotePropertyValue 1
    Update-UnjunkCleanupProgress -State $state -Results $results -CurrentId 'winsxs'
    Assert ($state.StepSnapshot.Steps[1].Status -eq 'Partial') 'A stage that removed empty folders is not mislabeled as fully skipped'
    $results[1].FileResult.RemovedDirectories=0;Update-UnjunkCleanupProgress -State $state -Results $results -CurrentId 'winsxs'
    $state.Message='DISM / foreground window / elapsed 00:01:02';Write-UnjunkProgressStatus $state
    $saved=Get-Content -LiteralPath $state.StatusPath -Raw|ConvertFrom-Json
    if($state.StatusWriteError){throw $state.StatusWriteError}
    Assert ($saved.Snapshot.FinishedCount -eq 2 -and $saved.Snapshot.Steps[2].Status -eq 'Running' -and $saved.Message -match 'DISM') 'Native heartbeat updates preserve earlier stage results in the status file'
    $results+=Result 'winsxs' 'Completed' $null $null 'Windows component cleanup completed.'
    Update-UnjunkCleanupProgress -State $state -Results $results -CurrentId 'disk-full'
    Assert ($state.StepSnapshot.Steps[2].ItemsText -eq 'Windows managed' -and $state.StepSnapshot.Steps[2].SizeText -eq '-' -and $state.StepSnapshot.CountedItems -eq 120) 'Native cleanup does not invent item or byte measurements'
    $state.Message='Windows Disk Cleanup is running. Completed stages are listed below.'
    $ui=New-UnjunkProgressWindow $state;& $ui.Refresh
    Assert ($ui.Stages.Items.Count -eq 4 -and $ui.Summary.Text -match '3 / 4' -and $ui.Progress.Value -eq 75) 'The window renders the ordered live checklist and completed totals'
    $ui.Stages.SelectedItem=$ui.Stages.Items[1]
    Assert ($ui.Details.Text -match 'awaiting a restart') 'Selecting a skipped stage exposes the complete explanation'
    if($RenderDirectory) {
        $null=New-Item -ItemType Directory -Path $RenderDirectory -Force
        $ui.Window.Content.Background='#10151C';$ui.Window.Content.Measure([Windows.Size]::new(1120,780));$ui.Window.Content.Arrange([Windows.Rect]::new(0,0,1120,780));$ui.Window.Content.UpdateLayout()
        $bitmap=New-Object Windows.Media.Imaging.RenderTargetBitmap(1120,780,96,96,[Windows.Media.PixelFormats]::Pbgra32);$bitmap.Render($ui.Window.Content)
        $encoder=New-Object Windows.Media.Imaging.PngBitmapEncoder;$encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
        $stream=[IO.File]::Create((Join-Path $RenderDirectory 'cleanup-progress.png'));try{$encoder.Save($stream)}finally{$stream.Dispose()}
    }
    $ui.Stop.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    Assert ($state.Cancel -and -not $ui.Stop.IsEnabled) 'Stop requests cancellation without terminating the running operation'
    Update-UnjunkCleanupProgress -State $state -Results $results -Finished
    Assert ($state.StepSnapshot.Steps[3].Status -eq 'Not run' -and $state.StepSnapshot.Steps[0].Status -eq 'Completed') 'Cancellation preserves completed work and identifies unrun stages'
    $ui.Complete=$true;& $ui.Refresh
    Assert ($ui.Done.IsEnabled -and -not $ui.Stop.IsEnabled -and -not $ui.Progress.IsIndeterminate) 'Finished results remain reviewable before opening the timeline'
    $ui.Window.Close();$ui=$null
    $results+=Result 'disk-full' 'Failed' $null $null 'Fixture native failure'
    Update-UnjunkCleanupProgress -State $state -Results $results -Finished
    Assert ($state.StepSnapshot.Steps[3].Status -eq 'Failed' -and $state.StepSnapshot.FinishedCount -eq 4 -and $state.StepSnapshot.CountedItems -eq 120) 'Failures are visible without claiming new cleanup totals'
    Assert ((Format-UnjunkProgressText $state.StepSnapshot) -match 'Completed \| Application caches' -and (Format-UnjunkProgressText $state.StepSnapshot) -match 'Failed \| Full Windows') 'Main and manual panels can display both completed and failed stages'

    # Exercise the real STA worker and dispatcher with a harmless replacement module.
    $fake=Join-Path $fixture 'FixtureWorker.psm1'
    $source=@'
Import-Module '__PROGRESS__'
function Invoke-UnjunkOneClick {
    param($Exclusions,$CloseApps,$ProgressState)
    $jobs=@(1..3|ForEach-Object {[pscustomobject]@{Id=[string]$_;Name=('Fixture stage '+$_);Kind='Files'}})
    $results=@();Update-UnjunkCleanupProgress -State $ProgressState -Jobs $jobs -Results $results
    foreach($job in $jobs) {
        if($ProgressState.Cancel){break}
        $ProgressState.Message='Running '+$job.Name;Update-UnjunkCleanupProgress -State $ProgressState -CurrentId $job.Id
        Start-Sleep -Milliseconds 600
        $results+= [pscustomobject]@{Id=$job.Id;Status='Completed';ItemsCleaned=2;Output='Fixture only';FileResult=[pscustomobject]@{DeletedBytes=50}}
        Update-UnjunkCleanupProgress -State $ProgressState -Results $results
    }
    $ProgressState.Message='Fixture finished';Update-UnjunkCleanupProgress -State $ProgressState -Finished
    [pscustomobject]@{Id='fixture';Status='Completed';Stages=$results.Count;Cancelled=[bool]$ProgressState.Cancel}
}
Export-ModuleMember -Function Invoke-UnjunkOneClick
'@
    $source=$source.Replace('__PROGRESS__',([IO.Path]::GetFullPath((Join-Path $ProjectDirectory 'Unjunk.Progress.psm1'))).Replace("'","''"))
    [IO.File]::WriteAllText($fake,$source)
    $script:factory=${function:New-UnjunkProgressWindow}
    function New-UnjunkProgressWindow($State) {
        $script:observedUI=& $script:factory $State
        $script:observedUI.Window.WindowStartupLocation='Manual';$script:observedUI.Window.Left=-20000;$script:observedUI.Window.Top=-20000;$script:observedUI.Window.ShowInTaskbar=$false
        return $script:observedUI
    }
    foreach($cancel in @($false,$true)) {
        $script:observedUI=$null;$script:sawRunning=$false;$script:sawCompletedDuringRun=$false;$script:cancelFixture=$cancel
        $observer=New-Object Windows.Threading.DispatcherTimer;$observer.Interval=[TimeSpan]::FromMilliseconds(30)
        $observer.Add_Tick({
            $view=$script:observedUI;if($null -eq $view){return}
            if(@($view.Stages.Items|Where-Object Status -eq 'Running').Count){
                $script:sawRunning=$true
                if(@($view.Stages.Items|Where-Object Status -eq 'Completed').Count){$script:sawCompletedDuringRun=$true}
                if($script:cancelFixture){$view.Stop.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}
            }
            if($view.Done.IsEnabled){$view.Done.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}
        })
        $observer.Start();$run=Show-UnjunkOneClickProgress -ModulePath $fake;$observer.Stop();$observer=$null
        if($cancel){Assert ($run.Cancelled -and $run.Stages -lt 3 -and $run.Stages -gt 0) 'Stop crosses the STA worker boundary and leaves later stages unrun'}
        else{Assert ($run.Stages -eq 3 -and $script:sawRunning -and $script:sawCompletedDuringRun) ("The real UI stays responsive and displays accomplishments during background work (stages=$($run.Stages), running=$script:sawRunning, completed=$script:sawCompletedDuringRun)")}
    }
    Write-Host "All $passed progress checks passed."
} finally {
    if($observer){$observer.Stop()}
    if($ui){$ui.Complete=$true;$ui.Window.Close()}
    $full=[IO.Path]::GetFullPath($fixture);$parent=[IO.Path]::GetFullPath($PSScriptRoot)+'\'
    if(-not $full.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $full -Leaf) -notlike 'ProgressFixture-*'){throw 'Unsafe fixture path'}
    Remove-Item -LiteralPath $full -Recurse -Force
}
