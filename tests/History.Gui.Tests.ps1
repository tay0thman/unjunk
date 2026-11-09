#requires -Version 5.1
# Offscreen WPF and synthetic history. No cleanup, native processes or mouse injection.
param([string]$RenderDirectory)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '..\Unjunk.History.Gui.ps1')
$passed=0
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
function Layout-View($View,[double]$Width=1080,[double]$Height=790){
    $View.Window.Content.Measure((New-Object Windows.Size($Width,$Height)))
    $View.Window.Content.Arrange((New-Object Windows.Rect(0,0,$Width,$Height)))
    $View.Window.Content.UpdateLayout()
    & $View.Draw
}
$runs=@(foreach($i in 0..2){
    $start=[DateTime]::UtcNow.Date.AddDays(-6+3*$i).AddHours(17)
    [pscustomobject]@{Id=[string]$i;StartedAt=$start.ToString('o');CompletedAt=$start.AddMinutes(8).ToString('o');CountedItems=@(125,300,78)[$i];DeletedFiles=@(122,297,75)[$i];RemovedApps=1;RemovedShadows=2;
        Disks=@([pscustomobject]@{Drive=$env:SystemDrive+'\';FinalFreeBytes=(100+$i)*1GB;NetFreedBytes=if($i -eq 1){-1GB}else{2GB}});
        Status=if($i -eq 1){'Partial'}else{'Completed'};RestartRequired=($i -eq 2);Error='';Warnings=@();ReportPath='Fixture report';
        Periods=@([pscustomobject]@{Name='Caches and temp';Kind='Files';StartedAt=$start.ToString('o');CompletedAt=$start.AddMinutes(2).ToString('o');ItemsCleaned=75;Status='Completed';Output='Fixture result'},
            [pscustomobject]@{Name='Windows Disk Cleanup';Kind='DiskCleanup';StartedAt=$start.AddMinutes(2).ToString('o');CompletedAt=$start.AddMinutes(8).ToString('o');ItemsCleaned=$null;Status='Completed';Output='Windows did not report item counts'})}
})
$otherDrive='Z:\'
$runs[0].Disks+=@([pscustomobject]@{Drive=$otherDrive;FinalFreeBytes=0L;NetFreedBytes=0L})
$runs[2].Disks+=@([pscustomobject]@{Drive=$otherDrive;FinalFreeBytes=50GB;NetFreedBytes=3GB})
$ui=New-UnjunkHistoryView -Runs $runs
Assert ($ui.Index -eq 2 -and $ui.Current.Id -eq '2' -and $ui.ItemsText.Text -eq '78') 'Timeline initially displays the latest run with its actual count'
Layout-View $ui
Assert ($ui.Chart.ActualWidth -gt 800 -and $ui.Chart.Children.Count -eq 9) 'Chart renders paired item and final-space bars for each cleanup date'
$itemBars=@($ui.Chart.Children|Where-Object {$_.Tag -and $_.Tag.Series -eq 'Items'})
$spaceBars=@($ui.Chart.Children|Where-Object {$_.Tag -and $_.Tag.Series -eq 'FreeSpace'})
Assert ($itemBars.Count -eq 3 -and $spaceBars.Count -eq 3 -and $spaceBars[2].Tag.Value -eq 102GB) 'Second bar uses actual total free bytes after the run, not the amount freed'
Assert ($itemBars[1].Height -gt $itemBars[2].Height -and $spaceBars[2].Height -gt $spaceBars[1].Height -and $ui.SpaceLegend.Text -match 'zoomed .* to ') 'Item counts and free-space bars use separate labeled scales with an explicit zoomed range'
$plot=$ui.Chart.ActualHeight-28
Assert ($spaceBars[2].Height-$spaceBars[0].Height -gt $plot*0.6) 'A 2 GiB change around 100 GiB spans most of the plot instead of looking flat'
Assert ($ui.SpaceScale.Minimum -lt 100GB -and $ui.SpaceScale.Maximum -gt 102GB -and $spaceBars[0].Height -gt 2 -and $spaceBars[2].Height -lt $plot) 'Zoomed bounds leave room above and below the recorded range'
Assert ($spaceBars[0].ToolTip -match '100.00 GiB' -and $spaceBars[2].ToolTip -match '102.00 GiB' -and $ui.FreeText.Text -eq '102.00 GiB') 'Normalized heights retain actual free space in tooltips and the selected card'
$originalHeights=@($spaceBars | ForEach-Object {$_.Height})
$ui.Scrubber.Value=0
$scrubbedBars=@($ui.Chart.Children|Where-Object {$_.Tag -and $_.Tag.Series -eq 'FreeSpace'})
Assert (@(0..2 | Where-Object {$scrubbedBars[$_].Height -ne $originalHeights[$_]}).Count -eq 0) 'Mouse selection does not rescale the free-space history'
$ui.Scrubber.Value=2
$ui.DriveBox.SelectedItem=$otherDrive
$spaceBars=@($ui.Chart.Children|Where-Object {$_.Tag -and $_.Tag.Series -eq 'FreeSpace'})
Assert ($spaceBars[2].Tag.Value -eq 50GB -and $ui.FreeText.Text -eq '50.00 GiB' -and $ui.SpaceLegend.Text.Contains($otherDrive)) 'Changing the measured drive updates both the second bars and final-space card'
Assert ($spaceBars[0].Tag.Value -eq 0 -and $null -eq $spaceBars[1].Tag.Value -and $spaceBars[1].ToolTip -match 'not measured') 'A measured zero and an unmeasured drive remain distinct'
Assert ($ui.SpaceScale.Minimum -eq 0 -and $ui.SpaceScale.Maximum -gt 50GB -and $spaceBars[0].Height -eq 2 -and $spaceBars[2].Height -gt $plot*0.8) 'Changing drives recomputes the range and preserves a true zero baseline'
$ui.Scrubber.Value=1
Assert ($ui.FreeText.Text -eq 'Not measured') 'Missing space measurements never display as zero free space'
$ui.DriveBox.SelectedItem=$env:SystemDrive+'\';$ui.Scrubber.Value=2
& $ui.Seek 1
Assert ($ui.Index -eq 0 -and $ui.ItemsText.Text -eq '125') 'Dragging to the left selects the first cleanup and refreshes its count'
& $ui.Seek ($ui.Chart.ActualWidth/2)
Assert ($ui.Index -eq 1 -and $ui.RunStatus.Text -eq 'Partial' -and $ui.FreedText.Text.StartsWith('-')) 'Dragging across the chart updates status and negative measured space'
Assert ($ui.FreeText.Text -eq '101.00 GiB' -and $ui.SelectedDate.Text -match '2026') 'Selected date and final free space follow the selected run'
$ui.Scrubber.Value=2
Assert ($ui.Index -eq 2 -and $ui.RunStatus.Text -match 'Restart required') 'Slider change selects the corresponding run without a playback timer'
Assert ($ui.Stages.Items[1].ItemLabel -eq 'Not reported') 'Stage table distinguishes unknown native counts from zero removals'
$ui.Stages.SelectedIndex=1
Assert ($ui.Details.Text -eq 'Windows did not report item counts') 'Selecting a stage shows its recorded result through a real callback'
& $ui.Seek -500
Assert ($ui.Index -eq 0) 'Dragging outside the left edge clamps to the first run'
& $ui.Seek ($ui.Chart.ActualWidth+500)
Assert ($ui.Index -eq 2) 'Dragging outside the right edge clamps to the last run'
if($RenderDirectory){
    $ui.Window.Content.UpdateLayout()
    $null=New-Item -ItemType Directory -Path $RenderDirectory -Force
    $bitmap=New-Object Windows.Media.Imaging.RenderTargetBitmap(1080,790,96,96,[Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($ui.Window.Content)
    $encoder=New-Object Windows.Media.Imaging.PngBitmapEncoder; $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $stream=[IO.File]::Create((Join-Path $RenderDirectory 'cleanup-timeline.png'));try{$encoder.Save($stream)}finally{$stream.Dispose()}
}
$chosen=New-UnjunkHistoryView -Runs $runs -SelectedRunId '1'
Assert ($chosen.Index -eq 1) 'A completed worker can select its exact run in history'
$single=New-UnjunkHistoryView -Runs @($runs[0])
Assert ($single.Index -eq 0 -and $single.Scrubber.Maximum -eq 0) 'A first cleanup works without earlier history'
Layout-View $single
$singleBar=@($single.Chart.Children|Where-Object {$_.Tag -and $_.Tag.Series -eq 'FreeSpace'})[0]
Assert ($singleBar.Height -gt $plot*0.3 -and $singleBar.Height -lt $plot*0.7 -and $single.FreeText.Text -eq '100.00 GiB') 'One measurement uses a stable finite range and still shows its actual value'
$flatRuns=$runs | ConvertTo-Json -Depth 8 | ConvertFrom-Json
foreach($run in $flatRuns){$run.Disks=@([pscustomobject]@{Drive=$env:SystemDrive+'\';FinalFreeBytes=100GB;NetFreedBytes=0L})}
$flat=New-UnjunkHistoryView -Runs $flatRuns
Layout-View $flat
$flatBars=@($flat.Chart.Children|Where-Object {$_.Tag -and $_.Tag.Series -eq 'FreeSpace'})
Assert ($flatBars.Count -eq 3 -and @($flatBars | Where-Object {$_.Height -ne $flatBars[0].Height -or [double]::IsNaN($_.Height)}).Count -eq 0 -and $flatBars[0].Height -gt 2) 'Unchanged free space stays level without inventing a trend or dividing by zero'
foreach($run in $flatRuns){$run.Disks[0].FinalFreeBytes=0L}
& $flat.Update
$flatBars=@($flat.Chart.Children|Where-Object {$_.Tag -and $_.Tag.Series -eq 'FreeSpace'})
Assert (@($flatBars | Where-Object {$_.Height -ne 2 -or $null -eq $_.Tag.Value -or $_.Tag.Value -ne 0}).Count -eq 0 -and $flat.SpaceScale.Maximum -gt $flat.SpaceScale.Minimum) 'All-zero measurements remain measured zeros on a valid scale'
foreach($run in $flatRuns){$run.Disks[0].FinalFreeBytes=$null}
& $flat.Update
$flatBars=@($flat.Chart.Children|Where-Object {$_.Tag -and $_.Tag.Series -eq 'FreeSpace'})
Assert ($flat.FreeText.Text -eq 'Not measured' -and $flat.SpaceLegend.Text -match 'not measured' -and @($flatBars | Where-Object {$null -ne $_.Tag.Value}).Count -eq 0) 'A drive with no usable measurements does not acquire a fabricated range'
$empty=New-UnjunkHistoryView -Runs @()
Assert (-not $empty.Scrubber.IsEnabled -and $empty.PositionText.Text -eq 'No runs recorded') 'Empty history opens with a useful state'
$source=Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\Unjunk.History.Gui.ps1') -Raw
Assert ($source -notmatch 'DispatcherTimer|Storyboard|BeginAnimation|Add_Tick') 'Timeline has no animation or automatic playback'
foreach($view in @($ui,$chosen,$single,$flat,$empty)){$view.Window.Close()}
Write-Host "All $passed history GUI checks passed."
