#requires -Version 5.1
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Core.psm1')

function Get-UnjunkHistoryDiskMeasurement($Run,[string]$Drive) {
    $matches=@($Run.Disks | Where-Object Drive -eq $Drive)
    if(-not $matches.Count){return $null}
    $disk=$matches[0]
    if($null -eq $disk.PSObject.Properties['FinalFreeBytes'] -or $null -eq $disk.FinalFreeBytes){return $null}
    $bytes=0.0
    if(-not [double]::TryParse([string]$disk.FinalFreeBytes,[ref]$bytes) -or [double]::IsNaN($bytes) -or [double]::IsInfinity($bytes) -or $bytes -lt 0){return $null}
    return $disk
}

function New-UnjunkHistoryView {
    param([object[]]$Runs=@(),[string]$SelectedRunId)
    [xml]$markup=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 Title="Unjunk | Cleanup timeline" Width="1120" Height="850" MinWidth="880" MinHeight="740" WindowStartupLocation="CenterScreen"
 Background="#10151C" Foreground="#E8EDF4" FontFamily="Segoe UI" FontSize="13">
 <Window.Resources>
  <Style TargetType="Button"><Setter Property="Padding" Value="14,7"/><Setter Property="Background" Value="#273341"/><Setter Property="Foreground" Value="#E8EDF4"/><Setter Property="BorderThickness" Value="0"/></Style>
  <Style TargetType="DataGridColumnHeader"><Setter Property="Background" Value="#202B38"/><Setter Property="Foreground" Value="#B8C8D9"/><Setter Property="Padding" Value="8"/></Style>
  <Style TargetType="DataGridRow"><Setter Property="Background" Value="#161F29"/><Setter Property="Foreground" Value="#E8EDF4"/><Setter Property="MinHeight" Value="30"/></Style>
 </Window.Resources>
 <Grid Margin="26">
  <Grid.RowDefinitions><RowDefinition Height="76"/><RowDefinition Height="100"/><RowDefinition Height="Auto"/><RowDefinition Height="165"/><RowDefinition Height="80"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
  <Grid><StackPanel><TextBlock Text="Cleanup timeline" FontSize="28" FontWeight="SemiBold"/><TextBlock x:Name="SelectedDate" Text="No cleanup history yet" Foreground="#95A7BB" Margin="0,5,0,0"/></StackPanel>
   <StackPanel HorizontalAlignment="Right"><TextBlock Text="MEASURED DRIVE" Foreground="#95A7BB" FontSize="10" Margin="0,0,0,5"/><ComboBox x:Name="DriveBox" Width="125" Padding="6"/></StackPanel>
  </Grid>
  <Grid Grid.Row="1"><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
   <Border Background="#19232E" CornerRadius="7" Padding="16,10" Margin="0,0,12,10"><StackPanel><TextBlock Text="COUNTED ITEMS REMOVED" Foreground="#93A6BC" FontSize="11"/><TextBlock x:Name="ItemsText" Text="--" FontSize="28" FontWeight="SemiBold"/><TextBlock x:Name="BreakdownText" Foreground="#93A6BC" FontSize="11" TextTrimming="CharacterEllipsis"/></StackPanel></Border>
   <Border Grid.Column="1" Background="#19232E" CornerRadius="7" Padding="16,10" Margin="0,0,12,10"><StackPanel><TextBlock Text="NET SPACE FREED" Foreground="#93A6BC" FontSize="11"/><TextBlock x:Name="FreedText" Text="--" FontSize="28" FontWeight="SemiBold"/><TextBlock Text="Measured change during this run" Foreground="#93A6BC" FontSize="11"/></StackPanel></Border>
   <Border Grid.Column="2" Background="#193834" CornerRadius="7" Padding="16,10" Margin="0,0,0,10"><StackPanel><TextBlock Text="FINAL FREE SPACE" Foreground="#9ABFB4" FontSize="11"/><TextBlock x:Name="FreeText" Text="--" FontSize="28" FontWeight="SemiBold" Foreground="#80E5CE"/><TextBlock x:Name="DriveText" Foreground="#9ABFB4" FontSize="11"/></StackPanel></Border>
  </Grid>
  <Grid Grid.Row="2" Margin="0,6,0,8"><Grid.RowDefinitions><RowDefinition/><RowDefinition/></Grid.RowDefinitions>
   <TextBlock Text="ITEMS REMOVED / FINAL FREE SPACE" Foreground="#93A6BC" FontSize="11"/><TextBlock x:Name="RunStatus" HorizontalAlignment="Right" Foreground="#80E5CE"/>
   <TextBlock x:Name="ItemsLegend" Grid.Row="1" Foreground="#80E5CE" FontSize="11" Margin="0,5,0,0"/>
   <TextBlock Grid.Row="1" Text="Separate scales" HorizontalAlignment="Center" Foreground="#8397AE" FontSize="10" Margin="0,5,0,0"/>
   <TextBlock x:Name="SpaceLegend" Grid.Row="1" HorizontalAlignment="Right" Foreground="#A7BAFF" FontSize="11" Margin="0,5,0,0" ToolTip="Blue bars use the labeled, zoomed range for this drive so changes are visible. Their baseline may be above zero. Hover a bar or select a date for actual free space."/>
  </Grid>
  <Border Grid.Row="3" Background="#161F29" CornerRadius="7" Padding="8"><Canvas x:Name="Chart" Background="Transparent" ClipToBounds="True" Cursor="SizeWE" ToolTip="Click or drag across the chart to choose a cleanup run."/></Border>
  <Grid Grid.Row="4" Margin="0,9,0,9"><Grid.RowDefinitions><RowDefinition Height="28"/><RowDefinition Height="22"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
   <Slider x:Name="Scrubber" Minimum="0" Maximum="1" SmallChange="1" LargeChange="1" TickFrequency="1" IsSnapToTickEnabled="True" IsMoveToPointEnabled="True" VerticalAlignment="Center" ToolTip="Drag to choose a cleanup date. Left and right arrow keys also work."/>
   <TextBlock x:Name="FirstDate" Grid.Row="1" Foreground="#93A6BC" FontSize="11"/><TextBlock x:Name="PositionText" Grid.Row="1" HorizontalAlignment="Center" Foreground="#B8C8D9" FontSize="11"/><TextBlock x:Name="LastDate" Grid.Row="1" HorizontalAlignment="Right" Foreground="#93A6BC" FontSize="11"/>
   <TextBlock Grid.Row="2" Text="Drag the handle or the chart to move between cleanup dates." HorizontalAlignment="Center" Foreground="#80E5CE" FontSize="12"/>
  </Grid>
  <Grid Grid.Row="5"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="76"/></Grid.RowDefinitions>
   <TextBlock Text="STAGES IN THE SELECTED CLEANUP" Foreground="#93A6BC" FontSize="11" Margin="0,0,0,8"/>
   <DataGrid x:Name="Stages" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True" CanUserAddRows="False" CanUserDeleteRows="False" HeadersVisibility="Column" RowHeaderWidth="0" SelectionMode="Single" Background="#161F29" Foreground="#E8EDF4" BorderBrush="#283747" GridLinesVisibility="None">
    <DataGrid.Columns><DataGridTextColumn Header="Cleanup period" Binding="{Binding Name}" Width="*" MinWidth="300"/><DataGridTextColumn Header="Started" Binding="{Binding TimeLabel}" Width="95"/><DataGridTextColumn Header="Duration" Binding="{Binding DurationLabel}" Width="85"/><DataGridTextColumn Header="Items" Binding="{Binding ItemLabel}" Width="105"/><DataGridTextColumn Header="Result" Binding="{Binding Status}" Width="90"/></DataGrid.Columns>
   </DataGrid>
   <TextBox x:Name="Details" Grid.Row="2" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Background="#19232E" Foreground="#B8C8D9" BorderThickness="0" Padding="10" Margin="0,8,0,0"/>
  </Grid>
  <Grid Grid.Row="6" Margin="0,14,0,0"><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
   <TextBlock Text="Counts include files, unused app copies, privacy records, cleared event channels and snapshots reported removed. Windows does not supply item counts for Disk Cleanup or DISM. Free-space changes can include other disk activity." TextWrapping="Wrap" Foreground="#8397AE" FontSize="11" Margin="0,0,18,0"/>
   <Button x:Name="CloseButton" Grid.Column="1" Content="Close" VerticalAlignment="Center" IsCancel="True"/>
  </Grid>
 </Grid>
</Window>
'@
    $reader=New-Object Xml.XmlNodeReader $markup
    $ui=@{Window=[Windows.Markup.XamlReader]::Load($reader);Runs=@($Runs | Sort-Object StartedAt | Select-Object -Last 120);Dragging=$false;
        Format=(Get-Command Format-UnjunkSize -CommandType Function);DiskMeasurement=(Get-Command Get-UnjunkHistoryDiskMeasurement -CommandType Function);Index=0;Current=$null}
    foreach($name in @('SelectedDate','DriveBox','ItemsText','BreakdownText','FreedText','FreeText','DriveText','RunStatus','ItemsLegend','SpaceLegend','Chart','Scrubber','FirstDate','PositionText','LastDate','Stages','Details','CloseButton')) { $ui[$name]=$ui.Window.FindName($name) }
    $ui.Draw={
        $ui.Chart.Children.Clear()
        $width=$ui.Chart.ActualWidth; $height=$ui.Chart.ActualHeight
        if($width -lt 1 -or $height -lt 1 -or -not $ui.Runs.Count) { return }
        $maximum=[Math]::Max(1,[double](($ui.Runs | Measure-Object CountedItems -Maximum).Maximum))
        $drive=[string]$ui.DriveBox.SelectedItem
        $diskValues=@(foreach($run in $ui.Runs){$measurement=& $ui.DiskMeasurement $run $drive;if($null -ne $measurement){[double]$measurement.FinalFreeBytes}})
        # Scale across all displayed runs, not the selected run, so scrubbing never changes bar heights.
        $spaceMinimum=0.0; $spaceMaximum=1.0
        if($diskValues.Count){
            $bounds=$diskValues | Measure-Object -Minimum -Maximum
            $span=[double]$bounds.Maximum-[double]$bounds.Minimum
            # Leave headroom at both ends; a flat history stays level without a zero-width scale.
            $padding=if($span -gt 0){[Math]::Max([double]1MB,$span*0.1)}else{[Math]::Max([double]1MB,[double]$bounds.Maximum*0.01)}
            $spaceMinimum=[Math]::Max(0.0,[double]$bounds.Minimum-$padding)
            $spaceMaximum=[double]$bounds.Maximum+$padding
        }
        $ui.SpaceScale=[pscustomobject]@{Minimum=$spaceMinimum;Maximum=$spaceMaximum}
        $ui.ItemsLegend.Text=('Items removed / max {0:N0}' -f $maximum)
        $ui.SpaceLegend.Text=if($diskValues.Count){'Free after on '+$drive+' / zoomed '+(& $ui.Format $spaceMinimum)+' to '+(& $ui.Format $spaceMaximum)}else{'Free after on '+$drive+' / not measured'}
        $slot=$width/$ui.Runs.Count; $plot=[Math]::Max(1,$height-28)
        for($i=0;$i -lt $ui.Runs.Count;$i++) {
            $entry=$ui.Runs[$i]; $selected=($i -eq $ui.Index)
            $bar=New-Object Windows.Shapes.Rectangle
            $bar.Width=[Math]::Max(0.5,[Math]::Min(34,($slot-6)/2)); $bar.Height=[Math]::Max(2,$plot*[double]$entry.CountedItems/$maximum)
            $bar.Fill=if($selected){'#80E5CE'}else{'#365B68'}; $bar.RadiusX=3; $bar.RadiusY=3
            $bar.Tag=[pscustomobject]@{Series='Items';RunId=$entry.Id;Value=[double]$entry.CountedItems}
            $bar.ToolTip=([DateTimeOffset]::Parse($entry.StartedAt).LocalDateTime.ToString('g')+" / $($entry.CountedItems) counted items / "+$entry.Status)
            $left=$slot*$i+($slot-(2*$bar.Width+2))/2
            [Windows.Controls.Canvas]::SetLeft($bar,$left); [Windows.Controls.Canvas]::SetTop($bar,$plot-$bar.Height+5)
            $null=$ui.Chart.Children.Add($bar)
            $disk=& $ui.DiskMeasurement $entry $drive
            $spaceBar=New-Object Windows.Shapes.Rectangle
            $spaceBar.Width=$bar.Width;$spaceBar.RadiusX=3;$spaceBar.RadiusY=3
            $spaceBar.Tag=[pscustomobject]@{Series='FreeSpace';RunId=$entry.Id;Drive=$drive;Value=$(if($null -ne $disk){[double]$disk.FinalFreeBytes}else{$null})}
            if($null -ne $disk){
                $fraction=([double]$disk.FinalFreeBytes-$spaceMinimum)/($spaceMaximum-$spaceMinimum)
                $spaceBar.Height=[Math]::Max(2.0,$plot*[Math]::Max(0.0,[Math]::Min(1.0,$fraction)))
                $spaceBar.Fill=if($selected){'#A7BAFF'}else{'#455576'}
                $spaceBar.ToolTip=([DateTimeOffset]::Parse($entry.CompletedAt).LocalDateTime.ToString('g')+' / '+(& $ui.Format ([double]$disk.FinalFreeBytes))+' free after cleanup on '+$drive)
            }else{
                $spaceBar.Height=2;$spaceBar.Fill='Transparent';$spaceBar.Stroke='#8397AE';$spaceBar.StrokeDashArray=New-Object Windows.Media.DoubleCollection
                $spaceBar.StrokeDashArray.Add(2);$spaceBar.StrokeDashArray.Add(2)
                $spaceBar.ToolTip='Final free space was not measured for '+$drive+' in this cleanup.'
            }
            [Windows.Controls.Canvas]::SetLeft($spaceBar,$left+$bar.Width+2);[Windows.Controls.Canvas]::SetTop($spaceBar,$plot-$spaceBar.Height+5)
            $null=$ui.Chart.Children.Add($spaceBar)
            if($ui.Runs.Count -le 10 -or $selected) {
                $label=New-Object Windows.Controls.TextBlock; $label.Text=[DateTimeOffset]::Parse($entry.StartedAt).LocalDateTime.ToString('MMM d')
                $label.Foreground=if($selected){'#80E5CE'}else{'#93A6BC'}; $label.FontSize=11; $label.Width=70; $label.TextAlignment='Center'; $label.IsHitTestVisible=$false
                [Windows.Controls.Canvas]::SetLeft($label,[Math]::Max(0,[Math]::Min($width-70,$slot*($i+0.5)-35))); [Windows.Controls.Canvas]::SetTop($label,$height-18)
                $null=$ui.Chart.Children.Add($label)
            }
        }
    }.GetNewClosure()
    $ui.Update={
        if(-not $ui.Runs.Count) {
            $ui.Scrubber.IsEnabled=$false; $ui.DriveBox.IsEnabled=$false; $ui.PositionText.Text='No runs recorded'
            $ui.Details.Text='Run One-click full clean to record a cleanup. Your completed runs will appear here.'
            return
        }
        $ui.Index=[Math]::Max(0,[Math]::Min($ui.Runs.Count-1,[int][Math]::Round($ui.Scrubber.Value)))
        $entry=$ui.Runs[$ui.Index]; $ui.Current=$entry
        $start=[DateTimeOffset]::Parse($entry.StartedAt).LocalDateTime; $finish=[DateTimeOffset]::Parse($entry.CompletedAt).LocalDateTime
        $ui.SelectedDate.Text=$start.ToString('dddd, MMMM d, yyyy  h:mm:ss tt')+'  -  '+$finish.ToString('h:mm:ss tt')
        $ui.ItemsText.Text=([long]$entry.CountedItems).ToString('N0')
        $packageCount=if($null -ne $entry.PSObject.Properties['RemovedProvisionedPackages']){$entry.RemovedProvisionedPackages}else{$entry.RemovedApps}
        $packageLabel=if($null -ne $entry.PSObject.Properties['RemovedProvisionedPackages']){'unused copies'}else{'packages'}
        $ui.BreakdownText.Text=('{0:N0} files / {1:N0} {2} / {3:N0} snapshots' -f $entry.DeletedFiles,$packageCount,$packageLabel,$entry.RemovedShadows)
        if($null -ne $entry.PSObject.Properties['RemovedPrivacyRecords']){$ui.BreakdownText.Text+=(' / {0:N0} privacy records' -f $entry.RemovedPrivacyRecords)}
        if($null -ne $entry.PSObject.Properties['ClearedEventLogs']){$ui.BreakdownText.Text+=(' / {0:N0} event channels' -f $entry.ClearedEventLogs)}
        $ui.BreakdownText.ToolTip=$ui.BreakdownText.Text
        $disk=& $ui.DiskMeasurement $entry ([string]$ui.DriveBox.SelectedItem)
        $ui.FreeText.Text='Not measured'; $ui.FreedText.Text='Not measured'; $ui.DriveText.Text=[string]$ui.DriveBox.SelectedItem
        if($null -ne $disk) {
            $ui.FreeText.Text=& $ui.Format ([double]$disk.FinalFreeBytes)
            $delta=[double]$disk.NetFreedBytes
            $ui.FreedText.Text=$(if($delta -lt 0){'-'}else{'+'})+(& $ui.Format ([Math]::Abs($delta)))
            $ui.DriveText.Text=([string]$disk.Drive+' available after cleanup')
        }
        $ui.RunStatus.Text=$entry.Status+$(if($entry.RestartRequired){' / Restart required'}else{''})
        $ui.PositionText.Text=('Cleanup {0} of {1}' -f ($ui.Index+1),$ui.Runs.Count)
        $ui.Stages.ItemsSource=@(foreach($period in $entry.Periods) {
            $duration=[DateTimeOffset]::Parse($period.CompletedAt)-[DateTimeOffset]::Parse($period.StartedAt)
            [pscustomobject]@{Name=$period.Name;TimeLabel=[DateTimeOffset]::Parse($period.StartedAt).LocalDateTime.ToString('HH:mm:ss');
                DurationLabel=('{0:N1}s' -f $duration.TotalSeconds);ItemLabel=$(if($null -eq $period.ItemsCleaned){'Not reported'}else{([long]$period.ItemsCleaned).ToString('N0')});Status=$period.Status;Output=$period.Output}
        })
        $ui.Details.Text=(@($entry.Error)+@($entry.Warnings)+@('Select a stage to see its result. Detailed report: '+$entry.ReportPath) | Where-Object {$_}) -join "`r`n"
        & $ui.Draw
    }.GetNewClosure()
    $ui.Seek={param([double]$X)
        if($ui.Runs.Count -and $ui.Chart.ActualWidth -gt 0) {
            $ui.Scrubber.Value=[Math]::Max(0,[Math]::Min($ui.Runs.Count-1,[Math]::Floor($X/$ui.Chart.ActualWidth*$ui.Runs.Count)))
        }
    }.GetNewClosure()
    $ui.Scrubber.Add_ValueChanged({& $ui.Update}.GetNewClosure())
    $ui.DriveBox.Add_SelectionChanged({& $ui.Update}.GetNewClosure())
    $ui.Chart.Add_SizeChanged({& $ui.Draw}.GetNewClosure())
    $ui.Chart.Add_PreviewMouseLeftButtonDown({param($sender,$eventArgs)
        $ui.Dragging=$true; $null=$ui.Chart.CaptureMouse(); & $ui.Seek $eventArgs.GetPosition($ui.Chart).X; $eventArgs.Handled=$true
    }.GetNewClosure())
    $ui.Chart.Add_MouseMove({param($sender,$eventArgs)
        if($ui.Dragging -and $eventArgs.LeftButton -eq [Windows.Input.MouseButtonState]::Pressed) { & $ui.Seek $eventArgs.GetPosition($ui.Chart).X }
    }.GetNewClosure())
    $ui.Chart.Add_MouseLeftButtonUp({$ui.Dragging=$false; $ui.Chart.ReleaseMouseCapture()}.GetNewClosure())
    $ui.Chart.Add_LostMouseCapture({$ui.Dragging=$false}.GetNewClosure())
    $ui.Stages.Add_SelectionChanged({if($ui.Stages.SelectedItem){$ui.Details.Text=$ui.Stages.SelectedItem.Output}}.GetNewClosure())
    $ui.CloseButton.Add_Click({$ui.Window.Close()}.GetNewClosure())
    $ui.Scrubber.Maximum=[Math]::Max(0,$ui.Runs.Count-1)
    $ui.DriveBox.ItemsSource=@($ui.Runs | ForEach-Object {$_.Disks} | Select-Object -ExpandProperty Drive -Unique | Sort-Object)
    $ui.DriveBox.SelectedItem=$env:SystemDrive+'\'
    if($ui.DriveBox.SelectedIndex -lt 0 -and $ui.DriveBox.Items.Count){$ui.DriveBox.SelectedIndex=0}
    $selectedIndex=[Math]::Max(0,$ui.Runs.Count-1)
    for($i=0;$i -lt $ui.Runs.Count;$i++){if($ui.Runs[$i].Id -eq $SelectedRunId){$selectedIndex=$i}}
    if($ui.Runs.Count){
        $ui.FirstDate.Text=[DateTimeOffset]::Parse($ui.Runs[0].StartedAt).LocalDateTime.ToString('MMM d, yyyy')
        $ui.LastDate.Text=[DateTimeOffset]::Parse($ui.Runs[-1].StartedAt).LocalDateTime.ToString('MMM d, yyyy')
    }
    $ui.Scrubber.Value=$selectedIndex
    & $ui.Update
    return $ui
}

function Show-UnjunkHistoryWindow {
    param([object[]]$Runs=@(),[string]$SelectedRunId)
    $view=New-UnjunkHistoryView -Runs $Runs -SelectedRunId $SelectedRunId
    $null=$view.Window.ShowDialog()
}
