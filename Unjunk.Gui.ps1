#requires -Version 5.1
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Core.psm1') -Force
. (Join-Path $PSScriptRoot 'Unjunk.Windows.Gui.ps1')
Import-Module (Join-Path $PSScriptRoot 'Unjunk.OneClick.psm1')
$script:oneClickProcess=$null
$script:maintenanceBusy=Test-UnjunkWindowsCleanupBusy
$script:unjunkModule=Join-Path $PSScriptRoot 'Unjunk.Core.psm1'
$script:scan=$null; $script:latestReport=$null; $script:worker=$null; $script:rows=@(); $script:reportPath=''
$script:settingsPath=Join-Path $env:LOCALAPPDATA 'Unjunk\settings.json'
$script:exclusions=@(); $script:tempAge=7; $script:crowdStrikeSafe=$true; $script:closeApps=$true
if(Test-Path -LiteralPath $script:settingsPath) {
    try {
        $settings=Get-Content -LiteralPath $script:settingsPath -Raw | ConvertFrom-Json
        $script:tempAge=[Math]::Max(1,[Math]::Min(365,[int]$settings.TempAgeDays))
        $script:exclusions=@($settings.Exclusions)
        if($null -ne $settings.PSObject.Properties['CrowdStrikeSafe']) { $script:crowdStrikeSafe=[bool]$settings.CrowdStrikeSafe }
        if($null -ne $settings.PSObject.Properties['CloseApps']) { $script:closeApps=[bool]$settings.CloseApps }
    } catch { }
}
$identity=[Security.Principal.WindowsIdentity]::GetCurrent()
$script:isAdmin=([Security.Principal.WindowsPrincipal]$identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

[xml]$xaml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 Title="Unjunk | Storage cleanup" Width="1220" Height="820" MinWidth="1020" MinHeight="720" WindowStartupLocation="CenterScreen"
 Background="#10151C" Foreground="#E8EDF4" FontFamily="Segoe UI" FontSize="13">
 <Window.Resources>
  <Style TargetType="Button"><Setter Property="Background" Value="#273341"/><Setter Property="Foreground" Value="#F1F5F9"/><Setter Property="Padding" Value="14,8"/><Setter Property="Margin" Value="0,0,8,0"/><Setter Property="BorderThickness" Value="0"/><Setter Property="Cursor" Value="Hand"/><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border x:Name="ButtonBorder" Background="{TemplateBinding Background}" CornerRadius="5" Padding="{TemplateBinding Padding}"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="ButtonBorder" Property="Opacity" Value="0.8"/></Trigger><Trigger Property="IsEnabled" Value="False"><Setter TargetName="ButtonBorder" Property="Opacity" Value="0.35"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style>
  <Style TargetType="TextBox"><Setter Property="Background" Value="#202B38"/><Setter Property="Foreground" Value="#E8EDF4"/><Setter Property="BorderBrush" Value="#39485B"/><Setter Property="Padding" Value="9,6"/><Setter Property="CaretBrush" Value="White"/></Style>
  <Style TargetType="DataGridColumnHeader"><Setter Property="Background" Value="#202B38"/><Setter Property="Foreground" Value="#B8C8D9"/><Setter Property="Padding" Value="8,10"/><Setter Property="BorderThickness" Value="0"/></Style>
  <Style TargetType="DataGridRow"><Setter Property="Background" Value="#161F29"/><Setter Property="Foreground" Value="#E8EDF4"/><Setter Property="MinHeight" Value="35"/><Style.Triggers><Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#254255"/></Trigger></Style.Triggers></Style>
  <Style TargetType="DataGridCell"><Setter Property="Padding" Value="8,6"/><Setter Property="BorderThickness" Value="0"/><Setter Property="Foreground" Value="#E8EDF4"/><Style.Triggers><Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#254255"/><Setter Property="Foreground" Value="White"/></Trigger></Style.Triggers></Style>
  <Style TargetType="ListBoxItem"><Setter Property="Padding" Value="12,9"/><Setter Property="Margin" Value="0,2"/><Style.Triggers><Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#224D4D"/><Setter Property="Foreground" Value="#80E5CE"/></Trigger></Style.Triggers></Style>
 </Window.Resources>
 <Grid Margin="24">
  <Grid.RowDefinitions><RowDefinition Height="64"/><RowDefinition Height="100"/><RowDefinition Height="*"/><RowDefinition Height="70"/></Grid.RowDefinitions>
  <Grid Grid.Row="0"><StackPanel><TextBlock Text="unjunk" FontSize="30" FontWeight="Bold"/><TextBlock Text="A clearer view of what you can clean." Foreground="#95A7BB"/></StackPanel><StackPanel HorizontalAlignment="Center"><StackPanel Orientation="Horizontal"><Button x:Name="OneClickButton" Content="One-click deep clean" Background="#54D0B3" Foreground="#102A25" HorizontalContentAlignment="Left" Margin="0,0,8,0" ToolTip="Starts immediately after Windows elevation. Every cleanup category: caches, app leftovers, Autodesk installers, temp, dumps, all Event Viewer channels (including security/audit), privacy (IDM, recents, Prefetch and hidden Windows history; separate Explorer Home cleanup keeps pins and other apps Jump Lists), shadow copies, foreground DISM, then full Disk Cleanup. The close-apps option may force-stop apps; save your work first. Overrides the local Safe profile for this run only. Continues if you close Unjunk."/><Button x:Name="HistoryButton" Content="Cleanup timeline" HorizontalContentAlignment="Left" Margin="0"/></StackPanel><TextBlock Text="Includes Explorer recents; keeps pins and app Jump Lists." Foreground="#95A7BB" FontSize="11" Margin="0,6,0,0" HorizontalAlignment="Center"/></StackPanel><StackPanel HorizontalAlignment="Right"><TextBlock x:Name="ModeText" Foreground="#80E5CE" HorizontalAlignment="Right"/><CheckBox x:Name="CrowdStrikeCheck" Content="CrowdStrike Safe cleaning" Foreground="#80E5CE" HorizontalAlignment="Right" Margin="0,7,0,0" ToolTip="Only user-profile caches and temporary files at least 7 days old. Blocks diagnostics, installers and Windows maintenance. Does not change Falcon settings or guarantee zero alerts."/><TextBlock Text="Conservative profile; Falcon alerts remain possible." FontSize="10" Foreground="#93A6BC" HorizontalAlignment="Right" Margin="0,4,0,0"/></StackPanel></Grid>
  <Grid Grid.Row="1" Margin="0,16,0,16"><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
   <Border Background="#19232E" CornerRadius="7" Padding="16,10" Margin="0,0,12,0"><StackPanel><TextBlock Text="FREE ON C:" Foreground="#93A6BC" FontSize="11"/><TextBlock x:Name="FreeText" Text="--" FontSize="26" FontWeight="SemiBold"/></StackPanel></Border>
   <Border Grid.Column="1" Background="#19232E" CornerRadius="7" Padding="16,10" Margin="0,0,12,0"><StackPanel><TextBlock Text="ELIGIBLE FILES FOUND" Foreground="#93A6BC" FontSize="11"/><TextBlock x:Name="FoundText" Text="Scan to calculate" FontSize="25" FontWeight="SemiBold"/></StackPanel></Border>
   <Border Grid.Column="2" Background="#193834" CornerRadius="7" Padding="16,10"><StackPanel><TextBlock Text="SELECTED FOR CLEANUP" Foreground="#9ABFB4" FontSize="11"/><TextBlock x:Name="SelectedText" Text="0 B" FontSize="26" FontWeight="SemiBold" Foreground="#80E5CE"/></StackPanel></Border>
  </Grid>
  <Grid Grid.Row="2"><Grid.ColumnDefinitions><ColumnDefinition Width="195"/><ColumnDefinition Width="18"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
   <DockPanel Grid.Column="0">
    <StackPanel DockPanel.Dock="Bottom" Margin="0,12,0,0">
     <CheckBox x:Name="RebootCheck" Foreground="#E4BC7A" Margin="0,0,0,12" ToolTip="Standard cleaning and administrator access required. Queues selected cache/temp file paths for the next Windows restart; does not delete them now. Off by default each launch."><TextBlock Text="Delete selected at next reboot" TextWrapping="Wrap" Width="170"/></CheckBox>
     <TextBlock Text="WINDOWS TOOLS" FontSize="11" Foreground="#7C90A7" Margin="0,0,0,10"/>
     <Button x:Name="WindowsButton" Content="Windows cleanup &amp; apps" HorizontalContentAlignment="Left" Margin="0,0,0,8"/>


     <Button x:Name="StorageButton" Content="Storage settings" HorizontalContentAlignment="Left" Margin="0,0,0,8"/>
     <Button x:Name="WinsxsButton" Content="Analyze WinSxS" HorizontalContentAlignment="Left" Margin="0,0,0,8"/>
     <Button x:Name="ReportsButton" Content="Open reports" HorizontalContentAlignment="Left" Margin="0,0,0,8"/>
     <Button x:Name="AdminButton" Content="Open as administrator" HorizontalContentAlignment="Left" Margin="0,0,0,10"/>
     <TextBlock Text="Windows cleanup offers system maintenance, app leftovers and shadow copies in Standard mode." TextWrapping="Wrap" Foreground="#7C90A7" FontSize="12" Margin="4,8,0,0"/>
    </StackPanel>
    <ListBox x:Name="CategoryList" Background="Transparent" Foreground="#B8C8D9" BorderThickness="0" ScrollViewer.HorizontalScrollBarVisibility="Disabled" ToolTip="Click a name to filter. Check a box to select that category across all search results.">
     <ListBox.ItemTemplate><DataTemplate><DockPanel Width="158"><CheckBox DockPanel.Dock="Left" IsThreeState="True" IsChecked="{Binding Checked, Mode=OneWay}" IsEnabled="{Binding CanToggle}" Tag="{Binding Name}" VerticalAlignment="Center" Margin="0,0,9,0"/><TextBlock Text="{Binding Name}" TextWrapping="Wrap" VerticalAlignment="Center"/></DockPanel></DataTemplate></ListBox.ItemTemplate>
    </ListBox>
   </DockPanel>
   <Grid Grid.Column="2"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="148"/></Grid.RowDefinitions>
    <Grid><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBox x:Name="SearchBox" ToolTip="Filter by name, category or path" Margin="0,0,12,10"/><Button x:Name="OptionsButton" Grid.Column="1" Content="Scan options" Margin="0,0,0,10"/></Grid>
    <DockPanel Grid.Row="1" Margin="0,0,0,10"><TextBlock x:Name="CountText" DockPanel.Dock="Right" VerticalAlignment="Center" Foreground="#93A6BC"/><WrapPanel><Button x:Name="RecommendedButton" Content="Select recommended"/><Button x:Name="SelectAllButton" Content="Select all"/><Button x:Name="ClearButton" Content="Clear selection"/><CheckBox x:Name="EmptyCheck" Content="Show empty" VerticalAlignment="Center" Foreground="#B8C8D9"/><CheckBox x:Name="CloseAppsCheck" Content="Close apps before cleanup" VerticalAlignment="Center" Foreground="#E4BC7A" Margin="12,6,0,6" ToolTip="Also applies to one-click cleanup. Requests normal close, then force-stops mapped cache apps and IDM after five seconds. Save your work first. System, security and installer processes are protected. C:\Autodesk cleanup also stops the Autodesk Access updater service."/></WrapPanel></DockPanel>
    <DataGrid x:Name="ResultsGrid" Grid.Row="2" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" CanUserResizeRows="False" HeadersVisibility="Column" GridLinesVisibility="None" Background="#161F29" BorderBrush="#283747" RowHeaderWidth="0" SelectionMode="Single" EnableRowVirtualization="True" ScrollViewer.CanContentScroll="True">
     <DataGrid.Columns>
      <DataGridTemplateColumn Width="42"><DataGridTemplateColumn.CellTemplate><DataTemplate><CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" IsEnabled="{Binding CanSelect}" HorizontalAlignment="Center" VerticalAlignment="Center"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn>
      <DataGridTextColumn Header="Cleanup target" Binding="{Binding Name}" Width="*" MinWidth="220" IsReadOnly="True"/>
      <DataGridTextColumn Header="Size" Binding="{Binding Size}" SortMemberPath="Bytes" Width="92" IsReadOnly="True"/>
      <DataGridTextColumn Header="Type" Binding="{Binding Risk}" Width="90" IsReadOnly="True"/>
      <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="175" IsReadOnly="True"/>
     </DataGrid.Columns>
    </DataGrid>
    <Border Grid.Row="3" Background="#19232E" CornerRadius="5" Padding="12" Margin="0,12,0,0"><TextBox x:Name="DetailsText" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Background="Transparent" BorderThickness="0" FontSize="12" Text="Start with Scan C:. Select a row to see its location, cleanup effects and largest files. Nothing is selected or deleted automatically."/></Border>
   </Grid>
  </Grid>
  <Grid Grid.Row="3" Margin="0,17,0,0"><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
   <StackPanel Margin="0,0,18,0"><TextBlock x:Name="StatusText" Text="Ready. Scan first, then choose what to clean." Foreground="#A7B9CD" TextTrimming="CharacterEllipsis"/><ProgressBar x:Name="Progress" Height="4" Minimum="0" Maximum="100" Background="#243442" Foreground="#50CDB0" BorderThickness="0" Margin="0,9,0,0"/></StackPanel>
   <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Top"><Button x:Name="ExportButton" Content="Export report" IsEnabled="False"/><Button x:Name="CancelButton" Content="Cancel" IsEnabled="False"/><Button x:Name="ScanButton" Content="Scan C:"/><Button x:Name="CleanButton" Content="Clean selected" Background="#54D0B3" Foreground="#102A25" IsEnabled="False" Margin="0"/></StackPanel>
  </Grid>
 </Grid>
</Window>
'@
$reader=New-Object Xml.XmlNodeReader $xaml
$script:window=[Windows.Markup.XamlReader]::Load($reader)
foreach($name in @('ModeText','CrowdStrikeCheck','RebootCheck','FreeText','FoundText','SelectedText','StorageButton','WinsxsButton','WindowsButton','OneClickButton','HistoryButton','ReportsButton','AdminButton','CategoryList','SearchBox','OptionsButton','CountText','RecommendedButton','SelectAllButton','ClearButton','EmptyCheck','ResultsGrid','DetailsText','StatusText','Progress','ExportButton','CancelButton','ScanButton','CleanButton')) {
    Set-Variable -Name $name -Value $script:window.FindName($name) -Scope Script
}
$script:ModeText.Text=if($script:isAdmin){'Administrator / C: drive'}else{'Standard user / C: drive'}
$script:AdminButton.IsEnabled=-not $script:isAdmin
$script:CrowdStrikeCheck.IsChecked=$script:crowdStrikeSafe
$script:CloseAppsCheck=$script:window.FindName('CloseAppsCheck')
$script:CloseAppsCheck.IsChecked=$script:closeApps
$script:WinsxsButton.ToolTip='Read-only DISM analysis. Administrator access is required.'
$script:categories=@(foreach($category in @('All categories','Windows','Browsers','Desktop apps','Graphics','Creative apps','Creative & desktop apps','Rhino plug-ins','Developer','Games','Downloaded installers','Diagnostics')) { [pscustomobject]@{Name=$category;Checked=$false;CanToggle=$false} })
$script:CategoryList.ItemsSource=$script:categories
$script:CategoryList.SelectedIndex=0
try { $script:window.Icon=New-Object Windows.Media.Imaging.BitmapImage ([Uri](Join-Path $PSScriptRoot 'unjunk.ico')) } catch { }

function Update-UnjunkFreeSpace {
    $drive=New-Object IO.DriveInfo ($env:SystemDrive+'\')
    $script:FreeText.Text=Format-UnjunkSize $drive.AvailableFreeSpace
}
function Update-UnjunkSelection {
    $selected=@($script:rows | Where-Object Selected)
    $size=[long](($selected | Measure-Object Bytes -Sum).Sum)
    $script:SelectedText.Text=Format-UnjunkSize $size
    $script:CleanButton.IsEnabled=($null -ne $script:scan -and -not $script:scan.Cancelled -and $null -eq $script:worker -and -not $script:maintenanceBusy -and $selected.Count -gt 0 -and $size -gt 0)
    $changed=$false
    foreach($category in $script:categories) {
        $state=Get-UnjunkCategoryState -Rows $script:rows -Category $category.Name -QueueOnReboot:([bool]$script:RebootCheck.IsChecked) -CloseApps:$script:closeApps
        $canToggle=$state.CanToggle -and $null -ne $script:scan -and -not $script:scan.Cancelled -and $null -eq $script:worker
        if($category.Checked -ne $state.Checked -or $category.CanToggle -ne $canToggle) {
            $category.Checked=$state.Checked; $category.CanToggle=$canToggle; $changed=$true
        }
    }
    if($changed) { $script:CategoryList.Items.Refresh() }
}
function Update-UnjunkView {
    $search=$script:SearchBox.Text.Trim(); $category=if($script:CategoryList.SelectedItem){[string]$script:CategoryList.SelectedItem.Name}else{'All categories'}
    $visible=@($script:rows | Where-Object {
        ($script:EmptyCheck.IsChecked -or $_.FileCount -gt 0 -or $_.Issues.Count -gt 0 -or $_.ScanOnly) -and
        ($category -eq 'All categories' -or $_.Category -eq $category) -and
        (-not $search -or ($_.Name+' '+$_.Root+' '+$_.Category).IndexOf($search,[StringComparison]::OrdinalIgnoreCase) -ge 0)
    })
    $script:ResultsGrid.ItemsSource=$visible
    $script:CountText.Text="$($visible.Count) targets"
    Update-UnjunkSelection
}
function Set-UnjunkBusy([bool]$Busy) {
    $script:OneClickButton.IsEnabled=(-not $Busy -and $null -eq $script:oneClickProcess -and -not $script:maintenanceBusy)
    foreach($control in @($script:ScanButton,$script:OptionsButton,$script:RecommendedButton,$script:SelectAllButton,$script:ClearButton,$script:ResultsGrid,$script:CrowdStrikeCheck,$script:WindowsButton,$script:CloseAppsCheck)) { $control.IsEnabled=-not $Busy }
    $script:CancelButton.IsEnabled=$Busy
    $script:AdminButton.IsEnabled=(-not $Busy -and -not $script:isAdmin -and -not $script:crowdStrikeSafe)
    $script:WinsxsButton.IsEnabled=(-not $Busy -and -not $script:crowdStrikeSafe)
    $script:StorageButton.IsEnabled=(-not $Busy -and -not $script:crowdStrikeSafe)
    $script:RebootCheck.IsEnabled=(-not $Busy -and -not $script:crowdStrikeSafe -and $script:isAdmin)
    if($script:crowdStrikeSafe -or -not $script:isAdmin) { $script:RebootCheck.IsChecked=$false }
    $script:CleanButton.IsEnabled=$false
}
function Start-UnjunkWork([string]$Operation, $Payload) {
    if($script:worker) { return }
    if($Operation -eq 'WinSxS' -and $script:crowdStrikeSafe) { return }
    $script:state=[hashtable]::Synchronized(@{Cancel=$false;Progress=0;Message='Starting...';Operation=$Operation})
    $ps=[PowerShell]::Create()
    $null=$ps.AddScript({
        param($ModulePath,$Operation,$Payload,$State)
        $ErrorActionPreference='Stop'
        Import-Module $ModulePath -Force
        switch($Operation) {
            Scan { Invoke-UnjunkScan -TempAgeDays $Payload.TempAgeDays -Exclusions $Payload.Exclusions -CrowdStrikeSafe $Payload.CrowdStrikeSafe -State $State }
            Clean { Invoke-UnjunkCleanup -Scan $Payload.Scan -TargetIds $Payload.Ids -QueueOnReboot:$Payload.QueueOnReboot -CloseApps:$Payload.CloseApps -State $State -Confirm:$false }
            WinSxS {
                Import-Module (Join-Path (Split-Path $ModulePath) 'Unjunk.Windows.psm1') -Force
                Invoke-UnjunkComponentAnalysis -State $State
            }
        }
    }).AddArgument($script:unjunkModule).AddArgument($Operation).AddArgument($Payload).AddArgument($script:state)
    $script:worker=[pscustomobject]@{PowerShell=$ps;Handle=$ps.BeginInvoke();Operation=$Operation}
    Set-UnjunkBusy $true
    # DISM servicing is allowed to finish; never terminate servicing mid-operation.
    if($Operation -eq 'WinSxS') { $script:CancelButton.IsEnabled=$false; $script:Progress.IsIndeterminate=$true }
}
function Show-UnjunkError([string]$Text) {
    $script:StatusText.Text=$Text
    $null=[Windows.MessageBox]::Show($script:window,$Text,'Unjunk',[Windows.MessageBoxButton]::OK,[Windows.MessageBoxImage]::Error)
}
$script:timer=New-Object Windows.Threading.DispatcherTimer
$script:timer.Interval=[TimeSpan]::FromMilliseconds(250)
$script:timer.Add_Tick({
    $script:maintenanceBusy=Test-UnjunkWindowsCleanupBusy
    $script:OneClickButton.IsEnabled=($null -eq $script:worker -and $null -eq $script:oneClickProcess -and -not $script:maintenanceBusy)
    $script:WinsxsButton.IsEnabled=($null -eq $script:worker -and -not $script:crowdStrikeSafe -and -not $script:maintenanceBusy)
    if($script:maintenanceBusy -and -not $script:worker){$script:StatusText.Text='Windows cleanup is already running. Its sequence includes DISM and Disk Cleanup.'}
    if($null -ne $script:oneClickProcess -and $script:oneClickProcess.HasExited) {
        $script:oneClickProcess.Dispose(); $script:oneClickProcess=$null
        $script:OneClickButton.IsEnabled=($null -eq $script:worker -and -not $script:maintenanceBusy)
        $script:scan=$null; foreach($row in $script:rows){$row.Selected=$false}; Update-UnjunkView; Update-UnjunkFreeSpace
        if(-not $script:worker){$script:StatusText.Text='One-click process finished. Open Cleanup timeline for its result.'}
    }
    if($null -ne $script:oneClickProcess -and -not $script:worker) {
        try { $status=Get-Content -LiteralPath (Join-Path $env:LOCALAPPDATA 'Unjunk\one-click-status.json') -Raw -ErrorAction Stop | ConvertFrom-Json; $script:StatusText.Text=[string]$status.Message } catch { }
    }
    if(-not $script:worker) { Update-UnjunkSelection; return }
    $script:Progress.Value=$script:state.Progress; $script:StatusText.Text=$script:state.Message
    if(-not $script:worker.Handle.IsCompleted) { return }
    $operation=$script:worker.Operation
    try {
        $result=@($script:worker.PowerShell.EndInvoke($script:worker.Handle))
        if($script:worker.PowerShell.HadErrors) { throw ($script:worker.PowerShell.Streams.Error | Out-String) }
        if($result.Count -ne 1) { throw 'The worker did not return a complete result.' }
        $script:latestReport=$result[0]
        if($operation -eq 'Scan') {
            $script:scan=$result[0]; $script:rows=@($script:scan.Targets | Sort-Object Bytes -Descending)
            $script:FoundText.Text=Format-UnjunkSize ([long](($script:rows | Measure-Object Bytes -Sum).Sum))
            Update-UnjunkView
            $script:StatusText.Text=if($script:scan.Cancelled){'Scan cancelled. Complete a new scan before cleaning.'}else{'Scan complete. Review the targets, then select what to clean.'}
        } elseif($operation -eq 'Clean') {
            $clean=$result[0]
            $script:StatusText.Text="Deleted $(Format-UnjunkSize $clean.DeletedBytes) / $($clean.DeletedFiles) files; $($clean.SkippedFiles) skipped. Scan again to refresh."
            if($clean.QueueOnReboot) { $script:StatusText.Text="Queued $($clean.QueuedFiles) files / $(Format-UnjunkSize $clean.QueuedBytes) for restart; $($clean.SkippedFiles) skipped. Space has not been freed yet." }
            if($clean.Cancelled) { $script:StatusText.Text='Cancelled. '+$script:StatusText.Text }
            $script:DetailsText.Text=$script:StatusText.Text+"`r`nFull file list: Open reports or Export report.`r`n`r`n"+(($clean.Results | Where-Object Status -in @('Skipped','Queued') | Select-Object -First 12 | ForEach-Object { $_.Status+' | '+$_.Path+': '+$_.Reason }) -join "`r`n")
            $script:scan=$null; foreach($row in $script:rows) { $row.Selected=$false; $row.Status='Rescan required' }
            $script:ResultsGrid.Items.Refresh(); Update-UnjunkFreeSpace
        } else {
            $script:DetailsText.Text=$result[0].Output
            $script:StatusText.Text="WinSxS analysis finished (exit $($result[0].ExitCode)). See details or export the report."
        }
        try { $script:reportPath=Save-UnjunkReport -Report $script:latestReport -Kind $operation.ToLowerInvariant() }
        catch { $script:StatusText.Text+=' Report could not be saved: '+$_.Exception.Message }
        $script:ExportButton.IsEnabled=$true
    } catch { Show-UnjunkError $_.Exception.Message }
    finally {
        $script:worker.PowerShell.Dispose(); $script:worker=$null; $script:Progress.IsIndeterminate=$false
        Set-UnjunkBusy $false; Update-UnjunkSelection
    }
})
$script:timer.Start()
$script:ScanButton.Add_Click({ Start-UnjunkWork Scan @{TempAgeDays=$script:tempAge;Exclusions=$script:exclusions;CrowdStrikeSafe=$script:crowdStrikeSafe} })
$script:RebootCheck.Add_Click({
    $script:CleanButton.Content=if($script:RebootCheck.IsChecked){'Queue for reboot'}else{'Clean selected'}
    $script:StatusText.Text=if($script:RebootCheck.IsChecked){'Reboot mode: selected cache/temp files will be queued, not deleted now. Review the selection.'}else{'Immediate cleanup mode. Locked files will be skipped.'}
})
$script:CloseAppsCheck.Add_Click({
    $script:closeApps=[bool]$script:CloseAppsCheck.IsChecked
    if(-not $script:closeApps -and -not $script:RebootCheck.IsChecked){foreach($row in $script:rows){if($row.RunningApps.Count){$row.Selected=$false}}}
    $script:ResultsGrid.Items.Refresh(); Update-UnjunkSelection
    $script:StatusText.Text=if($script:closeApps){'Selected cache apps will be closed, then force-stopped if needed. Save your work before cleanup.'}else{'Running-app caches will be skipped.'}
    try {
        $null=New-Item -ItemType Directory -Path (Split-Path $script:settingsPath) -Force
        @{TempAgeDays=$script:tempAge;Exclusions=$script:exclusions;CrowdStrikeSafe=$script:crowdStrikeSafe;CloseApps=$script:closeApps} | ConvertTo-Json | Set-Content -LiteralPath $script:settingsPath -Encoding UTF8 -ErrorAction Stop
    } catch { Show-UnjunkError ('Could not save the app-closing option: '+$_.Exception.Message) }
})
$script:CrowdStrikeCheck.Add_Click({
    $script:crowdStrikeSafe=[bool]$script:CrowdStrikeCheck.IsChecked
    $script:RebootCheck.IsChecked=$false; $script:CleanButton.Content='Clean selected'
    $script:scan=$null; $script:rows=@(); $script:FoundText.Text='Scan to calculate'; Update-UnjunkView; Set-UnjunkBusy $false
    $script:StatusText.Text='Cleaning profile changed. Scan again before selecting files.'
    try {
        $null=New-Item -ItemType Directory -Path (Split-Path $script:settingsPath) -Force
        @{TempAgeDays=$script:tempAge;Exclusions=$script:exclusions;CrowdStrikeSafe=$script:crowdStrikeSafe;CloseApps=$script:closeApps} | ConvertTo-Json | Set-Content -LiteralPath $script:settingsPath -Encoding UTF8 -ErrorAction Stop
    } catch { Show-UnjunkError ('Could not save the cleaning profile: '+$_.Exception.Message) }
})
$script:CancelButton.Add_Click({ $script:state.Cancel=$true; $script:CancelButton.IsEnabled=$false; $script:state.Message='Cancelling after the current file...' })
$script:RecommendedButton.Add_Click({
    foreach($row in $script:rows) { $row.Selected=($row.Recommended -and $row.FileCount -gt 0 -and ($row.RunningApps.Count -eq 0 -or $script:RebootCheck.IsChecked -or ($script:closeApps -and $row.Risk -eq 'Cache')) -and $row.Issues.Count -eq 0) }
    $script:ResultsGrid.Items.Refresh(); Update-UnjunkSelection
})
$script:ClearButton.Add_Click({ foreach($row in $script:rows) { $row.Selected=$false }; $script:ResultsGrid.Items.Refresh(); Update-UnjunkSelection })
$script:SelectAllButton.Add_Click({ Set-UnjunkCategorySelection -Rows $script:rows -Category 'All categories' -QueueOnReboot:([bool]$script:RebootCheck.IsChecked) -CloseApps:$script:closeApps; $script:ResultsGrid.Items.Refresh(); Update-UnjunkSelection })
$script:SearchBox.Add_TextChanged({ Update-UnjunkView })
$script:CategoryList.Add_SelectionChanged({ Update-UnjunkView })
$script:CategoryList.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent,[Windows.RoutedEventHandler]{
    param($sender,$eventArgs)
    $checkbox=$eventArgs.Source -as [Windows.Controls.CheckBox]
    if(-not $checkbox -or $script:worker -or -not $script:scan -or $script:scan.Cancelled) { return }
    $category=[string]$checkbox.Tag
    # Derive the action from the existing selection, not WPF's three-state click cycle.
    $state=Get-UnjunkCategoryState -Rows $script:rows -Category $category -QueueOnReboot:([bool]$script:RebootCheck.IsChecked) -CloseApps:$script:closeApps
    $select=$state.Checked -ne $true
    Set-UnjunkCategorySelection -Rows $script:rows -Category $category -Selected $select -QueueOnReboot:([bool]$script:RebootCheck.IsChecked) -CloseApps:$script:closeApps
    $script:ResultsGrid.Items.Refresh(); Update-UnjunkSelection
    $script:StatusText.Text=if($select){"Selected eligible targets in $category, including search-hidden targets. Review individual checkboxes before cleaning."}else{"Cleared selection in $category."}
    $eventArgs.Handled=$true
})
$script:EmptyCheck.Add_Click({ Update-UnjunkView })
$script:ResultsGrid.Add_SelectionChanged({
    $row=$script:ResultsGrid.SelectedItem
    if($row) {
        $largest=@($row.Files | Sort-Object Length -Descending | Select-Object -First 6 | ForEach-Object { (Format-UnjunkSize $_.Length)+'   '+$_.Path })
        $script:DetailsText.Text="$($row.Root)`r`n$($row.Note)`r`n$($row.FileCount) eligible files; minimum age $($row.MinimumAgeDays) days. Skipped: $($row.SkippedRecent) recent, $($row.SkippedLinks) links, $($row.SkippedExcluded) exclusions.`r`n"+(@($row.Issues)+$largest -join "`r`n")
    }
})
$script:CleanButton.Add_Click({
    $selected=@($script:rows | Where-Object Selected)
    if(-not $script:scan -or -not $selected.Count) { return }
    $total=[long](($selected | Measure-Object Bytes -Sum).Sum)
    $review=@($selected | Where-Object Risk -in @('Review','Diagnostics'))
    $queue=[bool]$script:RebootCheck.IsChecked
    if($queue -and $review.Count) { Show-UnjunkError 'Reboot deletion supports cache and temporary-file targets only. Deselect installer and diagnostic targets.'; return }
    $message="Permanently delete up to $(Format-UnjunkSize $total) from $($selected.Count) selected targets?`r`n`r`n"+(($selected | ForEach-Object {'- '+$_.Name+' ('+$_.Size+')'}) -join "`r`n")
    if($review.Count) { $message+="`r`n`r`nYour selection includes review items (such as backups, old plug-ins, installers or diagnostic history). Confirm you no longer need those files." }
    if($queue) {
        $message="Queue $($selected.Count) targets (up to $(Format-UnjunkSize $total)) for deletion at the next Windows restart?`r`n`r`n"+(($selected | ForEach-Object {'- '+$_.Name+' ('+$_.Size+')'}) -join "`r`n")
        $message+="`r`n`r`nNothing will be deleted now. Windows deletes by path at startup: files changed or replaced at those same paths before restart can also be deleted. Unjunk cannot recheck them then. There is no undo button for queued entries. Turning this option off or cancelling later does not remove entries already queued. Restart Windows when you are ready."
    } elseif($script:closeApps) { $message+="`r`n`r`nMapped apps for selected caches will be closed, then force-stopped if necessary. Unsaved work can be lost. Locked and changed files are skipped. This cannot be undone." }
    else { $message+="`r`n`r`nRunning-app, locked and changed files will be skipped. This cannot be undone." }
    $message+="`r`n`r`nAll selected rows are included, even if hidden by a filter."
    if([Windows.MessageBox]::Show($script:window,$message,'Confirm selected cleanup',[Windows.MessageBoxButton]::YesNo,[Windows.MessageBoxImage]::Warning,[Windows.MessageBoxResult]::No) -eq 'Yes') {
        Start-UnjunkWork Clean @{Scan=$script:scan;Ids=@($selected | Select-Object -ExpandProperty Id);QueueOnReboot=$queue;CloseApps=$script:closeApps}
    }
})
$script:StorageButton.Add_Click({ Start-Process 'ms-settings:storagesense' })
$script:OneClickButton.Add_Click({
    if($script:worker -or $script:oneClickProcess){return}
    try {
        $script:oneClickProcess=Start-UnjunkOneClick
        $script:OneClickButton.IsEnabled=$false
        $script:scan=$null; foreach($row in $script:rows){$row.Selected=$false}; Update-UnjunkView
        $script:StatusText.Text='One-click full cleanup started. You can close Unjunk; the results timeline opens when it finishes.'
    } catch { Show-UnjunkError ('One-click launch cancelled or failed: '+$_.Exception.Message) }
})
$script:HistoryButton.Add_Click({try{$null=Start-UnjunkOneClick -History}catch{Show-UnjunkError $_.Exception.Message}})
$script:WindowsButton.Add_Click({
    do {
        $action=Show-UnjunkWindowsWindow -Owner $script:window -CrowdStrikeSafe $script:crowdStrikeSafe -IsAdmin $script:isAdmin -TempAgeDays $script:tempAge -Exclusions $script:exclusions -CloseApps $script:closeApps
        if($action -eq 'EnableWindowsCleanup') {
            if($script:crowdStrikeSafe) {
                $script:CrowdStrikeCheck.IsChecked=$false
                $script:CrowdStrikeCheck.RaiseEvent((New-Object Windows.RoutedEventArgs ([Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
            }
            if(-not $script:isAdmin) {
                $script:AdminButton.RaiseEvent((New-Object Windows.RoutedEventArgs ([Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
                break
            }
        }
    } while($action -eq 'EnableWindowsCleanup')
    # Maintenance may have changed files or packages that appeared in an earlier file scan.
    $script:scan=$null; foreach($row in $script:rows){$row.Selected=$false}; Update-UnjunkView; Update-UnjunkFreeSpace
})
$script:ReportsButton.Add_Click({ $directory=Join-Path $env:LOCALAPPDATA 'Unjunk\Reports'; $null=New-Item -ItemType Directory -Path $directory -Force; Start-Process explorer.exe -ArgumentList ('"'+$directory+'"') })
$script:WinsxsButton.Add_Click({
    if(-not $script:isAdmin) { $null=[Windows.MessageBox]::Show($script:window,'Use Open as administrator, then Analyze WinSxS. DISM requires administrator access.','Administrator access required'); return }
    Start-UnjunkWork WinSxS $null
})
$script:AdminButton.Add_Click({
    try {
        $launcher=Join-Path $PSScriptRoot 'Unjunk.exe'
        if(Test-Path -LiteralPath $launcher) {
            Start-Process $launcher -Verb RunAs -ErrorAction Stop
        } else {
            $entry=Join-Path $PSScriptRoot 'unjunk.ps1'
            Start-Process "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -Verb RunAs -WindowStyle Hidden -ArgumentList ('-NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "'+$entry+'" -Gui') -ErrorAction Stop
        }
        $script:window.Close()
    } catch { Show-UnjunkError 'Administrator launch was cancelled or could not start.' }
})
$script:ExportButton.Add_Click({
    $dialog=New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter='JSON report (*.json)|*.json'; $dialog.FileName='Unjunk-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'.json'
    if($dialog.ShowDialog($script:window)) {
        try { $null=Save-UnjunkReport -Report $script:latestReport -Path $dialog.FileName; $script:StatusText.Text='Report saved: '+$dialog.FileName }
        catch { Show-UnjunkError $_.Exception.Message }
    }
})
$script:OptionsButton.Add_Click({
    $dialog=New-Object Windows.Window
    $dialog.Title='Unjunk scan options'; $dialog.Width=560; $dialog.Height=420; $dialog.Owner=$script:window; $dialog.WindowStartupLocation='CenterOwner'; $dialog.ResizeMode='NoResize'
    $panel=New-Object Windows.Controls.StackPanel; $panel.Margin='22'; $dialog.Content=$panel
    $label=New-Object Windows.Controls.TextBlock; $label.Text='Keep temporary files newer than this many days (1-365):'; $panel.Children.Add($label) | Out-Null
    $ageBox=New-Object Windows.Controls.TextBox; $ageBox.Text=[string]$script:tempAge; $ageBox.Margin='0,8,0,16'; $panel.Children.Add($ageBox) | Out-Null
    $label=New-Object Windows.Controls.TextBlock; $label.Text='Excluded files or folders (one absolute path per line):'; $panel.Children.Add($label) | Out-Null
    $excludeBox=New-Object Windows.Controls.TextBox; $excludeBox.Text=$script:exclusions -join "`r`n"; $excludeBox.AcceptsReturn=$true; $excludeBox.Height=170; $excludeBox.VerticalScrollBarVisibility='Auto'; $excludeBox.Margin='0,8,0,16'; $panel.Children.Add($excludeBox) | Out-Null
    $save=New-Object Windows.Controls.Button; $save.Content='Save options'; $save.Padding='12,7'; $panel.Children.Add($save) | Out-Null
    $save.Add_Click({
        $days=0
        if(-not [int]::TryParse($ageBox.Text,[ref]$days) -or $days -lt 1 -or $days -gt 365) { $null=[Windows.MessageBox]::Show($dialog,'Enter an age from 1 to 365 days.'); return }
        $paths=@($excludeBox.Text -split '\r?\n' | ForEach-Object {$_.Trim()} | Where-Object {$_})
        foreach($path in $paths) { if($path -notmatch '^[a-zA-Z]:\\' -or $path -match '[*?]') { $null=[Windows.MessageBox]::Show($dialog,'Use absolute paths without wildcards.'); return } }
        try {
            $null=New-Item -ItemType Directory -Path (Split-Path $script:settingsPath) -Force
            @{TempAgeDays=$days;Exclusions=$paths;CrowdStrikeSafe=$script:crowdStrikeSafe;CloseApps=$script:closeApps} | ConvertTo-Json | Set-Content -LiteralPath $script:settingsPath -Encoding UTF8 -ErrorAction Stop
            $script:tempAge=$days; $script:exclusions=$paths; $script:scan=$null
            foreach($row in $script:rows) { $row.Selected=$false }; Update-UnjunkView
            $script:StatusText.Text='Options saved. Scan again to apply them.'; $dialog.Close()
        } catch { $null=[Windows.MessageBox]::Show($dialog,$_.Exception.Message) }
    })
    $null=$dialog.ShowDialog()
})
$script:window.Add_Closing({
    param($sender,$eventArgs)
    if($script:worker) { $eventArgs.Cancel=$true; $script:StatusText.Text='Wait for the operation to finish, or cancel the scan/cleanup first.' }
})
$script:window.Add_Closed({$script:timer.Stop()})
Update-UnjunkFreeSpace
Set-UnjunkBusy $false
$null=$script:window.ShowDialog()
