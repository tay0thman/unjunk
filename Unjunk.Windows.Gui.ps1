#requires -Version 5.1
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Core.psm1')
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Windows.psm1') -Force

function Show-UnjunkWindowsConfirmation($Owner,[object[]]$Jobs,[bool]$CloseApps=$false) {
    $dialog=New-Object Windows.Window
    $dialog.Title='Review Windows cleanup'; $dialog.Width=850; $dialog.Height=670; $dialog.MinHeight=400
    $dialog.Owner=$Owner; $dialog.WindowStartupLocation='CenterOwner'; $dialog.Background='#10151C'; $dialog.Foreground='#E8EDF4'
    $panel=New-Object Windows.Controls.DockPanel; $panel.Margin='22'; $dialog.Content=$panel
    $heading=New-Object Windows.Controls.TextBlock; $heading.Text="Run $($Jobs.Count) selected operations?"; $heading.FontSize=22; $heading.Margin='0,0,0,12'
    [Windows.Controls.DockPanel]::SetDock($heading,'Top'); $null=$panel.Children.Add($heading)
    $buttons=New-Object Windows.Controls.StackPanel; $buttons.Orientation='Horizontal'; $buttons.HorizontalAlignment='Right'; $buttons.Margin='0,14,0,0'
    [Windows.Controls.DockPanel]::SetDock($buttons,'Bottom'); $null=$panel.Children.Add($buttons)
    $cancel=New-Object Windows.Controls.Button; $cancel.Content='Cancel'; $cancel.Padding='18,9'; $cancel.Margin='0,0,12,0'; $cancel.IsCancel=$true; $null=$buttons.Children.Add($cancel)
    $run=New-Object Windows.Controls.Button; $run.Content='Run selected'; $run.Padding='18,9'; $null=$buttons.Children.Add($run)
    $details=New-Object Windows.Controls.TextBox; $details.IsReadOnly=$true; $details.TextWrapping='Wrap'; $details.VerticalScrollBarVisibility='Auto'
    $details.Background='#19232E'; $details.Foreground='#E8EDF4'; $details.Padding='14'; $details.BorderThickness=0
    $details.Text="File deletion, removal of unused provisioning, and history clearing are permanent. Review all selected operations below. Windows operations run without additional Unjunk prompts. Stop waits for the current Windows operation to finish. Windows is not restarted automatically.`r`n`r`n"+(($Jobs | ForEach-Object {$_.Name+"`r`n"+$_.Detail}) -join "`r`n`r`n")
    $null=$panel.Children.Add($details)
    if($CloseApps -and @($Jobs | Where-Object Id -eq 'app-caches').Count){$details.Text="Selected cache apps and IDM will be closed, then force-stopped if necessary. Unsaved work can be lost. Save your work first.`r`n`r`n"+$details.Text}
    $run.Add_Click({$dialog.DialogResult=$true}.GetNewClosure())
    return $dialog.ShowDialog() -eq $true
}

function Show-UnjunkWindowsWindow {
    param($Owner,[bool]$CrowdStrikeSafe,[bool]$IsAdmin,[int]$TempAgeDays=7,[string[]]$Exclusions=@(),[bool]$CloseApps=$true)
    [xml]$markup=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 Title="Unjunk | Windows cleanup and apps" Width="1240" Height="860" MinWidth="1000" MinHeight="700" WindowStartupLocation="CenterOwner"
 Background="#10151C" Foreground="#E8EDF4" FontFamily="Segoe UI" FontSize="13">
 <Window.Resources>
  <Style TargetType="Button"><Setter Property="Padding" Value="12,8"/><Setter Property="Margin" Value="0,0,8,0"/><Setter Property="Background" Value="#273341"/><Setter Property="Foreground" Value="#F1F5F9"/><Setter Property="BorderThickness" Value="0"/><Setter Property="Cursor" Value="Hand"/></Style>
  <Style TargetType="CheckBox"><Setter Property="Foreground" Value="#E8EDF4"/></Style>
  <Style TargetType="DataGridColumnHeader"><Setter Property="Background" Value="#202B38"/><Setter Property="Foreground" Value="#B8C8D9"/><Setter Property="Padding" Value="8"/></Style>
  <Style TargetType="DataGridRow"><Setter Property="Background" Value="#161F29"/><Setter Property="Foreground" Value="#E8EDF4"/><Setter Property="MinHeight" Value="34"/></Style>
 </Window.Resources>
 <Grid Margin="24">
  <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="125"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
  <StackPanel><TextBlock Text="Windows cleanup &amp; app leftovers" FontSize="28" FontWeight="SemiBold"/><TextBlock Text="Clean junk and remnants of uninstalled apps. Installed apps are preserved." Foreground="#95A7BB" Margin="0,5,0,14"/></StackPanel>
  <Border Grid.Row="1" Background="#193834" CornerRadius="6" Padding="14" Margin="0,0,0,16"><DockPanel><Button x:Name="EnableButton" DockPanel.Dock="Right" Content="Enable Windows cleanup"/><TextBlock x:Name="ModeNotice" TextWrapping="Wrap" VerticalAlignment="Center" Foreground="#B3DCD2"/></DockPanel></Border>
  <Grid Grid.Row="2"><Grid.ColumnDefinitions><ColumnDefinition Width="440"/><ColumnDefinition Width="18"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
   <DockPanel>
    <StackPanel DockPanel.Dock="Top" Margin="0,0,0,12"><CheckBox x:Name="CleanAllCheck" FontSize="19" FontWeight="SemiBold" Content="Clean all" IsThreeState="True"/><TextBlock Text="Selects all file and system cleanup options below, including shadow copies. Machine-wide app leftovers are selected separately." Foreground="#95A7BB" TextWrapping="Wrap" Margin="24,6,8,0"/></StackPanel>
    <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="OptionsPanel" Margin="0,0,10,0"/></ScrollViewer>
   </DockPanel>
   <DockPanel Grid.Column="2">
    <StackPanel DockPanel.Dock="Top"><TextBlock Text="UNINSTALLED APP LEFTOVERS" FontSize="12" Foreground="#95A7BB"/><TextBlock Text="Machine-wide: unused provisioning and leftover data across local profiles. Apps registered to any account are preserved." TextWrapping="Wrap" Margin="0,6,0,10"/>
     <DockPanel Margin="0,0,0,10"><StackPanel DockPanel.Dock="Right" Orientation="Horizontal"><Button x:Name="AllAppsButton" Content="Select all leftovers"/><Button x:Name="ClearAppsButton" Content="Clear leftovers" Margin="0"/></StackPanel><TextBox x:Name="AppSearch" Margin="0,0,10,0" Padding="8,5" Background="#202B38" Foreground="#E8EDF4" ToolTip="Filter apps by package name"/></DockPanel>
    </StackPanel>
    <DataGrid x:Name="AppsGrid" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" CanUserResizeRows="False" HeadersVisibility="Column" RowHeaderWidth="0" Background="#161F29" BorderBrush="#283747" GridLinesVisibility="None" SelectionMode="Single">
     <DataGrid.Columns><DataGridTemplateColumn Width="35"><DataGridTemplateColumn.CellTemplate><DataTemplate><CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" VerticalAlignment="Center"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTextColumn Header="Removed app" Binding="{Binding Name}" Width="*" MinWidth="200" IsReadOnly="True"/><DataGridTextColumn Header="Leftover type" Binding="{Binding Version}" Width="115" IsReadOnly="True"/><DataGridTextColumn Header="Applies to" Binding="{Binding Scope}" Width="130" IsReadOnly="True"/></DataGrid.Columns>
    </DataGrid>
   </DockPanel>
  </Grid>
  <Border Grid.Row="3" Background="#19232E" CornerRadius="5" Padding="12" Margin="0,14,0,12"><TextBox x:Name="Details" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Background="Transparent" Foreground="#B8C8D9" BorderThickness="0" Text="Scan Windows to inventory files, machine-wide app leftovers, Disk Cleanup categories and shadow copies. No cleanup runs during the scan."/></Border>
  <Grid Grid.Row="4"><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><StackPanel Margin="0,0,16,0"><TextBlock x:Name="Status" Text="Ready to scan." TextWrapping="Wrap" Foreground="#A7B9CD"/><ProgressBar x:Name="Progress" Height="4" Margin="0,8,0,0" Minimum="0" Maximum="100"/></StackPanel><StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Top"><Button x:Name="ClearButton" Content="Clear selection"/><Button x:Name="ExportButton" Content="Export report" IsEnabled="False"/><Button x:Name="StopButton" Content="Stop" IsEnabled="False"/><Button x:Name="ScanButton" Content="Scan Windows"/><Button x:Name="RunButton" Content="Run selected" Background="#54D0B3" Foreground="#102A25" IsEnabled="False" Margin="0"/></StackPanel></Grid>
 </Grid>
</Window>
'@
    $ui=@{Window=[Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $markup));Scan=$null;Worker=$null;Report=$null;
        Options=@();Apps=@();CanRun=($IsAdmin -and -not $CrowdStrikeSafe);Safe=$CrowdStrikeSafe;TempAge=$TempAgeDays;Exclusions=$Exclusions;
        Module=(Join-Path $PSScriptRoot 'Unjunk.Windows.psm1');Syncing=$false;LastSummary=''}
    # GetNewClosure callbacks run in dynamic modules. Capture script-local helpers and
    # module commands now instead of resolving their names from a later click callback.
    $ui.ConfirmSelection=${function:Show-UnjunkWindowsConfirmation}
    $ui.BuildPlan=Get-Command Get-UnjunkWindowsPlan -CommandType Function -ErrorAction Stop
    $ui.BusyChecker=Get-Command Test-UnjunkWindowsCleanupBusy -CommandType Function -ErrorAction Stop
    $ui.WriteReport=Get-Command Save-UnjunkReport -CommandType Function -ErrorAction Stop
    if($Owner) { $ui.Window.Owner=$Owner }
    foreach($name in @('EnableButton','ModeNotice','CleanAllCheck','OptionsPanel','AllAppsButton','ClearAppsButton','AppSearch','AppsGrid','Details','Status','Progress','ClearButton','ExportButton','StopButton','ScanButton','RunButton')) { $ui[$name]=$ui.Window.FindName($name) }
    $ui.EnableButton.Visibility=if($ui.CanRun){'Collapsed'}else{'Visible'}
    $ui.CloseAppsCheck=New-Object Windows.Controls.CheckBox
    $ui.CloseAppsCheck.Content='Close apps before cleanup';$ui.CloseAppsCheck.IsChecked=$CloseApps;$ui.CloseAppsCheck.Margin='0,0,0,16'
    $ui.CloseAppsCheck.ToolTip='Requests a normal close, then stops mapped cache apps and IDM still running after five seconds. Unsaved work can be lost. System, security and installer processes are protected. C:\Autodesk cleanup also stops the Autodesk Access updater service.'
    $null=$ui.OptionsPanel.Children.Add($ui.CloseAppsCheck)
    $ui.ModeNotice.Text=if($ui.CanRun){'Choose your options, then Scan Windows before running. Clean all includes restore points and the full Windows Disk Cleanup profile.'}else{'You can choose cleanup options now. Scanning and running require Standard cleaning and administrator access. Use Enable Windows cleanup to switch profile and request administrator access.'}
    foreach($option in @(Get-UnjunkWindowsOptions)) {
        $box=New-Object Windows.Controls.CheckBox; $box.Tag=$option.Id; $box.Margin='0,0,0,14'
        $stack=New-Object Windows.Controls.StackPanel
        $label=New-Object Windows.Controls.TextBlock; $label.Text=$option.Name; $label.FontWeight='SemiBold'; $label.TextWrapping='Wrap'; $null=$stack.Children.Add($label)
        $note=New-Object Windows.Controls.TextBlock; $note.Text=$option.Note; $note.TextWrapping='Wrap'; $note.Foreground='#95A7BB'; $note.FontSize=12; $note.Margin='0,4,8,0'; $null=$stack.Children.Add($note)
        $box.Content=$stack; $null=$ui.OptionsPanel.Children.Add($box)
        $ui.Options+=@([pscustomobject]@{Id=$option.Id;Box=$box})
    }
    $update={
        $busy=$null -ne $ui.Worker; $outsideBusy=(-not $busy -and (& $ui.BusyChecker))
        $ready=$ui.CanRun -and $null -ne $ui.Scan -and -not $ui.Scan.Cancelled -and -not $busy -and -not $outsideBusy
        $selectedOptions=@($ui.Options | Where-Object {$_.Box.IsChecked -eq $true})
        $appCount=@($ui.Apps | Where-Object Selected).Count
        $ui.Syncing=$true
        $ui.CleanAllCheck.IsChecked=if($selectedOptions.Count -eq $ui.Options.Count){$true}elseif($selectedOptions.Count){$null}else{$false}
        $ui.Syncing=$false
        $ui.RunButton.IsEnabled=$ready -and ($selectedOptions.Count+$appCount -gt 0)
        $ui.ScanButton.IsEnabled=$ui.CanRun -and -not $busy -and -not $outsideBusy
        # Choosing options is harmless and must not depend on a completed scan.
        foreach($control in @($ui.OptionsPanel,$ui.CleanAllCheck,$ui.ClearButton)) { $control.IsEnabled=-not $busy }
        foreach($control in @($ui.AllAppsButton,$ui.ClearAppsButton,$ui.AppsGrid)) { $control.IsEnabled=$ready }
        $ui.EnableButton.IsEnabled=-not $busy
        $ui.RunButton.ToolTip=if($busy -or $outsideBusy){'Another cleanup is active. Wait for it to finish; one-click already includes DISM and Disk Cleanup.'}elseif(-not $ui.CanRun){'Enable Windows cleanup, then scan before running.'}elseif(-not $ready){'Scan Windows first. Your cleanup selections will be kept.'}else{'Review and run the selected operations.'}
        $ui.StopButton.IsEnabled=$busy -and -not $ui.State.Cancel
        $summary="$($selectedOptions.Count) cleanup options / $appCount app leftovers selected."
        if($outsideBusy){$ui.Status.Text='Another Unjunk Windows cleanup is still running. This window will be ready when it finishes.'; $ui.LastSummary=''}
        if($ready -and $summary -ne $ui.LastSummary) { $ui.Status.Text=$summary; $ui.LastSummary=$summary }
    }.GetNewClosure()
    $start={param($operation,$payload)
        if($ui.Worker) { return }
        $ui.State=[hashtable]::Synchronized(@{Cancel=$false;Progress=0;Message='Starting...'})
        $ps=[PowerShell]::Create()
        $runspace=[RunspaceFactory]::CreateRunspace();$runspace.ApartmentState='STA';$runspace.ThreadOptions='ReuseThread';$runspace.Open();$ps.Runspace=$runspace
        $null=$ps.AddScript({param($module,$operation,$payload,$state)
            $ErrorActionPreference='Stop'; Import-Module $module -Force
            if($operation -eq 'Scan') { Invoke-UnjunkWindowsScan -TempAgeDays $payload.TempAge -Exclusions $payload.Exclusions -CrowdStrikeSafe:$false -State $state }
            else { Invoke-UnjunkWindowsCleanup -Scan $payload.Scan -OptionIds $payload.OptionIds -AppIds $payload.AppIds -CloseApps:$payload.CloseApps -State $state -Confirm:$false }
        }).AddArgument($ui.Module).AddArgument($operation).AddArgument($payload).AddArgument($ui.State)
        try{$ui.Worker=@{PowerShell=$ps;Runspace=$runspace;Handle=$ps.BeginInvoke();Operation=$operation}}
        catch{$ps.Dispose();$runspace.Dispose();throw}
        $ui.Progress.IsIndeterminate=$true; & $update
    }.GetNewClosure()
    $ui.ScanButton.Add_Click({ & $start 'Scan' @{TempAge=$ui.TempAge;Exclusions=$ui.Exclusions} }.GetNewClosure())
    $ui.EnableButton.Add_Click({$ui.Window.Tag='EnableWindowsCleanup';$ui.Window.Close()}.GetNewClosure())
    $ui.CleanAllCheck.Add_Click({
        if($ui.Syncing) { return }
        # Derive the toggle from the previous selection, independent of the three-state click cycle.
        $select=@($ui.Options | Where-Object {$_.Box.IsChecked -eq $true}).Count -ne $ui.Options.Count
        foreach($option in $ui.Options) { $option.Box.IsChecked=$select }; & $update
    }.GetNewClosure())
    foreach($option in $ui.Options) { $option.Box.Add_Click($update) }
    $ui.AllAppsButton.Add_Click({foreach($app in $ui.Apps){$app.Selected=$true};$ui.AppsGrid.Items.Refresh();& $update}.GetNewClosure())
    $ui.ClearAppsButton.Add_Click({foreach($app in $ui.Apps){$app.Selected=$false};$ui.AppsGrid.Items.Refresh();& $update}.GetNewClosure())
    $ui.ClearButton.Add_Click({foreach($option in $ui.Options){$option.Box.IsChecked=$false};foreach($app in $ui.Apps){$app.Selected=$false};$ui.AppsGrid.Items.Refresh();& $update}.GetNewClosure())
    $ui.AppSearch.Add_TextChanged({$query=$ui.AppSearch.Text.Trim();$ui.AppsGrid.ItemsSource=@($ui.Apps | Where-Object {-not $query -or $_.Name.IndexOf($query,[StringComparison]::OrdinalIgnoreCase) -ge 0})}.GetNewClosure())
    $ui.RunButton.Add_Click({
        try {
            if(-not $ui.RunButton.IsEnabled) { return }
            $ids=@($ui.Options | Where-Object {$_.Box.IsChecked -eq $true} | Select-Object -ExpandProperty Id)
            $appIds=@($ui.Apps | Where-Object Selected | Select-Object -ExpandProperty Id)
            $plan=@(& $ui.BuildPlan -Scan $ui.Scan -OptionIds $ids -AppIds $appIds)
            if(& $ui.ConfirmSelection -Owner $ui.Window -Jobs $plan -CloseApps:([bool]$ui.CloseAppsCheck.IsChecked)) { & $start 'Clean' @{Scan=$ui.Scan;OptionIds=$ids;AppIds=$appIds;CloseApps=[bool]$ui.CloseAppsCheck.IsChecked} }
        } catch {$ui.Details.Text=$_.Exception.Message;$ui.Status.Text='Selection could not run. See details.'}
    }.GetNewClosure())
    $ui.StopButton.Add_Click({$ui.State.Cancel=$true;$ui.StopButton.IsEnabled=$false;$ui.Status.Text='Stopping after the current Windows operation finishes...'}.GetNewClosure())
    $ui.ExportButton.Add_Click({
        $save=New-Object Microsoft.Win32.SaveFileDialog; $save.Filter='JSON report (*.json)|*.json';$save.FileName='Unjunk-Windows-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'.json'
        if($save.ShowDialog($ui.Window)) {try {$null=& $ui.WriteReport -Report $ui.Report -Path $save.FileName;$ui.Status.Text='Report exported.'}catch{$ui.Details.Text=$_.Exception.Message}}
    }.GetNewClosure())
    $timer=New-Object Windows.Threading.DispatcherTimer; $timer.Interval=[TimeSpan]::FromMilliseconds(250)
    $timer.Add_Tick({
        if(-not $ui.Worker) { & $update; return }
        $ui.Status.Text=$ui.State.Message
        if(-not $ui.Worker.Handle.IsCompleted) { return }
        $operation=$ui.Worker.Operation
        try {
            $result=@($ui.Worker.PowerShell.EndInvoke($ui.Worker.Handle))
            if($ui.Worker.PowerShell.HadErrors) { throw ($ui.Worker.PowerShell.Streams.Error | Out-String) }
            if($result.Count -ne 1) { throw 'The worker returned an incomplete result.' }
            $ui.Report=$result[0]; $ui.ExportButton.IsEnabled=$true
            if($operation -eq 'Scan') {
                $ui.Scan=$result[0];$ui.Apps=@($ui.Scan.Apps);$ui.AppsGrid.ItemsSource=$ui.Apps
                # Keep the user's cleanup choices when a scan completes or is refreshed.
                $ui.Details.Text="Scan complete. $($ui.Apps.Count) machine-wide app leftovers; $(@($ui.Scan.Shadows).Count) shadow copies. Registered apps are preserved.`r`n`r`nFull Disk Cleanup categories: "+($ui.Scan.Handlers -join ', ')
                if($ui.Scan.AppError) {$ui.Details.Text+="`r`nApp inventory failed: "+$ui.Scan.AppError}
                if($ui.Scan.ShadowError) {$ui.Details.Text+="`r`nShadow-copy inventory failed: "+$ui.Scan.ShadowError}
                if($null -ne $ui.Scan.PSObject.Properties['PrivacyError'] -and $ui.Scan.PrivacyError){$ui.Details.Text+="`r`nPrivacy inventory failed: "+$ui.Scan.PrivacyError}
                if($null -ne $ui.Scan.PSObject.Properties['ExplorerHistoryError'] -and $ui.Scan.ExplorerHistoryError){$ui.Details.Text+="`r`nExplorer Home inventory failed: "+$ui.Scan.ExplorerHistoryError}
                if($null -ne $ui.Scan.PSObject.Properties['EventLogError'] -and $ui.Scan.EventLogError){$ui.Details.Text+="`r`nEvent Viewer inventory failed: "+$ui.Scan.EventLogError}
                $ui.LastSummary='';$ui.Status.Text=if($ui.Scan.Cancelled){'Scan cancelled. Scan again before running.'}else{'Scan complete. Choose your cleanup options.'}
            } else {
                $ui.Scan=$null
                $ui.Status.Text="$($ui.Report.Completed) completed / $($ui.Report.Partial) partial / $($ui.Report.Failed) failed. Scan again to refresh."
                if($ui.Report.Cancelled){$ui.Status.Text='Stopped. '+$ui.Status.Text}
                if($ui.Report.RestartRequired){$ui.Status.Text+=' Windows reports that a restart is needed.'}
                $ui.Details.Text=$ui.Status.Text+"`r`n`r`n"+(($ui.Report.Results | ForEach-Object {$_.Status+' | '+$_.Name+"`r`n"+$_.Output}) -join "`r`n`r`n")
                foreach($option in $ui.Options){$option.Box.IsChecked=$false};foreach($app in $ui.Apps){$app.Selected=$false};$ui.AppsGrid.Items.Refresh()
            }
            try {$path=& $ui.WriteReport -Report $ui.Report -Kind ('windows-'+$operation.ToLowerInvariant());$ui.Details.Text+="`r`n`r`nReport: "+$path}catch{$ui.Details.Text+="`r`nCould not save report: "+$_.Exception.Message}
        } catch {
            $ui.Scan=$null; $ui.Status.Text='Operation failed. See details and scan again.';$ui.Details.Text=$_.Exception.GetBaseException().Message
        } finally {
            $ui.Worker.PowerShell.Dispose();$ui.Worker.Runspace.Dispose();$ui.Worker=$null;$ui.Progress.IsIndeterminate=$false;& $update
        }
    }.GetNewClosure())
    $ui.Window.Add_Closing({param($sender,$eventArgs) if($ui.Worker){$eventArgs.Cancel=$true;$ui.Status.Text='Wait for the current operation to finish. Stop prevents later operations.'}}.GetNewClosure())
    $ui.State=@{Cancel=$false}; & $update; $timer.Start()
    try {$null=$ui.Window.ShowDialog()}finally{$timer.Stop()}
    return [string]$ui.Window.Tag
}
