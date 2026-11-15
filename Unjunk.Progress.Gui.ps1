#requires -Version 5.1
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase

function New-UnjunkProgressWindow($State) {
    [xml]$markup=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 Title="Unjunk | One-click cleanup progress" Width="1120" Height="780" MinWidth="850" MinHeight="580" WindowStartupLocation="CenterScreen"
 Background="#10151C" Foreground="#E8EDF4" FontFamily="Segoe UI" FontSize="13">
 <Window.Resources>
  <Style TargetType="Button"><Setter Property="Padding" Value="16,9"/><Setter Property="Margin" Value="8,0,0,0"/><Setter Property="Background" Value="#273341"/><Setter Property="Foreground" Value="#E8EDF4"/><Setter Property="BorderThickness" Value="0"/></Style>
  <Style TargetType="DataGridColumnHeader"><Setter Property="Background" Value="#202B38"/><Setter Property="Foreground" Value="#A8BDCA"/><Setter Property="Padding" Value="9"/></Style>
  <Style x:Key="CellText" TargetType="TextBlock"><Setter Property="TextWrapping" Value="Wrap"/><Setter Property="Margin" Value="8,6"/><Setter Property="VerticalAlignment" Value="Center"/></Style>
  <Style TargetType="DataGridCell"><Setter Property="BorderThickness" Value="0"/><Style.Triggers><Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#244B58"/><Setter Property="Foreground" Value="#FFFFFF"/></Trigger></Style.Triggers></Style>
  <Style TargetType="DataGridRow"><Setter Property="Background" Value="#161F29"/><Setter Property="Foreground" Value="#E8EDF4"/><Setter Property="MinHeight" Value="40"/>
   <Style.Triggers>
    <DataTrigger Binding="{Binding Status}" Value="Running"><Setter Property="Background" Value="#17413E"/></DataTrigger>
    <DataTrigger Binding="{Binding Status}" Value="Completed"><Setter Property="Foreground" Value="#80E5CE"/></DataTrigger>
    <DataTrigger Binding="{Binding Status}" Value="Skipped"><Setter Property="Foreground" Value="#F3CB82"/></DataTrigger>
    <DataTrigger Binding="{Binding Status}" Value="Partial"><Setter Property="Foreground" Value="#F3CB82"/></DataTrigger>
    <DataTrigger Binding="{Binding Status}" Value="Failed"><Setter Property="Foreground" Value="#FF9999"/></DataTrigger>
   </Style.Triggers>
  </Style>
 </Window.Resources>
 <Grid Margin="24">
  <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="150"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
  <StackPanel Margin="0,0,0,18">
   <TextBlock Text="One-click cleanup" FontSize="28" FontWeight="SemiBold"/>
   <TextBlock x:Name="Current" Text="Preparing cleanup..." TextWrapping="Wrap" Foreground="#80E5CE" FontSize="15" Margin="0,12,0,8"/>
   <TextBlock x:Name="Summary" Text="Scanning targets. The ordered checklist will appear here." Foreground="#A8BDCA"/>
   <ProgressBar x:Name="Progress" Height="5" Margin="0,12,0,0" Minimum="0" Maximum="100" IsIndeterminate="True"/>
  </StackPanel>
  <DataGrid x:Name="Stages" Grid.Row="1" IsReadOnly="True" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" HeadersVisibility="Column" RowHeaderWidth="0" GridLinesVisibility="None" Background="#161F29" BorderBrush="#283747" SelectionMode="Single">
   <DataGrid.Columns>
    <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100" ElementStyle="{StaticResource CellText}"/>
    <DataGridTextColumn Header="Cleanup stage" Binding="{Binding Name}" Width="2*" MinWidth="220" ElementStyle="{StaticResource CellText}"/>
    <DataGridTextColumn Header="Items cleaned" Binding="{Binding ItemsText}" Width="130" ElementStyle="{StaticResource CellText}"/>
    <DataGridTextColumn Header="File bytes freed" Binding="{Binding SizeText}" Width="120" ElementStyle="{StaticResource CellText}"/>
    <DataGridTextColumn Header="Result so far" Binding="{Binding Summary}" Width="2*" MinWidth="180" ElementStyle="{StaticResource CellText}"/>
   </DataGrid.Columns>
  </DataGrid>
  <Border Grid.Row="2" Background="#19232E" Padding="14" Margin="0,14,0,14" CornerRadius="5"><TextBox x:Name="Details" Text="Select a stage to read its result or the reason it was skipped. Windows-managed operations do not report exact item counts." IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Background="Transparent" Foreground="#C2D0DD" BorderThickness="0"/></Border>
  <DockPanel Grid.Row="3"><StackPanel DockPanel.Dock="Right" Orientation="Horizontal"><Button x:Name="Stop" Content="Stop after current stage"/><Button x:Name="Done" Content="View cleanup timeline" Background="#54D0B3" Foreground="#102A25" IsEnabled="False"/></StackPanel><TextBlock Text="Minimize to keep cleaning. Stop preserves completed work." Foreground="#95A7BB" VerticalAlignment="Center" TextWrapping="Wrap" Margin="0,0,12,0"/></DockPanel>
 </Grid>
</Window>
'@
    $ui=@{Window=[Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $markup));State=$State;Revision='';Complete=$false;Run=$null;Error=''}
    foreach($name in @('Current','Summary','Progress','Stages','Details','Stop','Done')){$ui[$name]=$ui.Window.FindName($name)}
    $ui.Refresh={
        $ui.Current.Text=[string]$ui.State.Message
        $snapshot=if($ui.State.ContainsKey('StepSnapshot')){$ui.State.StepSnapshot}else{$null}
        if($null -ne $snapshot -and $snapshot.Revision -ne $ui.Revision) {
            $selected=if($ui.Stages.SelectedItem){$ui.Stages.SelectedItem.Id}else{''}
            $ui.Revision=$snapshot.Revision;$ui.Stages.ItemsSource=@($snapshot.Steps);$ui.Summary.Text=$snapshot.Summary
            $ui.Progress.IsIndeterminate=$snapshot.Total -eq 0 -and -not $ui.Complete
            if($snapshot.Total){$ui.Progress.Value=100.0*$snapshot.FinishedCount/$snapshot.Total}
            if($selected){$ui.Stages.SelectedItem=@($snapshot.Steps|Where-Object Id -eq $selected)|Select-Object -First 1}
        }
        if($ui.State.Cancel){$ui.Stop.IsEnabled=$false}
        if($ui.Complete){$ui.Stop.IsEnabled=$false;$ui.Done.IsEnabled=$true;$ui.Progress.IsIndeterminate=$false}
    }.GetNewClosure()
    $ui.Stages.Add_SelectionChanged({if($ui.Stages.SelectedItem){$ui.Details.Text=$ui.Stages.SelectedItem.Name+"`r`n"+$ui.Stages.SelectedItem.Detail}}.GetNewClosure())
    $ui.Stop.Add_Click({$ui.State.Cancel=$true;$ui.Stop.IsEnabled=$false;$ui.Current.Text='Stopping after the current operation finishes. Completed work is kept.'}.GetNewClosure())
    $ui.Done.Add_Click({$ui.Window.Close()}.GetNewClosure())
    $ui.Window.Add_Closing({param($sender,$eventArgs)if(-not $ui.Complete){$eventArgs.Cancel=$true;$ui.Window.WindowState='Minimized'}}.GetNewClosure())
    return $ui
}

function Show-UnjunkOneClickProgress {
    param([string[]]$Exclusions=@(),[bool]$CloseApps=$true,[string]$ModulePath=(Join-Path $PSScriptRoot 'Unjunk.OneClick.psm1'))
    $state=[hashtable]::Synchronized(@{Cancel=$false;Progress=0;Message='Scanning full cleanup targets...';Finished=$false})
    $ui=New-UnjunkProgressWindow $state
    $runspace=[RunspaceFactory]::CreateRunspace();$runspace.ApartmentState='STA';$runspace.ThreadOptions='ReuseThread';$runspace.Open()
    $worker=[PowerShell]::Create();$worker.Runspace=$runspace
    $null=$worker.AddScript({param($module,$exclusions,$closeApps,$state)
        $ErrorActionPreference='Stop';Import-Module $module -Force
        Invoke-UnjunkOneClick -Exclusions $exclusions -CloseApps $closeApps -ProgressState $state
    }).AddArgument($ModulePath).AddArgument($Exclusions).AddArgument($CloseApps).AddArgument($state)
    $timer=New-Object Windows.Threading.DispatcherTimer;$timer.Interval=[TimeSpan]::FromMilliseconds(250)
    $handle=$null
    try {
        $handle=$worker.BeginInvoke()
        $timer.Add_Tick({
            & $ui.Refresh
            if($ui.Complete -or -not $handle.IsCompleted){return}
            try {
                $result=@($worker.EndInvoke($handle))
                if($worker.HadErrors){throw ($worker.Streams.Error|Out-String)}
                if($result.Count -ne 1){throw 'The cleanup worker did not return a complete result.'}
                $ui.Run=$result[0]
            } catch {
                $ui.Error=$_.Exception.GetBaseException().Message;$ui.State.Message='Cleanup stopped: '+$ui.Error;$ui.Details.Text=$ui.Error
            }
            $ui.Complete=$true;& $ui.Refresh
        }.GetNewClosure())
        $timer.Start();$null=$ui.Window.ShowDialog()
    } finally {$timer.Stop();$worker.Dispose();$runspace.Dispose()}
    if($ui.Error){throw $ui.Error}
    return $ui.Run
}
