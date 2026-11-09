#requires -Version 5.1
# Exercise real WPF click callbacks in an isolated script scope, without showing windows
# or running cleanup. The worker module returns synthetic scan data only.
param([string]$ProjectDirectory,[switch]$ConfirmationOnly)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
if(-not $ProjectDirectory){$ProjectDirectory=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))}
$fixtureRoot=Join-Path $env:TEMP ('Unjunk-gui-callback-tests-'+[Guid]::NewGuid().ToString('N'))
$previousLocal=$env:LOCALAPPDATA
$null=New-Item -ItemType Directory -Path $fixtureRoot
$env:LOCALAPPDATA=$fixtureRoot
$script:checks=0
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:checks++;Write-Host "PASS: $Message"}
function Click($Control){
    if(-not $Control.IsEnabled){throw ('Cannot click disabled control: '+$Control.GetType().Name)}
    $type=if($Control -is [Windows.Controls.Primitives.ToggleButton]){[Windows.Controls.Primitives.ToggleButton]}else{[Windows.Controls.Primitives.ButtonBase]}
    $method=$type.GetMethod('OnClick',[Reflection.BindingFlags]'Instance,NonPublic')
    $null=$method.Invoke($Control,@())
}
$fixtureModule=Join-Path $fixtureRoot 'ScanFixture.psm1'
@'
function Invoke-UnjunkWindowsScan {
    param($TempAgeDays,$Exclusions,$CrowdStrikeSafe,$State)
    $mode=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'mode.txt') -Raw
    if($mode.Trim() -eq 'Fail'){throw 'Simulated scan failure'}
    [pscustomobject]@{Cancelled=($mode.Trim() -eq 'Cancel');Apps=@();Shadows=@();Handlers=@('Temporary Files');AppError='';ShadowError=''}
}
function Invoke-UnjunkWindowsCleanup {
    param($Scan,$OptionIds,$AppIds,$State,[switch]$CloseApps,[switch]$Confirm)
    # Simulated return value only. No files, apps, snapshots or native processes are touched.
    [pscustomobject]@{Completed=1;Partial=0;Failed=0;Cancelled=$false;RestartRequired=$false;Results=@([pscustomobject]@{Status='Completed';Name='Simulated operation';Output='Fixture result'})}
}
Export-ModuleMember -Function Invoke-UnjunkWindowsScan,Invoke-UnjunkWindowsCleanup
'@ | Set-Content -LiteralPath $fixtureModule -Encoding UTF8
Set-Content -LiteralPath (Join-Path $fixtureRoot 'mode.txt') -Value 'Success'
$windowsSource=Get-Content -LiteralPath (Join-Path $ProjectDirectory 'Unjunk.Windows.Gui.ps1') -Raw
$windowsSource=$windowsSource.Replace('$PSScriptRoot',("'"+$ProjectDirectory.Replace("'","''")+"'"))
$windowsSource=$windowsSource.Replace('Show-UnjunkWindowsConfirmation','Show-UnjunkWindowsConfirmationScopeFixture')
$windowsSource=$windowsSource.Replace('$dialog.Owner=$Owner;', '$dialog.Tag=$Owner;')
$windowsSource=$windowsSource.Replace('return $dialog.ShowDialog() -eq $true','$global:UnjunkConfirmationFixture=$dialog; return $false')
$windowsSource=$windowsSource.Replace('$timer.Add_Tick({','$poll={')
$windowsSource=[regex]::Replace($windowsSource,'    \}\.GetNewClosure\(\)\)\r?\n    \$ui.Window.Add_Closing','    }.GetNewClosure(); $timer.Add_Tick($poll)'+[Environment]::NewLine+'    $ui.Window.Add_Closing')
$windowsSource=$windowsSource.Replace('try {$null=$ui.Window.ShowDialog()}finally{$timer.Stop()}','$timer.Stop(); $ui.Refresh=$update; $ui.Poll=$poll')
$windowsSource=$windowsSource.Replace('return [string]$ui.Window.Tag','return $ui')
$windowsSource=$windowsSource.Replace('$ui.BusyChecker=Get-Command Test-UnjunkWindowsCleanupBusy -CommandType Function -ErrorAction Stop','$ui.BusyChecker={$false}')
$panels=@()
function New-Panel([bool]$Safe=$false,[bool]$Admin=$true){
    # The helper is deliberately local to this invocation, like the real GUI script.
    . ([scriptblock]::Create($windowsSource))
    Show-UnjunkWindowsWindow -CrowdStrikeSafe $Safe -IsAdmin $Admin
}
function Finish-Scan($Panel){
    $deadline=[DateTime]::UtcNow.AddSeconds(10)
    while($Panel.Worker -and -not $Panel.Worker.Handle.IsCompleted -and [DateTime]::UtcNow -lt $deadline){[Threading.Thread]::Sleep(10)}
    if($Panel.Worker -and -not $Panel.Worker.Handle.IsCompleted){throw 'Fixture scan timed out'}
    & $Panel.Poll
}
try {
    $panel=New-Panel; $panels+=@($panel);$panel.Module=$fixtureModule
    # Only module commands are global; the confirmation helper must survive its local scope.
    Import-Module (Join-Path $ProjectDirectory 'Unjunk.Windows.psm1') -Global
    $panel.BusyChecker={$false}
    Import-Module (Join-Path $ProjectDirectory 'Unjunk.Core.psm1') -Global
    Assert (-not (Get-Command Show-UnjunkWindowsConfirmationScopeFixture -ErrorAction SilentlyContinue)) 'Confirmation helper is absent from the callback command lookup scope'
    if(-not $ConfirmationOnly){
        Assert ($panel.CleanAllCheck.IsEnabled -and $panel.Options[0].Box.IsEnabled -and -not $panel.RunButton.IsEnabled) 'Cleanup options are clickable before scanning; execution still requires a scan'
        Click $panel.CleanAllCheck
        Assert (@($panel.Options | Where-Object {$_.Box.IsChecked}).Count -eq 14 -and -not $panel.RunButton.IsEnabled) 'A real Clean all click selects options before scanning without enabling execution'
        Click $panel.CleanAllCheck
        $option=@($panel.Options | Where-Object Id -eq 'winsxs')[0]
        Click $option.Box
        Assert ($option.Box.IsChecked -and $null -eq $panel.CleanAllCheck.IsChecked) 'An individual checkbox toggles and updates the master checkbox'
    } else { @($panel.Options | Where-Object Id -eq 'winsxs')[0].Box.IsChecked=$true }
    Click $panel.ScanButton
    if(-not $ConfirmationOnly){Assert (-not $panel.OptionsPanel.IsEnabled -and -not $panel.CleanAllCheck.IsEnabled -and -not $panel.RunButton.IsEnabled) 'Selection is locked only while an operation is running'}
    Finish-Scan $panel
    if($ConfirmationOnly){@($panel.Options | Where-Object Id -eq 'winsxs')[0].Box.IsChecked=$true; & $panel.Refresh}
    Assert ($panel.Scan -and @($panel.Options | Where-Object {$_.Box.IsChecked}).Count -eq 1 -and $panel.RunButton.IsEnabled) 'Scan completion preserves the selected option and enables Run selected'
    $global:UnjunkConfirmationFixture=$null
    Click $panel.RunButton
    if(-not $global:UnjunkConfirmationFixture){throw ('Confirmation callback failed: '+$panel.Details.Text)}
    Assert ($null -ne $global:UnjunkConfirmationFixture) 'Run selected reaches the real confirmation builder from its closure'
    $details=@($global:UnjunkConfirmationFixture.Content.Children | Where-Object {$_ -is [Windows.Controls.TextBox]})[0]
    Assert ($details.Text -match 'WinSxS component cleanup' -and $global:UnjunkConfirmationFixture.Tag -eq $panel.Window) 'Confirmation contains the selected operation and receives the correct owner'
    Assert ($null -eq $panel.Worker -and $panel.RunButton.IsEnabled) 'Cancelling confirmation does not start a cleanup worker or discard selection'
    $panel.BusyChecker={$true};& $panel.Refresh
    Assert (-not $panel.RunButton.IsEnabled -and -not $panel.ScanButton.IsEnabled -and $panel.Status.Text -match 'still running') 'An external maintenance run disables Run and Scan with a clear explanation'
    $panel.BusyChecker={$false};& $panel.Refresh
    Assert ($panel.RunButton.IsEnabled -and $panel.ScanButton.IsEnabled) 'The window becomes available after the external run finishes'
    $global:UnjunkConfirmationFixture.Close();$global:UnjunkConfirmationFixture=$null
    if(-not $ConfirmationOnly){
        Set-Content -LiteralPath (Join-Path $fixtureRoot 'mode.txt') -Value 'Fail'
        Click $panel.ScanButton;Finish-Scan $panel
        Assert ($null -eq $panel.Scan -and $panel.CleanAllCheck.IsEnabled -and $option.Box.IsEnabled -and $option.Box.IsChecked -and -not $panel.RunButton.IsEnabled) 'A failed scan leaves cleanup choices editable and preserves the selection'
        Set-Content -LiteralPath (Join-Path $fixtureRoot 'mode.txt') -Value 'Cancel'
        Click $panel.ScanButton;Finish-Scan $panel
        Assert ($panel.Scan.Cancelled -and $panel.CleanAllCheck.IsEnabled -and $option.Box.IsChecked -and -not $panel.RunButton.IsEnabled) 'A cancelled scan leaves choices editable but cannot enable execution'
        Set-Content -LiteralPath (Join-Path $fixtureRoot 'mode.txt') -Value 'Success'
        Click $panel.ScanButton;Finish-Scan $panel
        $panel.ConfirmSelection={param($Owner,$Jobs) $true}
        Click $panel.RunButton;Finish-Scan $panel
        Assert ($null -eq $panel.Worker -and $null -eq $panel.Scan -and $panel.Report.Completed -eq 1 -and $panel.Details.Text -match 'Fixture result') 'Accepting confirmation completes the simulated worker and report callbacks'
        $safePanel=New-Panel -Safe $true; $panels+=@($safePanel)
        Click $safePanel.CleanAllCheck
        Assert ($safePanel.Options[0].Box.IsChecked -and -not $safePanel.ScanButton.IsEnabled -and -not $safePanel.RunButton.IsEnabled) 'Safe mode permits selecting options without permitting Windows operations'
        $standardPanel=New-Panel -Admin $false; $panels+=@($standardPanel)
        Click $standardPanel.CleanAllCheck
        Assert ($standardPanel.Options[0].Box.IsChecked -and -not $standardPanel.ScanButton.IsEnabled -and -not $standardPanel.RunButton.IsEnabled) 'A standard user can select options while administrator execution remains blocked'
    }
    Write-Host "All $script:checks GUI callback checks passed."
} finally {
    foreach($item in $panels){if($item.Worker){$item.Worker.PowerShell.Dispose()};$item.Window.Close()}
    if($global:UnjunkConfirmationFixture){$global:UnjunkConfirmationFixture.Close()}
    Remove-Variable UnjunkConfirmationFixture -Scope Global -ErrorAction SilentlyContinue
    $env:LOCALAPPDATA=$previousLocal
    $resolved=[IO.Path]::GetFullPath($fixtureRoot);$prefix=[IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')+'\Unjunk-gui-callback-tests-'
    if(-not $resolved.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase)){throw 'Fixture cleanup boundary check failed'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
