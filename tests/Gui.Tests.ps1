#requires -Version 5.1
# Construct and render WPF controls offscreen. Never show a window or invoke cleanup.
param([string]$RenderDirectory)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
$project=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$windowsSource=Get-Content -LiteralPath (Join-Path $project 'Unjunk.Windows.Gui.ps1') -Raw
$windowsSource=$windowsSource.Replace('$PSScriptRoot',("'"+$project.Replace("'","''")+"'"))
$windowsSource=$windowsSource.Replace('try {$null=$ui.Window.ShowDialog()}finally{$timer.Stop()}','$timer.Stop()')
$windowsSource=$windowsSource.Replace('return [string]$ui.Window.Tag','return $ui')
$windowsSource=$windowsSource.Replace('$ui.BusyChecker=Get-Command Test-UnjunkWindowsCleanupBusy -CommandType Function -ErrorAction Stop','$ui.BusyChecker={$false}')
. ([scriptblock]::Create($windowsSource))
$ui=Show-UnjunkWindowsWindow -CrowdStrikeSafe $false -IsAdmin $true
if($ui.Options.Count -ne 14 -or $ui.RunButton.IsEnabled -or -not $ui.ScanButton.IsEnabled){throw 'Windows panel initial state is invalid'}
if($ui.CleanAllCheck.Content -ne 'Clean all' -or $ui.AllAppsButton.Content -ne 'Select all leftovers'){throw 'Selection controls are missing'}
if($ui.Options.Id -notcontains 'shadow-copies'){throw 'Shadow-copy option is missing'}
Write-Host 'PASS: Windows panel constructs with fourteen options, Clean all and separate app selection'
$ui.Scan=[pscustomobject]@{Cancelled=$false}
$ui.Apps=@([pscustomobject]@{Name='Example.App';Version='1.0';Scope='All existing users';Selected=$false},[pscustomobject]@{Name='Hidden.App';Version='1.0';Scope='New user accounts';Selected=$false})
$ui.AppsGrid.ItemsSource=$ui.Apps
$ui.CleanAllCheck.RaiseEvent((New-Object Windows.RoutedEventArgs ([Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
if(@($ui.Options | Where-Object {$_.Box.IsChecked}).Count -ne 14 -or @($ui.Apps | Where-Object Selected).Count){throw 'Clean all must select every cleanup option and no apps'}
Write-Host 'PASS: Clean all selects every cleanup option without selecting apps'
$ui.CleanAllCheck.RaiseEvent((New-Object Windows.RoutedEventArgs ([Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
if(@($ui.Options | Where-Object {$_.Box.IsChecked}).Count){throw 'Second Clean all click must clear the options'}
Write-Host 'PASS: Clean all toggles off in one click'
$ui.AppSearch.Text='Example'
$ui.AllAppsButton.RaiseEvent((New-Object Windows.RoutedEventArgs ([Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
if(@($ui.Apps | Where-Object Selected).Count -ne 2){throw 'Select all leftovers must include search-hidden rows'}
Write-Host 'PASS: Select all leftovers includes search-hidden rows'
$ui.ClearButton.RaiseEvent((New-Object Windows.RoutedEventArgs ([Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
if(@($ui.Apps | Where-Object Selected).Count -or $ui.RunButton.IsEnabled){throw 'Clearing selection must disable execution'}
Write-Host 'PASS: Clear selection disables execution and clears app choices'
$ui.AppSearch.Text=''
$safe=Show-UnjunkWindowsWindow -CrowdStrikeSafe $true -IsAdmin $true
if($safe.ScanButton.IsEnabled -or $safe.RunButton.IsEnabled -or $safe.EnableButton.Visibility -ne 'Visible'){throw 'Safe-profile gating failed'}
Write-Host 'PASS: Safe profile exposes the opt-in route and disables native operations'
$nonAdmin=Show-UnjunkWindowsWindow -CrowdStrikeSafe $false -IsAdmin $false
if($nonAdmin.ScanButton.IsEnabled -or $nonAdmin.RunButton.IsEnabled){throw 'Non-admin gating failed'}
Write-Host 'PASS: Non-administrator panel cannot scan or run all-user maintenance'
$mainSource=Get-Content -LiteralPath (Join-Path $project 'Unjunk.Gui.ps1') -Raw
$mainSource=$mainSource.Replace('$PSScriptRoot',("'"+$project.Replace("'","''")+"'"))
$mainSource=$mainSource.Replace('$null=$script:window.ShowDialog()','$script:timer.Stop()')
$mainSource=$mainSource.Replace('$script:maintenanceBusy=Test-UnjunkWindowsCleanupBusy','$script:maintenanceBusy=$false')
. ([scriptblock]::Create($mainSource))
if(-not $script:window.FindName('WindowsButton') -or -not $script:window.FindName('SelectAllButton')){throw 'Main window integration is missing'}
if(-not $script:OneClickButton.IsEnabled -or $script:OneClickButton.Content -ne 'One-click deep clean' -or -not $script:HistoryButton.IsEnabled){throw 'One-click deep clean and timeline controls must be available in Safe mode'}
Write-Host 'PASS: Main window constructs with Windows cleanup and Select all controls'
function Render($Window,[string]$Name) {
    $width=[int]$Window.Width-32;$height=[int]$Window.Height-48
    $Window.Content.Measure((New-Object Windows.Size($width,$height)))
    $Window.Content.Arrange((New-Object Windows.Rect(0,0,$width,$height)))
    $Window.Content.UpdateLayout()
    if($RenderDirectory) {
        $null=New-Item -ItemType Directory -Path $RenderDirectory -Force
        $bitmap=New-Object Windows.Media.Imaging.RenderTargetBitmap($width,$height,96,96,[Windows.Media.PixelFormats]::Pbgra32)
        $bitmap.Render($Window.Content)
        $encoder=New-Object Windows.Media.Imaging.PngBitmapEncoder;$encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
        $stream=[IO.File]::Create((Join-Path $RenderDirectory ($Name+'.png')))
        try{$encoder.Save($stream)}finally{$stream.Dispose()}
    }
}
Render $ui.Window 'windows-cleanup'
Render $script:window 'main-window'
if($ui.OptionsPanel.ActualWidth -lt 300 -or $ui.AppsGrid.ActualWidth -lt 450 -or $ui.RunButton.ActualWidth -lt 50){throw 'Windows layout clipped a primary control'}
Write-Host 'PASS: Offscreen layout gives cleanup options, app list and action buttons usable space'
foreach($item in @($ui.Window,$safe.Window,$nonAdmin.Window,$script:window)){$item.Close()}
Write-Host 'All 9 GUI construction/layout checks passed.'
