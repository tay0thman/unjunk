#requires -Version 5.1
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Windows.psm1') -Force
$module=Get-Module Unjunk.Windows
$name='Local\Unjunk-BusyFixture-'+[Guid]::NewGuid().ToString('N')
& $module {param($name) $script:maintenanceMutexName=$name} $name
if(Test-UnjunkWindowsCleanupBusy){throw 'A missing mutex must not report busy'}
$state=[hashtable]::Synchronized(@{Ready=$false;Release=$false})
$worker=[PowerShell]::Create()
$null=$worker.AddScript({param($name,$state)
    $mutex=New-Object Threading.Mutex($false,$name)
    try {$null=$mutex.WaitOne();$state.Ready=$true;while(-not $state.Release){[Threading.Thread]::Sleep(10)}}
    finally {$mutex.ReleaseMutex();$mutex.Dispose()}
}).AddArgument($name).AddArgument($state)
$handle=$worker.BeginInvoke()
try {
    $deadline=[DateTime]::UtcNow.AddSeconds(5)
    while(-not $state.Ready -and [DateTime]::UtcNow -lt $deadline){[Threading.Thread]::Sleep(10)}
    if(-not $state.Ready){throw 'Busy fixture failed to acquire its private mutex'}
    if(-not (Test-UnjunkWindowsCleanupBusy)){throw 'Active maintenance on a different thread must report busy'}
    if(-not (Test-UnjunkWindowsCleanupBusy)){throw 'Checking activity must not release another thread''s lock'}
    $state.Release=$true;$null=$worker.EndInvoke($handle)
    if(Test-UnjunkWindowsCleanupBusy){throw 'Completed maintenance must stop reporting busy'}
    'All 4 maintenance-lock checks passed.'
} finally {$state.Release=$true;$worker.Dispose()}
