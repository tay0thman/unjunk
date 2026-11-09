#requires -Version 5.1
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Core.psm1') -Force
$module=Get-Module Unjunk.Core
$fixture=Join-Path $PSScriptRoot ('SystemScanFixture-'+[Guid]::NewGuid().ToString('N'))
$previousLocal=$env:LOCALAPPDATA
$env:LOCALAPPDATA=$fixture
try {
    foreach($name in @('Temp','Cache','Logs')){$dir=Join-Path $fixture $name;$null=New-Item -ItemType Directory -Path $dir -Force;Set-Content -LiteralPath (Join-Path $dir 'test.dat') -Value 'fixture'}
    & $module {param($root)
        $script:fixture=$root
        function script:Get-UnjunkTargets {
            New-UnjunkTarget 'Temp' 'Tests' (Join-Path $script:fixture 'Temp') -Risk Temporary
            New-UnjunkTarget 'Cache' 'Tests' (Join-Path $script:fixture 'Cache') -Risk Cache
            New-UnjunkTarget 'Logs' 'Tests' (Join-Path $script:fixture 'Logs') -Risk Diagnostics -MinimumAgeDays 30
        }
    } $fixture
    $scan=Invoke-UnjunkSystemFileScan -TempAgeDays 0 -IncludeCaches -IncludeLogs
    if(@($scan.Targets | Where-Object {$_.SystemOption -eq 'system-temp' -and $_.FileCount -eq 1}).Count -ne 1){throw 'User temp must be a separate temp stage'}
    if(@($scan.Targets | Where-Object {$_.SystemOption -eq 'app-caches' -and $_.FileCount -eq 1}).Count -ne 1){throw 'Application junk belongs to the cache stage'}
    if(@($scan.Targets | Where-Object {$_.SystemOption -eq 'app-logs' -and $_.FileCount -eq 1 -and $_.MinimumAgeDays -eq 0}).Count -ne 1){throw 'Selected diagnostic logs include all ages'}
    $limited=Invoke-UnjunkSystemFileScan -TempAgeDays 0 -IncludeCaches
    if(@($limited.Targets | Where-Object SystemOption -eq 'app-logs').Count){throw 'Log enumeration requires the explicit include flag'}
    'All 4 system-scan checks passed.'
} finally {
    $env:LOCALAPPDATA=$previousLocal
    $path=[IO.Path]::GetFullPath($fixture);$parent=[IO.Path]::GetFullPath($PSScriptRoot)+'\'
    if(-not $path.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $path -Leaf) -notlike 'SystemScanFixture-*'){throw 'Unsafe fixture path'}
    Remove-Item -LiteralPath $path -Recurse -Force
}
