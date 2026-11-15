#requires -Version 5.1
# Disposable Cinebench profiles and mock processes only. No real app data is cleaned.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Core.psm1') -Force
$module=Get-Module Unjunk.Core
$fixture=Join-Path $PSScriptRoot ('CinebenchFixture-'+[Guid]::NewGuid().ToString('N'))
$previousLocal=$env:LOCALAPPDATA;$previousRoaming=$env:APPDATA;$previousUser=$env:USERPROFILE
$originalCatalog=& $module {(Get-Item Function:Get-UnjunkTargets).ScriptBlock}
$originalClose=& $module {(Get-Item Function:Close-UnjunkTargetApps).ScriptBlock}
$passed=0;$links=@()
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
function FixtureFile([string]$Path,[int]$Age=45){
    $null=New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($Path)) -Force
    [IO.File]::WriteAllText($Path,'fixture data')
    (Get-Item -LiteralPath $Path).LastWriteTimeUtc=[DateTime]::UtcNow.AddDays(-$Age)
    return $Path
}
try {
    $env:LOCALAPPDATA=Join-Path $fixture 'Local';$env:APPDATA=Join-Path $fixture 'Roaming';$env:USERPROFILE=$fixture
    $maxon=Join-Path $env:APPDATA 'Maxon'
    $root=Join-Path $maxon 'Cinebench2026_win_x86_64_60A1F7D6'
    $cacheLeaves=@('Redshift\Cache','EBWebView\ShaderCache','EBWebView\GrShaderCache','EBWebView\GPUPersistentCache',
        'EBWebView\Default\Cache','EBWebView\Default\Code Cache','EBWebView\Default\GPUCache','EBWebView\Default\DawnWebGPUCache','EBWebView\Default\DawnGraphiteCache','EBWebView\Profile 2\Cache')
    $logLeaves=@('Redshift\Log','_bugreports','EBWebView\Crashpad\reports')
    $expected=@(foreach($leaf in @($cacheLeaves+$logLeaves)){$path=Join-Path $root $leaf;$null=FixtureFile (Join-Path $path 'old.bin');$path})
    $second=Join-Path $maxon 'Cinebench2026_win_x86_64_AABB1234\Redshift\Cache'
    $null=FixtureFile (Join-Path $second 'old.bin');$expected+= $second
    $null=FixtureFile (Join-Path $root 'Redshift\Cache\Textures\17\56\texture.rstexbin')
    $null=FixtureFile (Join-Path $root 'Redshift\Cache\Meshes\mesh.bin')
    $null=FixtureFile (Join-Path $root 'Redshift\Cache\CompiledShaders\shader.bin')
    $protected=@(foreach($leaf in @('cinebenchprefs\Cinema 4D.prf','cinebenchprefs\results.txt','Redshift\preferences.xml','language.prf',
        'plugins\custom.dll','library\scripts\custom.py','userrepository\asset.rstexbin','description.db\module.bin','dc\builtinrepository\data.bin',
        'EBWebView\Local State','EBWebView\Default\History','EBWebView\Default\Network\Cookies','EBWebView\Default\Login Data',
        'EBWebView\Default\Preferences','EBWebView\Default\Local Storage\data.bin','EBWebView\Default\IndexedDB\data.bin',
        'EBWebView\Default\Service Worker\CacheStorage\offline.bin','EBWebView\Guest Profile\Cache\keep.bin','UnknownCache\keep.bin')){FixtureFile (Join-Path $root $leaf)})
    foreach($other in @('Cinema4D2026_win_x86_64_60A1F7D6','Cinebench2025_win_x86_64_60A1F7D6','Cinebench2026_win_x86_64_NOTAHASH','Cinebench2026_win_x86_64_60A1F7D6_backup')){
        $protected+=FixtureFile (Join-Path $maxon ($other+'\Redshift\Cache\keep.bin'))
    }
    $outside=Join-Path $fixture 'outside';$protected+=FixtureFile (Join-Path $outside 'keep.bin')
    foreach($link in @((Join-Path $maxon 'Cinebench2026_win_x86_64_FFFFFFFF'),
        (Join-Path $maxon 'Cinebench2026_win_x86_64_FEEDBEEF\Redshift\Cache'),(Join-Path $root 'Redshift\Cache\redirect'))){
        $null=New-Item -ItemType Directory -Path (Split-Path -Parent $link) -Force
        $null=New-Item -ItemType Junction -Path $link -Target $outside;$links+=$link
    }
    $hashes=@{};foreach($path in $protected){$hashes[$path]=(Get-FileHash -LiteralPath $path).Hash}
    $targets=@(Get-UnjunkTargets | Where-Object Name -like 'Cinebench 2026 / *')
    Assert ($targets.Count -eq $expected.Count -and @($expected|Where-Object {$_ -notin $targets.Root}).Count -eq 0) 'Catalog discovers only explicit cache/log leaves across Cinebench 2026 build hashes'
    Assert (@($targets.Root|Sort-Object -Unique).Count -eq $targets.Count) 'Cinebench targets are not duplicated or overlapping'
    Assert (@($targets|Where-Object {$_.Root -eq $root -or $_.Root -eq $maxon}).Count -eq 0) 'Neither the whole Cinebench profile nor the Maxon directory is a cleanup target'
    Assert (@($targets|Where-Object {$_.Processes.Count -ne 1 -or $_.Processes[0] -ne 'Cinebench*'}).Count -eq 0) 'Process mapping covers Cinebench without closing unrelated Maxon apps or WebView hosts'
    Assert (@($targets|Where-Object {$_.Risk -eq 'Cache' -and (-not $_.Recommended -or $_.Category -ne 'Creative apps')}).Count -eq 0) 'Generated caches are selectable and recommended under Creative apps'
    Assert (@($targets|Where-Object {$_.Risk -eq 'Diagnostics' -and ($_.Recommended -or $_.MinimumAgeDays -ne 30)}).Count -eq 0) 'Logs keep the normal 30-day diagnostic policy and are not recommended by default'
    & $module {
        $script:cinebenchRunning=$false;$script:cinebenchCloseCalls=0
        function script:Get-UnjunkTargets {Get-UnjunkCinebenchTargets}
        function script:Get-Process {
            [CmdletBinding()]param([string[]]$Name)
            if($script:cinebenchRunning -and @($Name|Where-Object {'Cinebench' -like $_}).Count){[pscustomobject]@{ProcessName='Cinebench'}}
        }
        function script:Close-UnjunkTargetApps {
            param($Target,$State)
            if($Target.Processes -notcontains 'Cinebench*'){throw 'Unexpected process-close target'}
            if(-not $script:cinebenchRunning){return}
            $script:cinebenchCloseCalls++;$script:cinebenchRunning=$false
            [pscustomobject]@{Name='Cinebench';ProcessId=-1;Status='Closed';Reason='Fixture process only'}
        }
    }
    $recent=FixtureFile (Join-Path $root 'Redshift\Cache\recent.bin') 0
    $recentLog=FixtureFile (Join-Path $root 'Redshift\Log\current.log') 0
    $excluded=FixtureFile (Join-Path $root 'Redshift\Cache\excluded.bin')
    $safe=Invoke-UnjunkScan -CrowdStrikeSafe $true -Exclusions @($excluded)
    Assert (@($safe.Targets|Where-Object {$_.Risk -eq 'Cache' -and ($_.PolicyBlocked -or $_.MinimumAgeDays -ne 7)}).Count -eq 0 -and @($safe.Targets.Files|Where-Object Path -eq $recent).Count -eq 0) 'Safe profile preserves fresh caches and retains its seven-day minimum'
    Assert (@($safe.Targets|Where-Object {$_.Risk -eq 'Diagnostics' -and -not $_.PolicyBlocked}).Count -eq 0) 'Safe profile does not silently enable diagnostics cleanup'
    $scan=Invoke-UnjunkScan -CrowdStrikeSafe $false -Exclusions @($excluded)
    Assert (@($scan.Targets.Files|Where-Object Path -eq $recent).Count -eq 1 -and @($scan.Targets.Files|Where-Object Path -eq $recentLog).Count -eq 0) 'Standard scan includes new caches while retaining recent logs'
    Assert (@($scan.Targets.Files|Where-Object {$_.Path -in $protected -or $_.Path -eq $excluded}).Count -eq 0) 'Preferences, results, assets, browser data, unrelated builds and exclusions never enter the scan'
    Assert (($scan.Targets|Where-Object Root -eq (Join-Path $root 'Redshift\Cache')).SkippedLinks -eq 1) 'Nested cache junctions are skipped during enumeration'
    $system=Invoke-UnjunkSystemFileScan -TempAgeDays 0 -IncludeCaches -IncludeLogs -Exclusions @($excluded)
    Assert ($system.Targets.Count -eq $targets.Count -and @($system.Targets|Where-Object {$_.Risk -eq 'Cache' -and $_.SystemOption -ne 'app-caches'}).Count -eq 0) 'Windows cleanup and one-click automatically include every Cinebench cache'
    Assert (@($system.Targets.Files|Where-Object Path -eq $recentLog).Count -eq 1 -and @($system.Targets|Where-Object {$_.Risk -eq 'Diagnostics' -and ($_.SystemOption -ne 'app-logs' -or $_.MinimumAgeDays -ne 0)}).Count -eq 0) 'One-click log cleanup includes fresh Cinebench diagnostics in the log stage'
    Assert (@((Invoke-UnjunkSystemFileScan -IncludeCaches).Targets|Where-Object Risk -eq 'Diagnostics').Count -eq 0) 'The Windows cache-only selection does not also select logs'
    & $module {$script:cinebenchRunning=$true}
    $cacheIds=@($targets|Where-Object Risk -eq 'Cache'|ForEach-Object {$_.Id})
    $preview=Invoke-UnjunkCleanup -Scan $system -TargetIds $cacheIds -CloseApps -WhatIf
    Assert ($preview.DeletedFiles -eq 0 -and (& $module {$script:cinebenchCloseCalls}) -eq 0 -and (Test-Path -LiteralPath $recent)) 'WhatIf neither deletes files nor closes Cinebench'
    $blocked=Invoke-UnjunkCleanup -Scan $system -TargetIds $cacheIds -Confirm:$false
    Assert ($blocked.DeletedFiles -eq 0 -and (Test-Path -LiteralPath $recent)) 'Running Cinebench blocks deletion when Close apps is off'
    $newFile=FixtureFile (Join-Path $root 'Redshift\Cache\created-after-scan.bin')
    $changed=Join-Path $root 'Redshift\Cache\old.bin';Add-Content -LiteralPath $changed -Value 'changed after scan'
    $cacheCount=[long](($system.Targets|Where-Object Risk -eq 'Cache'|Measure-Object FileCount -Sum).Sum)
    $cleaned=Invoke-UnjunkCleanup -Scan $system -TargetIds $cacheIds -CloseApps -Confirm:$false
    Assert ((& $module {$script:cinebenchCloseCalls}) -eq 1 -and $cleaned.DeletedFiles -eq ($cacheCount-1) -and -not (Test-Path -LiteralPath $recent)) 'Close apps uses the mapped Cinebench process before deleting only scanned cache files'
    Assert ((Test-Path -LiteralPath $newFile) -and (Test-Path -LiteralPath $changed) -and (Test-Path -LiteralPath $excluded)) 'New, changed and excluded cache files survive cleanup'
    $logIds=@($targets|Where-Object Risk -eq 'Diagnostics'|ForEach-Object {$_.Id})
    $logs=Invoke-UnjunkCleanup -Scan $system -TargetIds $logIds -Confirm:$false
    Assert ($logs.DeletedFiles -eq 4 -and -not (Test-Path -LiteralPath $recentLog)) 'Explicit log cleanup removes only the inventoried diagnostic files'
    Assert (@($protected|Where-Object {-not (Test-Path -LiteralPath $_) -or (Get-FileHash -LiteralPath $_).Hash -ne $hashes[$_]}).Count -eq 0) 'All protected data and junction destinations remain byte-for-byte intact'
    Write-Host "All $passed Cinebench checks passed."
} finally {
    & $module {param($catalog,$close)
        Set-Item Function:script:Get-UnjunkTargets -Value $catalog
        Set-Item Function:script:Close-UnjunkTargetApps -Value $close
        Remove-Item Function:script:Get-Process -ErrorAction SilentlyContinue
    } $originalCatalog $originalClose
    $env:LOCALAPPDATA=$previousLocal;$env:APPDATA=$previousRoaming;$env:USERPROFILE=$previousUser
    $path=[IO.Path]::GetFullPath($fixture);$parent=[IO.Path]::GetFullPath($PSScriptRoot)+'\'
    if(-not $path.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $path -Leaf) -notlike 'CinebenchFixture-*'){throw 'Unsafe Cinebench fixture path'}
    foreach($link in $links){if((Test-Path -LiteralPath $link) -and [IO.Path]::GetFullPath($link).StartsWith($path+'\',[StringComparison]::OrdinalIgnoreCase)){[IO.Directory]::Delete($link)}}
    if(Test-Path -LiteralPath $path){Remove-Item -LiteralPath $path -Recurse -Force}
}
