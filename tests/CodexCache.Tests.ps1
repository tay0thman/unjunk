#requires -Version 5.1
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Core.psm1') -Force
$module=Get-Module Unjunk.Core
$fixture=Join-Path $PSScriptRoot ('CodexCacheFixture-'+[Guid]::NewGuid().ToString('N'))
$previousLocal=$env:LOCALAPPDATA; $previousRoaming=$env:APPDATA; $previousUser=$env:USERPROFILE
$originalCatalog=& $module {(Get-Item Function:Get-UnjunkTargets).ScriptBlock}
$script:passed=0; $links=@()
function Assert([bool]$Condition,[string]$Message) {
    if(-not $Condition){throw ('FAIL: '+$Message)}
    $script:passed++; Write-Host ('PASS: '+$Message)
}
function FixtureFile([string]$Path,[switch]$Recent) {
    $null=New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($Path)) -Force
    [IO.File]::WriteAllText($Path,'fixture data')
    if(-not $Recent){(Get-Item -LiteralPath $Path).LastWriteTimeUtc=[DateTime]::UtcNow.AddDays(-14)}
    return $Path
}
try {
    $env:LOCALAPPDATA=Join-Path $fixture 'AppData\Local'
    $env:APPDATA=Join-Path $fixture 'AppData\Roaming'
    $env:USERPROFILE=$fixture
    $desktop=Join-Path $env:APPDATA 'Codex'
    $packaged=Join-Path $env:LOCALAPPDATA 'Packages\OpenAI.Codex_2p2nqsd0c76g0\LocalCache\Roaming\Codex'
    $expected=@(); $protected=@()
    foreach($homePath in @($desktop,$packaged)) {
        foreach($leaf in @('ShaderCache','GrShaderCache','GPUPersistentCache','Default\Cache','Default\Code Cache','Default\GPUCache','Default\DawnWebGPUCache','Default\DawnGraphiteCache','Default\Partitions\codex-browser-app\Cache','Default\Partitions\codex-browser-app\Code Cache','Default\Partitions\codex-browser-app\GPUCache','Default\Partitions\codex-browser-app\DawnWebGPUCache','Default\Partitions\codex-browser-app\DawnGraphiteCache','Profile 2\Cache')) {
            $root=Join-Path $homePath ('web\Codex\'+$leaf)
            $expected+= $root
            $null=FixtureFile (Join-Path $root 'old.bin')
        }
        $expected+= (Join-Path $homePath 'Cache')
        $null=FixtureFile (Join-Path $homePath 'Cache\old.bin')
        foreach($leaf in @('web\Codex\Local State','web\Codex\artifact-sessions\artifact.txt','web\Codex\executor-plugins\plugin.txt','web\Codex\extensions_crx_cache\extension.crx','web\Codex\Guest Profile\Cache\keep.bin','web\Codex\Default\Preferences','web\Codex\Default\History','web\Codex\Default\Login Data','web\Codex\Default\Network\Cookies','web\Codex\Default\IndexedDB\database.dat','web\Codex\Default\Local Storage\data.dat','web\Codex\Default\Session Storage\data.dat','web\Codex\Default\Service Worker\CacheStorage\offline.dat','web\Codex\Default\WebStorage\data.dat','web\Codex\Default\Partitions\codex-browser-app\Network\Cookies','web\Codex\Default\Partitions\codex-browser-app\Local Storage\data.dat')) {
            $protected+= (FixtureFile (Join-Path $homePath $leaf))
        }
    }
    foreach($leaf in @('.codex\auth.json','.codex\config.toml','.codex\sessions\session.jsonl','.codex\state_5.sqlite','.codex\cache\plugin-data.json','AppData\Local\OpenAI\Codex\runtimes\runtime.exe','AppData\Local\Packages\OpenAI.CodexPrimaryRuntime.v26-1105-1297-0_3k8sg7r9htsxt\LocalCache\Roaming\Codex\Cache\runtime.dat')) {
        $protected+= (FixtureFile (Join-Path $fixture $leaf))
    }
    $protectedHashes=@{}; foreach($path in $protected){$protectedHashes[$path]=(Get-FileHash -LiteralPath $path).Hash}
    # Verify discovery rejects both a linked profile and a linked cache leaf.
    $outside=Join-Path $fixture 'Outside'
    $linkedFile=FixtureFile (Join-Path $outside 'Cache\keep.bin')
    $profileLink=Join-Path $desktop 'web\Codex\Profile 3'
    $leafLink=Join-Path $desktop 'web\Codex\Profile 2\GPUCache'
    $null=New-Item -ItemType Junction -Path $profileLink -Target $outside; $links+= $profileLink
    $null=New-Item -ItemType Junction -Path $leafLink -Target (Join-Path $outside 'Cache'); $links+= $leafLink
    $catalog=@(Get-UnjunkTargets)
    $targets=@($catalog | Where-Object Name -like 'Codex / *')
    Assert ($targets.Count -eq $expected.Count -and @($expected | Where-Object {$_ -notin $targets.Root}).Count -eq 0) 'Desktop and Windows app caches include all verified browser, code, graphics and in-app browser leaves'
    Assert (@($targets.Root | Sort-Object -Unique).Count -eq $targets.Count) 'Codex targets are deduplicated'
    Assert (@($targets | Where-Object {$_.Risk -ne 'Cache' -or $_.Category -ne 'Desktop apps' -or -not $_.Recommended}).Count -eq 0) 'Codex caches are selectable and recommended in Desktop apps'
    Assert (@($targets | Where-Object {$_.Root -eq $profileLink -or $_.Root.StartsWith($profileLink+'\') -or $_.Root -eq $leafLink}).Count -eq 0) 'Linked profiles and cache folders are never offered'
    Assert (@($targets | Where-Object {$_.Processes -notcontains 'ChatGPT' -or $_.Processes -notcontains 'Codex' -or $_.Processes -contains 'codex*'}).Count -eq 0) 'Process guard covers ChatGPT.exe and Codex.exe without blocking on the persistent sandbox service'
    # Restrict public scan/cleanup to this disposable catalog. Simulate app processes only.
    & $module {
        param($t)
        $script:codexFixtureTargets=@($t); $script:codexFixtureProcess=''
        function script:Get-UnjunkTargets { $script:codexFixtureTargets }
        function script:Get-Process {
            [CmdletBinding()]param([string[]]$Name)
            if($script:codexFixtureProcess -and $Name -contains $script:codexFixtureProcess){[pscustomobject]@{ProcessName=$script:codexFixtureProcess}}
        }
    } $targets
    $recent=FixtureFile (Join-Path $desktop 'web\Codex\Default\Cache\recent.bin') -Recent
    $excluded=FixtureFile (Join-Path $desktop 'web\Codex\Default\Cache\excluded.bin')
    $safe=Invoke-UnjunkScan -CrowdStrikeSafe $true -Exclusions @($excluded)
    Assert (@($safe.Targets | Where-Object {$_.PolicyBlocked -or $_.MinimumAgeDays -ne 7}).Count -eq 0 -and @($safe.Targets.Files | Where-Object Path -eq $recent).Count -eq 0) 'Safe profile retains recent Codex caches and its seven-day cutoff'
    $scan=Invoke-UnjunkScan -CrowdStrikeSafe $false -Exclusions @($excluded)
    Assert (@($scan.Targets.Files | Where-Object Path -eq $recent).Count -eq 1 -and @($scan.Targets.Files | Where-Object Path -eq $excluded).Count -eq 0) 'Full cleanup includes recent caches and honors saved exclusions'
    Assert (@($scan.Targets.Files | Where-Object {$_.Path -in $protected}).Count -eq 0) 'Chats, credentials, settings, browser storage, extensions and runtimes never enter the scan'
    $windows=Invoke-UnjunkSystemFileScan -TempAgeDays 0 -IncludeCaches -Exclusions @($excluded)
    Assert ($windows.Targets.Count -eq $targets.Count -and @($windows.Targets | Where-Object SystemOption -ne 'app-caches').Count -eq 0) 'Windows cleanup and the one-click cache stage include every Codex target'
    $without=Invoke-UnjunkSystemFileScan -TempAgeDays 0
    Assert ($without.Targets.Count -eq 0) 'Codex caches require Application caches to be included in Windows cleanup'
    $preview=Invoke-UnjunkCleanup -Scan $scan -TargetIds $targets.Id -WhatIf
    Assert ($preview.DeletedFiles -eq 0 -and (Test-Path -LiteralPath $recent)) 'Preview does not delete Codex cache fixtures'
    foreach($process in @('ChatGPT','Codex','codex-computer-use-swift')) {
        & $module {param($p) $script:codexFixtureProcess=$p} $process
        $runningScan=Invoke-UnjunkScan -CrowdStrikeSafe $false
        Set-UnjunkCategorySelection -Rows $runningScan.Targets -Category 'Desktop apps' -Selected $true
        Assert (@($runningScan.Targets | Where-Object Selected).Count -eq 0) ("Select all skips caches while $process is running")
        # Recheck at cleanup time even if the app was closed when scanning.
        $blocked=Invoke-UnjunkCleanup -Scan $scan -TargetIds $targets.Id -Confirm:$false
        Assert ($blocked.DeletedFiles -eq 0 -and $blocked.SkippedFiles -eq $targets.Count -and (Test-Path -LiteralPath $recent)) ("Cleanup rechecks $process and preserves its caches")
    }
    & $module {$script:codexFixtureProcess=''}
    $newFile=FixtureFile (Join-Path $desktop 'web\Codex\Default\Cache\created-after-scan.bin')
    $changed=Join-Path $desktop 'web\Codex\Default\Cache\old.bin'
    Add-Content -LiteralPath $changed -Value 'changed after scan'
    $result=Invoke-UnjunkCleanup -Scan $scan -TargetIds $targets.Id -Confirm:$false
    Assert ($result.DeletedFiles -eq $expected.Count -and -not (Test-Path -LiteralPath $recent)) 'Closed-app cleanup deletes only eligible scanned cache files and reports their count'
    Assert ((Test-Path -LiteralPath $newFile) -and (Test-Path -LiteralPath $changed) -and (Test-Path -LiteralPath $excluded)) 'New, changed and excluded cache files survive'
    $intact=@($protected | Where-Object {(Test-Path -LiteralPath $_) -and (Get-FileHash -LiteralPath $_).Hash -eq $protectedHashes[$_]})
    Assert ($intact.Count -eq $protected.Count -and (Test-Path -LiteralPath $linkedFile)) 'All protected Codex data and linked-folder contents remain byte-for-byte intact'
    Write-Host ("All $script:passed Codex cache checks passed.")
} finally {
    & $module {
        param($c)
        Set-Item Function:script:Get-UnjunkTargets -Value $c
        Remove-Item Function:script:Get-Process -ErrorAction SilentlyContinue
        Remove-Variable codexFixtureTargets,codexFixtureProcess -Scope Script -ErrorAction SilentlyContinue
    } $originalCatalog
    $env:LOCALAPPDATA=$previousLocal; $env:APPDATA=$previousRoaming; $env:USERPROFILE=$previousUser
    $path=[IO.Path]::GetFullPath($fixture); $parent=[IO.Path]::GetFullPath($PSScriptRoot)+'\'
    if(-not $path.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $path -Leaf) -notlike 'CodexCacheFixture-*'){throw 'Unsafe fixture path'}
    foreach($link in $links){if((Test-Path -LiteralPath $link) -and [IO.Path]::GetFullPath($link).StartsWith($path+'\',[StringComparison]::OrdinalIgnoreCase)){[IO.Directory]::Delete($link)}}
    if(Test-Path -LiteralPath $path){Remove-Item -LiteralPath $path -Recurse -Force}
}
