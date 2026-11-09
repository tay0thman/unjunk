#requires -Version 5.1
Set-StrictMode -Version 2.0

function Format-UnjunkSize([double]$Bytes) {
    if ($Bytes -ge 1GB) { return ('{0:N2} GiB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MiB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N1} KiB' -f ($Bytes / 1KB)) }
    return ('{0:N0} B' -f $Bytes)
}

function Get-UnjunkFullPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'An empty path is not allowed.' }
    $expanded = [Environment]::ExpandEnvironmentVariables($Path)
    if ($expanded -notmatch '^[a-zA-Z]:\\' -or $expanded.Substring(2).Contains(':')) {
        throw "Only absolute local drive paths are supported: $Path"
    }
    return [IO.Path]::GetFullPath($expanded).TrimEnd('\')
}

function Test-UnjunkWithin([string]$Path, [string]$Root) {
    return $Path.Equals($Root, [StringComparison]::OrdinalIgnoreCase) -or
        $Path.StartsWith($Root.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)
}

function Test-UnjunkExcluded([string]$Path, [string[]]$Exclusions) {
    foreach ($excluded in $Exclusions) {
        if (Test-UnjunkWithin $Path $excluded) { return $true }
    }
    return $false
}

function Assert-UnjunkRoot([string]$Root) {
    $full = Get-UnjunkFullPath $Root
    if ($full.Length -lt 4 -or $full -notlike "$env:SystemDrive\*") { throw "Unsupported cleanup root: $full" }
    $protected = @("$env:SystemRoot\Installer", "$env:SystemRoot\WinSxS", "$env:SystemRoot\servicing",
        "$env:SystemRoot\System32", "$env:ProgramData\Package Cache", "$env:ProgramData\Microsoft\Windows Defender",
        "$env:ProgramFiles", ${env:ProgramFiles(x86)}, "$env:SystemDrive\System Volume Information")
    foreach ($path in $protected) {
        if ($path -and (Test-UnjunkWithin $full (Get-UnjunkFullPath $path))) { throw "Protected location: $full" }
    }
    foreach ($path in @($env:SystemRoot, $env:USERPROFILE, $env:LOCALAPPDATA, $env:APPDATA, $env:ProgramData, "$env:SystemDrive\Users")) {
        if ($path -and $full.Equals((Get-UnjunkFullPath $path), [StringComparison]::OrdinalIgnoreCase)) { throw "Root is too broad: $full" }
    }
    return $full
}

function Test-UnjunkNoLinks([string]$Path) {
    $cursor = $Path
    while ($cursor) {
        $item = Get-Item -LiteralPath $cursor -Force -ErrorAction Stop
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return $false }
        $parent = [IO.Path]::GetDirectoryName($cursor)
        if ($parent -eq $cursor) { break }
        $cursor = $parent
    }
    return $true
}

function New-UnjunkTarget {
    param([string]$Name, [string]$Category, [string]$Root, [string[]]$Processes = @(),
        [string]$Risk = 'Cache', [int]$MinimumAgeDays = 0, [string]$Note = 'Recreated when the application needs it.',
        [string]$FilePattern = '*', [bool]$Recurse = $true, [bool]$ScanOnly = $false, $RhinoPackage = $null)
    $full = Get-UnjunkFullPath $Root
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $id = ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($full.ToLowerInvariant())))).Replace('-', '').Substring(0, 16) }
    finally { $sha.Dispose() }
    [pscustomobject]@{ Id = $id; Name = $Name; Category = $Category; Root = $full; Processes = @($Processes);
        Risk = $Risk; Recommended = ($Risk -in @('Cache', 'Temporary')); MinimumAgeDays = $MinimumAgeDays; Note = $Note;
        FilePattern = $FilePattern; Recurse = $Recurse; ScanOnly=$ScanOnly; RhinoPackage=$RhinoPackage }
}

function ConvertTo-UnjunkPackageVersion([string]$Text) {
    $match=[regex]::Match($Text,'^(?<core>(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)(?:\.(?:0|[1-9]\d*))?)(?:-(?<pre>[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$')
    if(-not $match.Success) { return $null }
    $numbers=@($match.Groups['core'].Value.Split('.'))
    while($numbers.Count -lt 4) { $numbers+= '0' }
    $pre=$match.Groups['pre'].Value
    foreach($part in @($pre.Split('.'))) { if($part -match '^0\d+$') { return $null } }
    [pscustomobject]@{Text=$Text;Numbers=$numbers;Prerelease=$pre}
}

function Compare-UnjunkPackageVersion($Left,$Right) {
    for($i=0;$i -lt 4;$i++) {
        $a=$Left.Numbers[$i]; $b=$Right.Numbers[$i]
        if($a.Length -ne $b.Length) { return [Math]::Sign($a.Length-$b.Length) }
        $cmp=[string]::CompareOrdinal($a,$b); if($cmp) { return [Math]::Sign($cmp) }
    }
    if(-not $Left.Prerelease -and -not $Right.Prerelease) { return 0 }
    if(-not $Left.Prerelease) { return 1 }; if(-not $Right.Prerelease) { return -1 }
    $a=@($Left.Prerelease.Split('.')); $b=@($Right.Prerelease.Split('.'))
    for($i=0;$i -lt [Math]::Min($a.Count,$b.Count);$i++) {
        $an=$a[$i] -match '^\d+$'; $bn=$b[$i] -match '^\d+$'
        if($an -and -not $bn) { return -1 }; if($bn -and -not $an) { return 1 }
        if($an -and $a[$i].Length -ne $b[$i].Length) { return [Math]::Sign($a[$i].Length-$b[$i].Length) }
        $cmp=[string]::CompareOrdinal($a[$i],$b[$i]); if($cmp) { return [Math]::Sign($cmp) }
    }
    return [Math]::Sign($a.Count-$b.Count)
}

function Test-UnjunkPackageManifest([string]$VersionRoot,[string]$Version) {
    $path=Join-Path $VersionRoot 'manifest.yml'
    try {
        if(-not (Test-UnjunkNoLinks $path)) { return $false }
        $lines=@(Get-Content -LiteralPath $path -ErrorAction Stop | Where-Object {$_ -match '^version\s*:'})
        if($lines.Count -ne 1) { return $false }
        $value=($lines[0] -replace '^version\s*:\s*','').Trim().Trim('"',"'")
        return $value -ceq $Version
    } catch { return $false }
}

function Get-UnjunkRhinoPackageTargets {
    param([string]$Root="$env:APPDATA\McNeel\Rhinoceros\packages\9.0")
    $targets=New-Object 'System.Collections.Generic.List[object]'
    $details=New-Object 'System.Collections.Generic.List[string]'
    $packages=0
    try {
        if(-not (Test-UnjunkNoLinks $Root)) { throw 'Package root is a link.' }
        $folders=@(Get-ChildItem -LiteralPath $Root -Directory -Force -ErrorAction Stop)
        foreach($package in $folders) {
            $packages++
            try {
                if(-not (Test-UnjunkNoLinks $package.FullName)) { throw 'Package folder is a link.' }
                $manifest=Join-Path $package.FullName 'manifest.txt'
                if(-not (Test-UnjunkNoLinks $manifest)) { throw 'Package pointer is a link.' }
                $active=(Get-Content -LiteralPath $manifest -Raw -ErrorAction Stop).Trim()
                if(-not (ConvertTo-UnjunkPackageVersion $active)) { throw 'Missing or unrecognized active version.' }
                $activeRoot=Join-Path $package.FullName $active
                if(-not (Test-Path -LiteralPath $activeRoot -PathType Container) -or -not (Test-UnjunkPackageManifest $activeRoot $active)) { throw 'Active version or its manifest cannot be verified.' }
                $versions=@(foreach($folder in @(Get-ChildItem -LiteralPath $package.FullName -Directory -Force -ErrorAction Stop)) {
                    $parsed=ConvertTo-UnjunkPackageVersion $folder.Name
                    if($parsed -and (Test-UnjunkNoLinks $folder.FullName)) { [pscustomobject]@{Folder=$folder;Version=$parsed} }
                })
                $latest=$null; $stable=$null
                foreach($version in $versions) {
                    if($null -eq $latest -or (Compare-UnjunkPackageVersion $version.Version $latest.Version) -gt 0) { $latest=$version }
                    if(-not $version.Version.Prerelease -and ($null -eq $stable -or (Compare-UnjunkPackageVersion $version.Version $stable.Version) -gt 0)) { $stable=$version }
                }
                $count=0
                foreach($version in $versions) {
                    if($version.Folder.Name -eq $active -or (Compare-UnjunkPackageVersion $version.Version $latest.Version) -ge 0) { continue }
                    if($null -ne $stable -and (Compare-UnjunkPackageVersion $version.Version $stable.Version) -eq 0) { continue }
                    if(-not (Test-UnjunkPackageManifest $version.Folder.FullName $version.Folder.Name)) { continue }
                    $info=[pscustomobject]@{PackagesRoot=$Root;PackageRoot=$package.FullName;ManifestPath=$manifest;ActiveVersion=$active;LatestVersion=$latest.Folder.Name;Version=$version.Folder.Name}
                    $note="Older package version. Keep active $active and newest $($latest.Folder.Name) (also newest stable). Close Rhino and Yak first. Removing this version removes its rollback copy."
                    $targets.Add((New-UnjunkTarget "Rhino 9 / $($package.Name) / old $($version.Folder.Name)" 'Rhino plug-ins' $version.Folder.FullName @('Rhino*','Yak') -Risk Review -Note $note -RhinoPackage $info))
                    $count++
                }
                $details.Add("$($package.Name): active $active; newest $($latest.Folder.Name); $count older candidate(s).")
            } catch { $details.Add("$($package.Name): preserved; $($_.Exception.Message)") }
        }
    } catch { $details.Add("Package scan unavailable: $($_.Exception.Message)") }
    $note="$packages package folders checked; $($targets.Count) older version candidate(s). Active, newest, newest stable, unknown versions and ambiguous packages are preserved.`r`n"+($details -join "`r`n")
    $targets.Add((New-UnjunkTarget 'Rhino 9 plug-in version inventory' 'Rhino plug-ins' $Root -Risk Review -Note $note -ScanOnly $true))
    return $targets.ToArray()
}

function Get-UnjunkCodexCacheTargets {
    # Enumerate only verified cache leaves, never a whole Codex profile or .codex home.
    $homes=@([pscustomobject]@{Name='Desktop';Root="$env:APPDATA\Codex"})
    $packages="$env:LOCALAPPDATA\Packages"
    try {
        if(Test-UnjunkNoLinks $packages) {
            foreach($package in @(Get-ChildItem -LiteralPath $packages -Directory -Filter 'OpenAI.Codex_*' -ErrorAction Stop)) {
                if($package.Name -match '^OpenAI\.Codex_[a-z0-9]{13}$' -and (Test-UnjunkNoLinks $package.FullName)) {
                    $homes+= [pscustomobject]@{Name='Windows app';Root=(Join-Path $package.FullName 'LocalCache\Roaming\Codex')}
                }
            }
        }
    } catch { } # A missing/inaccessible package directory is not a cleanup target.
    $profileCaches=@('Cache','Code Cache','GPUCache','DawnWebGPUCache','DawnGraphiteCache')
    $shaderCaches=@('ShaderCache','GrShaderCache','GPUPersistentCache')
    # The Windows desktop executable is ChatGPT.exe. Avoid codex*: its sandbox service can stay running.
    $processes=@('ChatGPT','Codex','codex-computer-use-swift')
    $note='Close Codex first or enable Close apps before cache cleanup. Recreated browser/code/graphics cache only; chats, projects, sign-in data, settings, extensions and runtimes are preserved.'
    foreach($codexHome in $homes) {
        try { if(-not (Test-UnjunkNoLinks $codexHome.Root)) { continue } } catch { continue }
        $locations=@([pscustomobject]@{Name=$codexHome.Name;Root=$codexHome.Root;Leaves=@($profileCaches+$shaderCaches)})
        $web=Join-Path $codexHome.Root 'web\Codex'
        try {
            if(Test-UnjunkNoLinks $web) {
                $locations+= [pscustomobject]@{Name=($codexHome.Name+' / Browser');Root=$web;Leaves=$shaderCaches}
                foreach($profile in @(Get-ChildItem -LiteralPath $web -Directory -ErrorAction Stop | Where-Object {$_.Name -eq 'Default' -or $_.Name -match '^Profile \d+$'})) {
                    if(-not (Test-UnjunkNoLinks $profile.FullName)) { continue }
                    $locations+= [pscustomobject]@{Name=($codexHome.Name+' / '+$profile.Name);Root=$profile.FullName;Leaves=$profileCaches}
                    $locations+= [pscustomobject]@{Name=($codexHome.Name+' / '+$profile.Name+' / In-app browser');Root=(Join-Path $profile.FullName 'Partitions\codex-browser-app');Leaves=$profileCaches}
                }
            }
        } catch { } # Do not broaden the target when a profile cannot be inspected.
        foreach($location in $locations) {
            foreach($leaf in $location.Leaves) {
                $path=Join-Path $location.Root $leaf
                try {
                    if((Test-Path -LiteralPath $path -PathType Container) -and (Test-UnjunkNoLinks $path)) {
                        New-UnjunkTarget ("Codex / $($location.Name) / $leaf") 'Desktop apps' $path $processes -Note $note
                    }
                } catch { }
            }
        }
    }
}

function Get-UnjunkTargets {
    $list = New-Object 'System.Collections.Generic.List[object]'
    $list.Add((New-UnjunkTarget 'User temporary files' 'Windows' "$env:LOCALAPPDATA\Temp" -Risk Temporary -MinimumAgeDays 7 -Note 'Only files older than the chosen age. Locked files are skipped.'))
    $list.Add((New-UnjunkTarget 'Windows temporary files' 'Windows' "$env:SystemRoot\Temp" -Risk Temporary -MinimumAgeDays 7))
    $list.Add((New-UnjunkTarget 'DirectX shader cache' 'Graphics' "$env:LOCALAPPDATA\D3DSCache" -Note 'Shaders rebuild; the next game launch may be slower.'))
    foreach ($name in @('DXCache', 'GLCache', 'OptixCache')) {
        $list.Add((New-UnjunkTarget "NVIDIA $name" 'Graphics' "$env:LOCALAPPDATA\NVIDIA\$name" -Note 'Close games and rendering applications first. Shaders rebuild.'))
    }
    $browsers = @(
        @{Name='Edge'; Root="$env:LOCALAPPDATA\Microsoft\Edge\User Data"; Process='msedge'},
        @{Name='Chrome'; Root="$env:LOCALAPPDATA\Google\Chrome\User Data"; Process='chrome'},
        @{Name='Brave'; Root="$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data"; Process='brave'}
    )
    foreach ($browser in $browsers) {
        $profiles = @(Get-ChildItem -LiteralPath $browser.Root -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -eq 'Default' -or $_.Name -match '^Profile \d+$' })
        foreach ($profile in $profiles) {
            foreach ($cache in @('Cache', 'Code Cache', 'GPUCache')) {
                $list.Add((New-UnjunkTarget "$($browser.Name) / $($profile.Name) / $cache" 'Browsers' (Join-Path $profile.FullName $cache) @($browser.Process)))
            }
        }
    }
    $apps = @(
        @{Name='Claude'; Folder='Claude'; Process='Claude'}, @{Name='Vagon'; Folder='Vagon'; Process='Vagon*'},
        @{Name='Gemini'; Folder='gemini'; Process='gemini'}, @{Name='Oculus'; Folder='OculusClient'; Process='OculusClient'},
        @{Name='Calibrite'; Folder='calibrite-profiler'; Process='calibrite-profiler'},
        @{Name='Discord'; Folder='discord'; Process='Discord'}, @{Name='Slack'; Folder='Slack'; Process='slack'},
        @{Name='VS Code'; Folder='Code'; Process='Code'}, @{Name='Vectorworks Install Manager'; Folder='vectorworks-install-manager'; Process='*Vectorworks*'}
    )
    foreach ($app in $apps) {
        foreach ($cache in @('Cache', 'Code Cache', 'GPUCache', 'DawnWebGPUCache', 'DawnGraphiteCache', 'GrShaderCache')) {
            $list.Add((New-UnjunkTarget "$($app.Name) / $cache" 'Desktop apps' "$env:APPDATA\$($app.Folder)\$cache" @($app.Process)))
        }
    }
    foreach($target in @(Get-UnjunkCodexCacheTargets)) { $list.Add($target) }
    foreach ($entry in @(
        @('pip download cache', 'Developer', "$env:LOCALAPPDATA\pip\Cache", 'python*'),
        @('NuGet HTTP cache', 'Developer', "$env:LOCALAPPDATA\NuGet\v3-cache", 'dotnet'),
        @('VS Code VSIX downloads', 'Developer', "$env:APPDATA\Code\CachedExtensionVSIXs", 'Code'),
        @('Steam web cache', 'Games', "$env:LOCALAPPDATA\Steam\htmlcache", 'steam*'),
        @('Adobe media cache', 'Creative apps', "$env:APPDATA\Adobe\Common\Media Cache Files", '*Adobe*'),
        @('Lightroom cache', 'Creative apps', "$env:LOCALAPPDATA\Adobe\Lightroom\Caches", 'Lightroom*'),
        @('V-Ray UI cache', 'Creative apps', "$env:APPDATA\Chaos\V-Ray for Rhinoceros\vrayneui\Partitions\neui\Cache", 'Rhino'),
        @('Rhino downloaded updates', 'Creative apps', "$env:LOCALAPPDATA\McNeel\McNeelUpdate\DownloadCache", 'Rhino'),
        @('UniGetUI media cache', 'Desktop apps', "$env:LOCALAPPDATA\UniGetUI\CachedMedia", 'UniGetUI')
    )) {
        $processes=@($entry[3])
        if($entry[0] -eq 'Adobe media cache') { $processes+=@('Adobe Premiere Pro','AfterFX','Adobe Media Encoder','Audition') }
        $list.Add((New-UnjunkTarget $entry[0] $entry[1] $entry[2] $processes))
    }
    $list.Add((New-UnjunkTarget 'Rhino 9 settings backup ZIPs' 'Creative apps' "$env:APPDATA\McNeel\Rhinoceros\9.0" @('Rhino*') -Risk Review -FilePattern '*.zip' -Recurse $false -Note 'ZIP files directly in the Rhino 9.0 folder only. Review the filenames before selecting: these may be settings backups. Deleting them removes those restore copies. Active settings and all subfolders are preserved.'))
    # Only known cache leaves are expanded. No complete web profiles or collaboration data.
    foreach ($entry in @(
        @('Revit browser code', "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit *\WebBrowserControl\EBWebView\*\Code Cache", 'Revit'),
        @('InDesign cache', "$env:LOCALAPPDATA\Adobe\InDesign\*\*\Caches", 'InDesign'),
        @('Bluebeam web cache', "$env:LOCALAPPDATA\Bluebeam\Revu\*\WebCache", 'Revu'),
        @('Epic web cache', "$env:LOCALAPPDATA\EpicGamesLauncher\Saved\webcache_*", 'EpicGamesLauncher'),
        @('Teams browser cache', "$env:LOCALAPPDATA\Packages\MSTeams_*\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_tfw\Cache", 'ms-teams')
    )) {
        foreach ($resolved in @(Resolve-Path -Path $entry[1] -ErrorAction SilentlyContinue)) {
            $list.Add((New-UnjunkTarget $entry[0] 'Creative & desktop apps' $resolved.ProviderPath @($entry[2])))
        }
    }
    foreach ($entry in @(
        @('Vectorworks Cloud installer', 'vectorworks-cloud-services-updater', '*Vectorworks*'),
        @('Calibrite installer', 'calibrite-profiler-updater', 'calibrite-profiler'),
        @('Vagon installer', 'vagon-app-updater', 'Vagon*'),
        @('Upscayl installer', 'upscayl-updater', 'upscayl')
    )) {
        $list.Add((New-UnjunkTarget $entry[0] 'Downloaded installers' "$env:LOCALAPPDATA\$($entry[1])\installer.exe" @($entry[2]) -Risk Review -MinimumAgeDays 7 -Note 'Downloaded installer only. Remove if the installation is finished and you do not need this copy.'))
    }
    foreach ($folder in @('AMD', 'NVIDIA')) {
        $list.Add((New-UnjunkTarget "$folder extracted installers" 'Downloaded installers' "$env:SystemDrive\$folder" -Risk Review -MinimumAgeDays 30 -Note 'Review the file preview. Keep if an installation is pending or you need offline setup media.'))
    }
    $list.Add((New-UnjunkTarget 'Autodesk extracted installers' 'Downloaded installers' "$env:SystemDrive\Autodesk" @('AdskInstall','AdskInstaller','AdskAccessUIHost','AdskAccessServiceHost','Installer','Setup','msiexec') -Risk Review -MinimumAgeDays 0 -Note 'Extracted installation/download files of all ages, including WI and Access. Close apps stops the Autodesk Access updater service and UI first (administrator required), without changing startup settings. Active installers and Network License Manager are preserved. Empty extraction folders are removed.'))
    foreach ($entry in @(
        @('Application crash dumps', "$env:LOCALAPPDATA\CrashDumps"),
        @('Windows minidumps', "$env:SystemRoot\Minidump"),
        @('Live kernel reports', "$env:SystemRoot\LiveKernelReports"),
        @('System memory dump', "$env:SystemRoot\MEMORY.DMP"),
        @('Autodesk error reports', "$env:LOCALAPPDATA\Autodesk\CER"),
        @('Google Drive diagnostic logs', "$env:LOCALAPPDATA\Google\DriveFS\Logs"),
        @('Zoom logs', "$env:APPDATA\Zoom\logs")
    )) { $list.Add((New-UnjunkTarget $entry[0] 'Diagnostics' $entry[1] -Risk Diagnostics -MinimumAgeDays 30 -Note 'Keep while investigating freezes or app failures. Removing these loses diagnostic history.')) }
    foreach($target in @(Get-UnjunkRhinoPackageTargets)) { $list.Add($target) }
    return @($list.ToArray() | Where-Object { $_.Root -like "$env:SystemDrive\*" } | Sort-Object Id -Unique)
}

function Get-UnjunkRunningApps($Target) {
    $names = @()
    foreach ($pattern in $Target.Processes) {
        $names += @(Get-Process -Name $pattern -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ProcessName)
    }
    return @($names | Sort-Object -Unique)
}

function Get-UnjunkAppProcesses($Target) {
    @(foreach($pattern in $Target.Processes) { Get-Process -Name $pattern -ErrorAction SilentlyContinue }) | Sort-Object Id -Unique
}

function Get-UnjunkInteractiveSession {
    $current=[Diagnostics.Process]::GetCurrentProcess()
    try { return $current.SessionId } finally { $current.Dispose() }
}

function Test-UnjunkAutodeskTarget($Target) {
    return $Target.Risk -eq 'Review' -and $null -ne $Target.PSObject.Properties['Root'] -and
        $Target.Root -eq (Get-UnjunkFullPath "$env:SystemDrive\Autodesk")
}

function Stop-UnjunkAutodeskAccess {
    param($State)
    if($State.Cancel){return}
    # Exact Autodesk Access service names only. Never alter startup type or stop licensing services.
    foreach($service in @(Get-Service -Name 'Autodesk Access Service Host','AdskAccessServiceHost' -ErrorAction SilentlyContinue)) {
        if($State.Cancel){break}
        try {
            if($service.Status -ne [ServiceProcess.ServiceControllerStatus]::Stopped){
                $State.Message='Stopping Autodesk Access before cleaning extracted installers...'
                $service.Stop()
                $service.WaitForStatus([ServiceProcess.ServiceControllerStatus]::Stopped,[TimeSpan]::FromSeconds(15))
                [pscustomobject]@{Name=$service.ServiceName;ProcessId=0;Status='ServiceStopped';Reason='Autodesk Access stopped for cleanup; startup type is unchanged.'}
            }
            # A stopped service may leave its host alive briefly. Only this exact Autodesk host is eligible.
            foreach($process in @(Get-Process -Name AdskAccessServiceHost -ErrorAction SilentlyContinue)) {
                try {
                    if($State.Cancel){break}
                    $null=$process.Handle
                    if(-not $process.HasExited){$process.Kill();if(-not $process.WaitForExit(3000)){throw 'Autodesk Access host did not exit.'}}
                    [pscustomobject]@{Name='AdskAccessServiceHost';ProcessId=$process.Id;Status='Terminated';Reason='Remaining Autodesk Access host stopped after its service stopped.'}
                } finally {$process.Dispose()}
            }
        } catch {
            [pscustomobject]@{Name=$service.ServiceName;ProcessId=0;Status='Skipped';Reason=('Could not stop Autodesk Access. Run Unjunk as administrator. '+$_.Exception.Message)}
        } finally {$service.Dispose()}
    }
}

function Test-UnjunkProtectedProcess($Process) {
    # Only mapped interactive apps may be closed, never Windows, servicing, security or this host.
    if($Process.Id -le 4 -or $Process.Id -eq $PID -or $Process.SessionId -eq 0 -or $Process.SessionId -ne (Get-UnjunkInteractiveSession)) { return $true }
    return $Process.ProcessName -match '^(System|Idle|Registry|smss|csrss|wininit|winlogon|services|lsass|svchost|dwm|explorer|powershell|pwsh|cmd|WindowsTerminal|conhost|OpenConsole|Unjunk|Dism|DismHost|cleanmgr|TrustedInstaller|TiWorker|msiexec|Setup|Installer|AdskInstall.*|MsMpEng|MsSense|Sense.*|NisSrv|.*Falcon.*|.*CrowdStrike.*|.*Sentinel.*|.*Cylance.*)$'
}

function Close-UnjunkTargetApps {
    param($Target,$State,[ValidateRange(0,10)][int]$GraceSeconds=5)
    $autodesk=Test-UnjunkAutodeskTarget $Target
    if(($Target.Risk -ne 'Cache' -and -not $autodesk) -or $State.Cancel) { return }
    $processTarget=$Target
    if($autodesk){
        if(@(Get-UnjunkRunningApps $Target | Where-Object {$_ -notin @('AdskAccessUIHost','AdskAccessServiceHost')}).Count){return}
        Stop-UnjunkAutodeskAccess -State $State
        $processTarget=[pscustomobject]@{Name=$Target.Name;Processes=@('AdskAccessUIHost')}
    }
    $pending=New-Object 'System.Collections.Generic.List[object]'
    try {
        foreach($process in @(Get-UnjunkAppProcesses $processTarget)) {
            if($State.Cancel){$process.Dispose();continue}
            try {
                if(Test-UnjunkProtectedProcess $process) { throw 'Protected process or another Windows session; left running.' }
                # Retain the original process handle so a reused PID cannot become the target.
                $null=$process.Handle
                if($process.HasExited) { $process.Dispose(); continue }
                $pending.Add($process)
                if($process.MainWindowHandle -ne [IntPtr]::Zero) { $null=$process.CloseMainWindow() }
            } catch {
                [pscustomobject]@{Name=$process.ProcessName;ProcessId=$process.Id;Status='Skipped';Reason=$_.Exception.Message}
                if(-not $pending.Contains($process)) { $process.Dispose() }
            }
        }
        $deadline=[DateTime]::UtcNow.AddSeconds($GraceSeconds)
        while(-not $State.Cancel -and [DateTime]::UtcNow -lt $deadline -and @($pending | Where-Object {-not $_.HasExited}).Count) {
            $State.Message='Closing apps before cache cleanup: '+$Target.Name
            Start-Sleep -Milliseconds 100
        }
        foreach($process in $pending) {
            $status='Closed'; $reason='Application exited before cache deletion.'
            try {
                if(-not $process.HasExited) {
                    if($State.Cancel) { $status='Skipped'; $reason='Cleanup cancelled before process termination.' }
                    else {
                        $process.Kill()
                        if(-not $process.WaitForExit(3000)) { throw 'Process did not exit; its cache will be skipped.' }
                        $status='Terminated'; $reason='Still running after the close request; stopped before cache deletion.'
                    }
                }
            } catch { $status='Skipped'; $reason=$_.Exception.Message }
            [pscustomobject]@{Name=$process.ProcessName;ProcessId=$process.Id;Status=$status;Reason=$reason}
        }
    } finally { foreach($process in $pending) { $process.Dispose() } }
}

function Test-UnjunkSafeTarget($Target) {
    if($Target.Risk -notin @('Cache','Temporary')) { return $false }
    $root=Get-UnjunkFullPath $Target.Root
    return (Test-UnjunkWithin $root (Get-UnjunkFullPath $env:LOCALAPPDATA)) -or
        (Test-UnjunkWithin $root (Get-UnjunkFullPath $env:APPDATA))
}

function Test-UnjunkTargetFile([string]$Path, $Target) {
    if($Target.ScanOnly) { return $false }
    $full=Get-UnjunkFullPath $Path
    $root=Get-UnjunkFullPath $Target.Root
    if(-not (Test-UnjunkWithin $full $root)) { return $false }
    if($root -eq (Get-UnjunkFullPath "$env:SystemDrive\Autodesk") -and $full.Substring($root.Length) -match '(?i)\\Network License Manager(\\|$)') { return $false }
    if([IO.Path]::GetFileName($full) -notlike $Target.FilePattern) { return $false }
    if(-not $Target.Recurse -and $full -ne $root -and [IO.Path]::GetDirectoryName($full) -ne $root) { return $false }
    return $true
}

function Get-UnjunkTargetScan {
    param($Target, [int]$TempAgeDays, [string[]]$Exclusions, $State, [bool]$CrowdStrikeSafe=$false)
    $root = Assert-UnjunkRoot $Target.Root
    $files = New-Object 'System.Collections.Generic.List[object]'
    $issues = New-Object 'System.Collections.Generic.List[string]'
    $bytes = 0L; $links = 0; $recent = 0; $excluded = 0; $exists = $false
    $age = $Target.MinimumAgeDays
    if ($Target.Risk -eq 'Temporary') { $age = [Math]::Max(0, $TempAgeDays) }
    if ($CrowdStrikeSafe) { $age=[Math]::Max(7,$age) }
    $policyBlocked=$CrowdStrikeSafe -and -not (Test-UnjunkSafeTarget $Target)
    $cutoff = [DateTime]::UtcNow.AddDays(-$age)
    $stack = New-Object 'System.Collections.Generic.Stack[string]'
    try {
        if($policyBlocked) { throw [OperationCanceledException]::new('Outside the CrowdStrike Safe cleaning profile.') }
        $rootItem = Get-Item -LiteralPath $root -Force -ErrorAction Stop
        $exists = $true
        if (Test-UnjunkNoLinks $root) { if(-not $Target.ScanOnly) { $stack.Push($root) } } else { $links++ }
    } catch [OperationCanceledException] {
        # Policy-blocked targets are not enumerated, even when running elevated.
    } catch [System.Management.Automation.ItemNotFoundException] {
        # Missing targets are normal for apps not installed on this machine.
    } catch { $issues.Add($_.Exception.Message) }
    while ($stack.Count -gt 0 -and -not $State.Cancel) {
        $path = $stack.Pop()
        if (Test-UnjunkExcluded $path $Exclusions) { $excluded++; continue }
        try {
            $item = Get-Item -LiteralPath $path -Force -ErrorAction Stop
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { $links++; continue }
            if ($item.PSIsContainer) {
                if(-not $Target.Recurse -and $item.FullName -ne $root) { continue }
                foreach ($child in @(Get-ChildItem -LiteralPath $path -Force -ErrorAction Stop)) { $stack.Push($child.FullName) }
            } else {
                if(-not (Test-UnjunkTargetFile $item.FullName $Target)) { continue }
                if ($item.LastWriteTimeUtc -gt $cutoff) { $recent++; continue }
                $files.Add([pscustomobject]@{ Path=$item.FullName; Length=[long]$item.Length; LastWriteTicks=$item.LastWriteTimeUtc.Ticks; CreationTicks=$item.CreationTimeUtc.Ticks })
                $bytes += $item.Length
            }
        } catch { if ($issues.Count -lt 50) { $issues.Add($_.Exception.Message) } }
    }
    $running = @(Get-UnjunkRunningApps $Target)
    $status = if ($Target.ScanOnly) { 'Inventory only' } elseif ($policyBlocked) { 'Blocked by Safe profile' } elseif ($issues.Count) { 'Partial / access errors' } elseif ($links) { 'Links skipped' } elseif (-not $exists) { 'Not present' } elseif ($running.Count) { 'Close: ' + ($running -join ', ') } elseif ($files.Count) { 'Ready' } else { 'No eligible files' }
    [pscustomobject]@{ Id=$Target.Id; Name=$Target.Name; Category=$Target.Category; Root=$root; Risk=$Target.Risk;
        Recommended=($Target.Recommended -and -not $policyBlocked); CanSelect=(-not $Target.ScanOnly); PolicyBlocked=$policyBlocked; Note=$Target.Note; Bytes=$bytes; Size=(Format-UnjunkSize $bytes); FileCount=$files.Count;
        Files=$files.ToArray(); RunningApps=$running; Status=$status; Issues=$issues.ToArray(); SkippedLinks=$links;
        SkippedRecent=$recent; SkippedExcluded=$excluded; MinimumAgeDays=$age; FilePattern=$Target.FilePattern; Recurse=$Target.Recurse; ScanOnly=$Target.ScanOnly; Selected=$false }
}

function Invoke-UnjunkScan {
    [CmdletBinding()]
    param([ValidateRange(1,365)][int]$TempAgeDays=7, [string[]]$Exclusions=@(), [bool]$CrowdStrikeSafe=$true, $State)
    if ($null -eq $State) { $State = @{Cancel=$false; Progress=0; Message=''} }
    $normalized = @($Exclusions | ForEach-Object { Get-UnjunkFullPath $_ })
    $targets = @(Get-UnjunkTargets)
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $index = 0
    foreach ($target in $targets) {
        if ($State.Cancel) { break }
        $State.Message = "Scanning $($target.Name)"; $State.Progress = [int](100 * $index / [Math]::Max(1,$targets.Count))
        $rows.Add((Get-UnjunkTargetScan $target $TempAgeDays $normalized $State $CrowdStrikeSafe)); $index++
    }
    $State.Progress = 100
    [pscustomobject]@{ Version=41; ScannedAt=[DateTime]::UtcNow.ToString('o'); Computer=$env:COMPUTERNAME;
        UserSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value; Cancelled=[bool]$State.Cancel;
        TempAgeDays=$TempAgeDays; CrowdStrikeSafe=$CrowdStrikeSafe; Exclusions=$normalized; Targets=$rows.ToArray() }
}

function Test-UnjunkAdministrator {
    $principal=[Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-UnjunkSystemFileScan {
    param([ValidateRange(0,365)][int]$TempAgeDays=7,[string[]]$Exclusions=@(),[switch]$IncludeCaches,[switch]$IncludeLogs,$State)
    if($null -eq $State) { $State=@{Cancel=$false;Progress=0;Message=''} }
    $normalized=@($Exclusions | ForEach-Object {Get-UnjunkFullPath $_})
    $roots=@{}
    $roots[(Get-UnjunkFullPath "$env:SystemRoot\Temp")]='system-temp'
    $roots[(Get-UnjunkFullPath "$env:LOCALAPPDATA\Temp")]='system-temp'
    $roots[(Get-UnjunkFullPath "$env:SystemRoot\LiveKernelReports")]='live-kernel'
    $roots[(Get-UnjunkFullPath "$env:SystemRoot\MEMORY.DMP")]='memory-dumps'
    $roots[(Get-UnjunkFullPath "$env:SystemRoot\Minidump")]='memory-dumps'
    $roots[(Get-UnjunkFullPath "$env:SystemDrive\Autodesk")]='autodesk-installers'
    $rows=@(foreach($target in @(Get-UnjunkTargets)) {
        if($State.Cancel) { break }
        if(-not $roots.ContainsKey($target.Root)) {
            if($IncludeCaches -and $target.Risk -eq 'Cache') { $roots[$target.Root]='app-caches' }
            elseif($IncludeLogs -and $target.Risk -eq 'Diagnostics') { $roots[$target.Root]='app-logs' }
            else { continue }
        }
        # Explicit diagnostic cleanup includes all ages; temp keeps the configured cutoff.
        if($target.Risk -eq 'Diagnostics') { $target.MinimumAgeDays=0 }
        if($roots[$target.Root] -eq 'autodesk-installers') { $target.MinimumAgeDays=0 }
        $State.Message='Scanning '+$target.Name
        $row=Get-UnjunkTargetScan $target $TempAgeDays $normalized $State $false
        $row | Add-Member -NotePropertyName SystemOption -NotePropertyValue $roots[$target.Root]
        $row
    })
    [pscustomobject]@{Version=41;ScannedAt=[DateTime]::UtcNow.ToString('o');Computer=$env:COMPUTERNAME;
        UserSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value;Cancelled=[bool]$State.Cancel;
        TempAgeDays=$TempAgeDays;CrowdStrikeSafe=$false;Exclusions=$normalized;Targets=$rows}
}

function Initialize-UnjunkRebootApi {
    if(-not ('Unjunk.NativeReboot' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace Unjunk {
    public static class NativeReboot {
        [DllImport("kernel32.dll", EntryPoint="MoveFileExW", CharSet=CharSet.Unicode, ExactSpelling=true, SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool MoveFileEx(string existingPath, string newPath, uint flags);
        public static bool QueueDelete(string path) { return MoveFileEx(path, null, 4); }
    }
}
'@ -ErrorAction Stop
    }
}

function Invoke-UnjunkNativeRebootDelete([string]$Path) {
    if(-not (Test-UnjunkAdministrator)) { throw 'Run Unjunk as administrator to queue deletion at reboot.' }
    Initialize-UnjunkRebootApi
    if(-not [Unjunk.NativeReboot]::QueueDelete($Path)) {
        $code=[Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw (New-Object ComponentModel.Win32Exception $code)
    }
}

function Remove-UnjunkCandidate {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param($File, [string]$Root, [string[]]$Exclusions=@(), [switch]$QueueOnReboot)
    $outcome = [ordered]@{Path=$File.Path; Bytes=0L; QueuedBytes=0L; Status='Skipped'; Reason=''}
    try {
        $rootPath = Assert-UnjunkRoot $Root
        $path = Get-UnjunkFullPath $File.Path
        if (-not (Test-UnjunkWithin $path $rootPath)) { throw 'File is outside the scanned target.' }
        if (Test-UnjunkExcluded $path $Exclusions) { throw 'Excluded path.' }
        if (-not (Test-UnjunkNoLinks $path)) { throw 'Link or junction detected.' }
        $current = Get-Item -LiteralPath $path -Force -ErrorAction Stop
        if ($current.PSIsContainer) { throw 'Directory deletion is not supported.' }
        if ($current.Length -ne $File.Length -or $current.LastWriteTimeUtc.Ticks -ne $File.LastWriteTicks -or $current.CreationTimeUtc.Ticks -ne $File.CreationTicks) {
            throw 'File changed after scanning; scan again.'
        }
        $action=if($QueueOnReboot){'Queue this file path for permanent deletion at the next Windows restart'}else{'Permanently delete scanned file'}
        if ($PSCmdlet.ShouldProcess($path, $action)) {
            if($QueueOnReboot) {
                Invoke-UnjunkNativeRebootDelete $path
                $outcome.Status='Queued'; $outcome.QueuedBytes=[long]$File.Length
                $outcome.Reason='Pending Windows restart. Windows deletes by path and cannot recheck file changes at boot.'
            } else {
                Remove-Item -LiteralPath $path -Force -ErrorAction Stop
                $outcome.Status='Deleted'; $outcome.Bytes=[long]$File.Length
            }
        } else { $outcome.Status='Preview'; $outcome.Reason='No deletion requested.' }
    } catch { $outcome.Reason=$_.Exception.Message }
    [pscustomobject]$outcome
}

function Test-UnjunkRhinoTargetCurrent($Target) {
    if($null -eq $Target.RhinoPackage) { return $true }
    $candidates=@(Get-UnjunkRhinoPackageTargets -Root $Target.RhinoPackage.PackagesRoot)
    return @($candidates | Where-Object {$_.Id -eq $Target.Id -and -not $_.ScanOnly}).Count -eq 1
}

function Remove-UnjunkEmptyPackageDirectories($Target,$State,[string[]]$Exclusions=@()) {
    # Only empty directories under this verified old-version root; never recursive deletion.
    $root=Assert-UnjunkRoot $Target.Root
    # The old manifest may already have been deleted. Verify the preserved version instead.
    $active=Get-Content -LiteralPath $Target.RhinoPackage.ManifestPath -Raw -ErrorAction SilentlyContinue
    if(-not $active -or $active.Trim() -ne $Target.RhinoPackage.ActiveVersion) { return }
    $latestRoot=Join-Path $Target.RhinoPackage.PackageRoot $Target.RhinoPackage.LatestVersion
    if(-not (Test-UnjunkPackageManifest $latestRoot $Target.RhinoPackage.LatestVersion)) { return }
    $stack=New-Object 'System.Collections.Generic.Stack[string]'; $stack.Push($root)
    $directories=New-Object 'System.Collections.Generic.List[string]'
    while($stack.Count -and -not $State.Cancel) {
        $path=$stack.Pop()
        try {
            if(-not (Test-UnjunkWithin $path $root) -or (Test-UnjunkExcluded $path $Exclusions) -or -not (Test-UnjunkNoLinks $path)) { continue }
            $directories.Add($path)
            foreach($child in @(Get-ChildItem -LiteralPath $path -Directory -Force -ErrorAction Stop)) { $stack.Push($child.FullName) }
        } catch { }
    }
    foreach($path in @($directories | Sort-Object Length -Descending)) {
        if($State.Cancel) { break }
        try {
            # Recheck the active pointer even after files were deleted. Never touch the package parent.
            $active=(Get-Content -LiteralPath $Target.RhinoPackage.ManifestPath -Raw -ErrorAction Stop).Trim()
            if($active -eq $Target.RhinoPackage.Version -or @(Get-UnjunkRunningApps $Target).Count) { break }
            if(-not (Test-UnjunkNoLinks $path) -or -not (Test-UnjunkWithin $path $root) -or (Test-UnjunkExcluded $path $Exclusions)) { continue }
            if(@(Get-ChildItem -LiteralPath $path -Force -ErrorAction Stop).Count -eq 0) {
                [IO.Directory]::Delete($path,$false)
                [pscustomobject]@{Path=$path;Bytes=0L;QueuedBytes=0L;Status='DirectoryRemoved';Reason='Empty old plug-in version directory.'}
            }
        } catch { }
    }
}

function Remove-UnjunkEmptyAutodeskDirectories($Target,$State,[string[]]$Exclusions=@()) {
    if(-not (Test-UnjunkAutodeskTarget $Target)){return}
    $root=Assert-UnjunkRoot $Target.Root
    if(-not (Test-Path -LiteralPath $root -PathType Container)){return}
    $stack=New-Object 'System.Collections.Generic.Stack[string]';$stack.Push($root)
    $directories=New-Object 'System.Collections.Generic.List[string]'
    while($stack.Count -and -not $State.Cancel){
        $path=$stack.Pop()
        try {
            if(-not (Test-UnjunkWithin $path $root) -or (Test-UnjunkExcluded $path $Exclusions) -or
                -not (Test-UnjunkTargetFile $path $Target) -or -not (Test-UnjunkNoLinks $path)){continue}
            $directories.Add($path)
            foreach($child in @(Get-ChildItem -LiteralPath $path -Directory -Force -ErrorAction Stop)){$stack.Push($child.FullName)}
        } catch { }
    }
    foreach($path in @($directories | Sort-Object Length -Descending)){
        if($State.Cancel -or @(Get-UnjunkRunningApps $Target).Count){break}
        try {
            if(-not (Test-UnjunkWithin $path $root) -or (Test-UnjunkExcluded $path $Exclusions) -or -not (Test-UnjunkNoLinks $path)){continue}
            if(@(Get-ChildItem -LiteralPath $path -Force -ErrorAction Stop).Count -eq 0){
                [IO.Directory]::Delete($path,$false)
                [pscustomobject]@{Path=$path;Bytes=0L;QueuedBytes=0L;Status='DirectoryRemoved';Reason='Empty Autodesk extraction folder.'}
            }
        } catch { }
    }
}

function Invoke-UnjunkCleanup {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
    param([Parameter(Mandatory)]$Scan, [Parameter(Mandatory)][string[]]$TargetIds, [switch]$QueueOnReboot, [switch]$CloseApps, $State)
    if ($null -eq $State) { $State=@{Cancel=$false; Progress=0; Message=''} }
    if ($Scan.Cancelled) { throw 'The scan was cancelled. Complete a new scan before cleaning.' }
    if($QueueOnReboot -and $Scan.CrowdStrikeSafe) { throw 'Delete at reboot is disabled by CrowdStrike Safe cleaning. Switch to Standard cleaning and scan again.' }
    if($QueueOnReboot -and -not $WhatIfPreference -and -not (Test-UnjunkAdministrator)) { throw 'Run Unjunk as administrator to queue deletion at reboot.' }
    if ($Scan.UserSid -ne [Security.Principal.WindowsIdentity]::GetCurrent().User.Value -or $Scan.Computer -ne $env:COMPUTERNAME) { throw 'This scan belongs to another user or computer.' }
    $age = [DateTime]::UtcNow - [DateTime]::Parse($Scan.ScannedAt).ToUniversalTime()
    if ($age.TotalMinutes -gt 30 -or $age.TotalMinutes -lt 0) { throw 'The scan has expired. Scan again before cleaning.' }
    $catalog=@{}; foreach($target in @(Get-UnjunkTargets)) { $catalog[$target.Id]=$target }
    $selected=@($Scan.Targets | Where-Object { $TargetIds -contains $_.Id })
    if (@($TargetIds | Sort-Object -Unique).Count -ne $selected.Count) { throw 'Unknown selection in scan.' }
    # Validate every selected target before any mutation, so a catalog change cannot cause a partial, unreported run.
    foreach($row in $selected) {
        if (-not $catalog.ContainsKey($row.Id) -or $catalog[$row.Id].Root -ne $row.Root) { throw 'Target no longer matches the cleanup catalog.' }
        $null=Assert-UnjunkRoot $row.Root
        if($catalog[$row.Id].ScanOnly) { throw 'The plug-in inventory row is informational and cannot be cleaned.' }
        foreach($file in $row.Files) {
            if(-not (Test-UnjunkTargetFile $file.Path $catalog[$row.Id])) { throw 'A scanned file does not match the current target filename/depth filter. Scan again.' }
        }
        if($QueueOnReboot -and $catalog[$row.Id].Risk -notin @('Cache','Temporary')) { throw 'Delete at reboot is limited to cache and temporary-file targets. Deselect diagnostics and installers.' }
        if($Scan.CrowdStrikeSafe -and -not (Test-UnjunkSafeTarget $catalog[$row.Id])) { throw 'Selection is blocked by CrowdStrike Safe cleaning.' }
        if($Scan.CrowdStrikeSafe -and @($row.Files | Where-Object { $_.LastWriteTicks -gt [DateTime]::UtcNow.AddDays(-7).Ticks }).Count) { throw 'Safe cleaning only permits files at least seven days old.' }
    }
    $results=New-Object 'System.Collections.Generic.List[object]'
    $appClosures=New-Object 'System.Collections.Generic.List[object]'
    $seen=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $total=[long](($selected | Measure-Object FileCount -Sum).Sum); $processed=0L
    foreach($row in $selected) {
        if ($State.Cancel) { break }
        $target=$catalog[$row.Id]
        if(-not (Test-UnjunkRhinoTargetCurrent $target)) {
            $results.Add([pscustomobject]@{Path=$row.Root;Bytes=0L;QueuedBytes=0L;Status='Skipped';Reason='Plug-in package selection changed; scan again.'}); continue
        }
        $running=@(Get-UnjunkRunningApps $catalog[$row.Id])
        $autodesk=Test-UnjunkAutodeskTarget $target
        $canClose=$CloseApps -and -not $QueueOnReboot -and ($target.Risk -eq 'Cache' -or $autodesk)
        $installers=@($running | Where-Object {$_ -notin @('AdskAccessUIHost','AdskAccessServiceHost')})
        if($autodesk -and $installers.Count){
            $results.Add([pscustomobject]@{Path=$row.Root;Bytes=0L;QueuedBytes=0L;Status='Skipped';Reason=('Active installation; wait for it to finish: '+($installers -join ', '))});continue
        }
        if ($running.Count -and -not $QueueOnReboot -and -not $canClose) {
            $results.Add([pscustomobject]@{Path=$row.Root;Bytes=0L;QueuedBytes=0L;Status='Skipped';Reason=('Close the application first: '+($running -join ', '))})
            $processed+=$row.FileCount; continue
        }
        $action=if($QueueOnReboot){'Queue selected file paths for deletion at next restart (replacement files at these paths can also be deleted)'}else{'Permanently delete selected files'}
        if($canClose) { $action='Close mapped apps (force-stop if needed; unsaved work can be lost), then '+$action.ToLowerInvariant() }
        if (-not $PSCmdlet.ShouldProcess("$($row.Name): $($row.FileCount) scanned files", $action)) { continue }
        if($canClose -and $row.FileCount -gt 0) {
            foreach($closure in @(Close-UnjunkTargetApps -Target $target -State $State)) { $appClosures.Add($closure) }
            if($State.Cancel) { break }
            $running=@(Get-UnjunkRunningApps $target)
            if($running.Count) {
                $results.Add([pscustomobject]@{Path=$row.Root;Bytes=0L;QueuedBytes=0L;Status='Skipped';Reason=('App still running after close attempt: '+($running -join ', '))}); continue
            }
        }
        $lastAppCheck=[DateTime]::UtcNow
        foreach($file in $row.Files) {
            if ($State.Cancel) { break }
            if($autodesk -or ($canClose -and ([DateTime]::UtcNow-$lastAppCheck).TotalSeconds -ge 1)) {
                $lastAppCheck=[DateTime]::UtcNow
                if(@(Get-UnjunkRunningApps $target).Count) {
                    $results.Add([pscustomobject]@{Path=$row.Root;Bytes=0L;QueuedBytes=0L;Status='Skipped';Reason='Application restarted; remaining cache files were preserved.'}); break
                }
            }
            if($null -ne $target.RhinoPackage) {
                $active=(Get-Content -LiteralPath $target.RhinoPackage.ManifestPath -Raw -ErrorAction SilentlyContinue)
                if(-not $active -or $active.Trim() -ne $target.RhinoPackage.ActiveVersion -or @(Get-UnjunkRunningApps $target).Count) {
                    $results.Add([pscustomobject]@{Path=$row.Root;Bytes=0L;QueuedBytes=0L;Status='Skipped';Reason='Rhino/Yak started or the active plug-in version changed; stopped this target.'}); break
                }
            }
            $processed++; $State.Progress=[int](100*$processed/[Math]::Max(1,$total)); $State.Message="Cleaning $($row.Name) ($processed / $total)"
            if (-not $seen.Add($file.Path)) { continue }
            $results.Add((Remove-UnjunkCandidate -File $file -Root $row.Root -Exclusions $Scan.Exclusions -QueueOnReboot:$QueueOnReboot -Confirm:$false))
        }
        if($null -ne $target.RhinoPackage -and -not $State.Cancel) {
            foreach($entry in @(Remove-UnjunkEmptyPackageDirectories $target $State $Scan.Exclusions)) { $results.Add($entry) }
        }
        if($autodesk -and -not $State.Cancel){
            foreach($entry in @(Remove-UnjunkEmptyAutodeskDirectories $target $State $Scan.Exclusions)){$results.Add($entry)}
        }
    }
    $State.Progress=100
    $deletedBytes=0L; $queuedBytes=0L
    foreach($result in $results) { $deletedBytes+=$result.Bytes; $queuedBytes+=$result.QueuedBytes }
    [pscustomobject]@{ CompletedAt=[DateTime]::UtcNow.ToString('o'); Cancelled=[bool]$State.Cancel; CrowdStrikeSafe=$Scan.CrowdStrikeSafe; QueueOnReboot=[bool]$QueueOnReboot;
        DeletedFiles=@($results | Where-Object Status -eq 'Deleted').Count;
        DeletedBytes=$deletedBytes;
        QueuedFiles=@($results | Where-Object Status -eq 'Queued').Count; QueuedBytes=$queuedBytes;
        RemovedDirectories=@($results | Where-Object Status -eq 'DirectoryRemoved').Count;
        SkippedFiles=@($results | Where-Object Status -eq 'Skipped').Count; Results=$results.ToArray();
        CloseApps=[bool]$CloseApps; ClosedProcesses=@($appClosures | Where-Object Status -in @('Closed','Terminated')).Count;
        StoppedServices=@($appClosures | Where-Object Status -eq 'ServiceStopped').Count; AppClosures=$appClosures.ToArray() }
}

function Save-UnjunkReport {
    param([Parameter(Mandatory)]$Report, [string]$Kind='scan', [string]$Path)
    if (-not $Path) {
        $directory=Join-Path $env:LOCALAPPDATA 'Unjunk\Reports'
        $null=New-Item -ItemType Directory -Path $directory -Force -WhatIf:$false -Confirm:$false
        $Path=Join-Path $directory ("{0}-{1}-{2}.json" -f $Kind,(Get-Date -Format 'yyyyMMdd-HHmmss'),[Guid]::NewGuid().ToString('N').Substring(0,6))
    }
    $Report | ConvertTo-Json -Depth 9 | Set-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop -WhatIf:$false -Confirm:$false
    return $Path
}

function Test-UnjunkRowSelectable($Row,[bool]$QueueOnReboot,[bool]$CloseApps=$false) {
    if(-not $Row.CanSelect -or $Row.PolicyBlocked -or $Row.FileCount -eq 0 -or $Row.Issues.Count) { return $false }
    if($QueueOnReboot) { return $Row.Risk -in @('Cache','Temporary') }
    return $Row.RunningApps.Count -eq 0 -or ($CloseApps -and ($Row.Risk -eq 'Cache' -or
        ((Test-UnjunkAutodeskTarget $Row) -and @($Row.RunningApps | Where-Object {$_ -notin @('AdskAccessUIHost','AdskAccessServiceHost')}).Count -eq 0)))
}

function Set-UnjunkCategorySelection {
    param([object[]]$Rows=@(),[string]$Category='All categories',[bool]$Selected=$true,[switch]$QueueOnReboot,[switch]$CloseApps)
    foreach($row in $Rows) {
        if($Category -ne 'All categories' -and $row.Category -ne $Category) { continue }
        $row.Selected=$Selected -and (Test-UnjunkRowSelectable $row ([bool]$QueueOnReboot) ([bool]$CloseApps))
    }
}

function Get-UnjunkCategoryState {
    param([object[]]$Rows=@(),[string]$Category='All categories',[switch]$QueueOnReboot,[switch]$CloseApps)
    $matching=@($Rows | Where-Object {$Category -eq 'All categories' -or $_.Category -eq $Category})
    $eligible=@($matching | Where-Object {Test-UnjunkRowSelectable $_ ([bool]$QueueOnReboot) ([bool]$CloseApps)})
    $selected=@($matching | Where-Object Selected)
    $checked=$false
    if($selected.Count) {
        $checked=if($eligible.Count -eq $selected.Count -and @($eligible | Where-Object {-not $_.Selected}).Count -eq 0){$true}else{$null}
    }
    [pscustomobject]@{Checked=$checked;CanToggle=($eligible.Count -gt 0 -or $selected.Count -gt 0);SelectedCount=$selected.Count;EligibleCount=$eligible.Count}
}

Export-ModuleMember -Function Format-UnjunkSize,Get-UnjunkTargets,Invoke-UnjunkScan,Invoke-UnjunkCleanup,Save-UnjunkReport,Set-UnjunkCategorySelection,Get-UnjunkCategoryState,Invoke-UnjunkSystemFileScan
