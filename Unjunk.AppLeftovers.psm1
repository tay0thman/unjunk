#requires -Version 5.1
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Core.psm1')

function Test-UnjunkProtectedAppName([string]$Name) {
    # Keep everyday terminal/OpenAI apps even if their registration is being repaired.
    return $Name -match '(?i)(^OpenAI\.|(^|\.)ChatGPT([.-]|$)|^Microsoft\.WindowsTerminal)' -or
        $Name -match '(?i)^(Microsoft\.(WindowsStore|StorePurchaseApp|DesktopAppInstaller|SecHealthUI|AAD\.|AccountsControl|Windows\.(ShellExperienceHost|StartMenuExperienceHost|CloudExperienceHost|Search|Cortana|Apprep|CredDialogHost|AssignedAccessLockApp|CapturePicker|FilePicker|PinningConfirmationDialog)|WindowsAppRuntime|VCLibs|UI\.Xaml|NET\.Native|Client\.)|Windows\.|MicrosoftWindows\.(Client\.|UndockedDevKit))'
}

function Test-UnjunkPackageInUse($Package) {
    foreach($flag in @('NonRemovable','IsFramework','IsResourcePackage','IsPartiallyStaged')) {
        if($null -ne $Package.PSObject.Properties[$flag] -and $Package.$flag) { return $true }
    }
    # Only a known staged-only package is unused. Missing/unknown registration state is preserved.
    if($null -eq $Package.PSObject.Properties['PackageUserInformation']) { return $true }
    $users=@($Package.PackageUserInformation)
    if(-not $users.Count) { return $true }
    foreach($user in $users) { if([string]$user.InstallState -ne 'Staged'){return $true} }
    return $false
}

function Get-UnjunkAppPresence {
    $packages=@(Get-AppxPackage -AllUsers -PackageTypeFilter Main,Bundle,Framework,Resource -ErrorAction Stop)
    if(-not $packages.Count) { throw 'Windows returned an empty machine-wide app inventory. App leftovers are preserved.' }
    $inUse=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach($package in $packages) {
        if(Test-UnjunkPackageInUse $package){$null=$inUse.Add([string]$package.Name)}
    }
    [pscustomobject]@{InUse=$inUse;Provisioned=@(Get-AppxProvisionedPackage -Online -ErrorAction Stop)}
}

function Test-UnjunkAppStillAbsent([string]$Name) {
    if(Test-UnjunkProtectedAppName $Name){return $false}
    $packages=@(Get-AppxPackage -AllUsers -Name $Name -PackageTypeFilter Main,Bundle,Framework,Resource -ErrorAction Stop)
    return @($packages | Where-Object {Test-UnjunkPackageInUse $_}).Count -eq 0
}

function Get-UnjunkAppProfileRoots {
    @(Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop | Where-Object {-not $_.Special -and $_.LocalPath} | ForEach-Object {
        $profile=[IO.Path]::GetFullPath([string]$_.LocalPath).TrimEnd('\')
        if($profile -notmatch '^[a-zA-Z]:\\' -or $profile.Substring(2).Contains(':')) { throw 'A user profile is not on a local drive.' }
        $root=Join-Path $profile 'AppData\Local\Packages'
        [pscustomobject]@{Root=$root;Sid=[string]$_.SID;Profile=$profile}
    })
}

function ConvertTo-UnjunkAppFamily([string]$PackageName) {
    $parts=$PackageName.Split('_'); $version=$null
    if($parts.Count -ne 5 -or -not [Version]::TryParse($parts[1],[ref]$version)){return ''}
    $family=$parts[0]+'_'+$parts[4]
    if($family -notmatch '^[A-Za-z0-9][A-Za-z0-9.-]{2,49}_[a-z0-9]{13}$'){return ''}
    $family
}

function Get-UnjunkAppLeftovers {
    param([string[]]$Exclusions=@(),$State)
    if($null -eq $State){$State=@{Cancel=$false;Progress=0;Message=''}}
    $presence=Get-UnjunkAppPresence
    $core=Get-Module Unjunk.Core
    $normalized=@(& $core {param($paths) foreach($path in $paths){Get-UnjunkFullPath $path}} $Exclusions)
    $stamp=[DateTime]::UtcNow.ToString('o'); $sid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    foreach($profile in @(Get-UnjunkAppProfileRoots)) {
        if($State.Cancel){return}
        if(-not (Test-Path -LiteralPath $profile.Root)){continue}
        if(-not (& $core {param($root) Test-UnjunkNoLinks $root} $profile.Root)){continue}
        foreach($folder in @(Get-ChildItem -LiteralPath $profile.Root -Directory -Force -ErrorAction Stop)) {
            if($State.Cancel){return}
            if($folder.Name -notmatch '^(?<app>[A-Za-z0-9][A-Za-z0-9.-]{2,49})_[a-z0-9]{13}$'){continue}
            $name=$Matches.app
            if((Test-UnjunkProtectedAppName $name) -or $presence.InUse.Contains($name)){continue}
            if(-not (& $core {param($root) Test-UnjunkNoLinks $root} $folder.FullName)){continue}
            $State.Message='Scanning leftover app data: '+$name
            $row=& $core {param($name,$root,$exclusions,$state)
                $target=New-UnjunkTarget $name 'App leftovers' $root -Risk Review -Note 'Residual app data; no account has this app registered.'
                Get-UnjunkTargetScan $target 0 $exclusions $state $false
            } $name $folder.FullName $normalized $State
            # Partial/inaccessible data needs a fresh successful scan, never a broad directory delete.
            if($row.Issues.Count){throw ($name+': leftover scan incomplete. '+($row.Issues -join '; '))}
            [pscustomobject]@{Id='leftover-data:'+$row.Id;Name=$name;Version='Residual data';Kind='DataLeftover';PackageName=$folder.Name;
                Scope=('Profile: '+(Split-Path $profile.Profile -Leaf));Root=$folder.FullName;ProfileSid=$profile.Sid;
                Files=@($row.Files);FileCount=$row.FileCount;Bytes=$row.Bytes;Size=$row.Size;Exclusions=$normalized;
                ScannedAt=$stamp;UserSid=$sid;Computer=$env:COMPUTERNAME;Selected=$false}
        }
    }
    foreach($package in $presence.Provisioned) {
        $name=[string]$package.DisplayName
        if($State.Cancel){return}
        if((Test-UnjunkProtectedAppName $name) -or $presence.InUse.Contains($name)){continue}
        $family=ConvertTo-UnjunkAppFamily ([string]$package.PackageName)
        if(-not $family -or -not $family.StartsWith($name+'_',[StringComparison]::OrdinalIgnoreCase)){continue}
        [pscustomobject]@{Id='leftover-provisioned:'+$package.PackageName;Name=$name;Version=[string]$package.Version;Kind='ProvisionedLeftover';
            PackageName=[string]$package.PackageName;Scope='Unused machine provisioning';Root='';ProfileSid='';Files=@();FileCount=0;Bytes=0L;Size='Windows-managed';Exclusions=$normalized;
            ScannedAt=$stamp;UserSid=$sid;Computer=$env:COMPUTERNAME;Selected=$false}
    }
}

function Invoke-UnjunkAppLeftoverCleanup {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param([Parameter(Mandatory)]$App,$State)
    if($null -eq $State){$State=@{Cancel=$false;Progress=0;Message=''}}
    if($App.Kind -notin @('DataLeftover','ProvisionedLeftover') -or (Test-UnjunkProtectedAppName $App.Name)){throw 'This is not an eligible app leftover.'}
    $age=[DateTime]::UtcNow-[DateTime]::Parse($App.ScannedAt).ToUniversalTime()
    if($App.UserSid -ne [Security.Principal.WindowsIdentity]::GetCurrent().User.Value -or $App.Computer -ne $env:COMPUTERNAME -or $age.TotalMinutes -lt 0 -or $age.TotalMinutes -gt 30){throw 'App-leftover scan is stale or belongs to another user/machine.'}
    $presence=Get-UnjunkAppPresence
    if($presence.InUse.Contains([string]$App.Name)){throw 'This app is registered to a user. Its packages and data are preserved.'}
    $empty=[pscustomobject]@{DeletedFiles=0;DeletedBytes=0L;SkippedFiles=0;QueuedFiles=0;QueuedBytes=0L;Results=@()}
    if($App.Kind -eq 'ProvisionedLeftover') {
        $current=@($presence.Provisioned | Where-Object {$_.PackageName -eq $App.PackageName -and $_.DisplayName -eq $App.Name})
        if($current.Count -ne 1 -or -not (ConvertTo-UnjunkAppFamily $App.PackageName)){throw 'Provisioned leftover changed. Scan again.'}
        if(-not (Test-UnjunkAppStillAbsent $App.Name)){throw 'App was installed after scanning; preserved.'}
        $count=0L; $restart=$false
        if(-not $State.Cancel -and $PSCmdlet.ShouldProcess($App.PackageName,'Remove unused machine-wide provisioning; preserve installed apps')){
            $result=Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName -ErrorAction Stop
            $count=1L; $restart=[bool]$result.RestartNeeded
        }
        return [pscustomobject]@{ExitCode=0;Output="Removed $count unused provisioned copy. No installed app was uninstalled.";RestartRequired=$restart;ItemsCleaned=$count;FileResult=$null;RemovedProvisionedPackages=$count}
    }
    $core=Get-Module Unjunk.Core
    $root=& $core {param($path) Assert-UnjunkRoot $path} $App.Root
    $profile=@(Get-UnjunkAppProfileRoots | Where-Object {$_.Sid -eq $App.ProfileSid -and [IO.Path]::GetDirectoryName($root) -eq $_.Root})
    if($profile.Count -ne 1 -or (Split-Path $root -Leaf) -ne $App.PackageName -or $App.PackageName -notmatch '^[A-Za-z0-9][A-Za-z0-9.-]{2,49}_[a-z0-9]{13}$' -or $App.PackageName.Split('_')[0] -ne $App.Name){throw 'App data does not match a verified local profile and package family.'}
    foreach($file in $App.Files){if(-not (& $core {param($p,$r) Test-UnjunkWithin (Get-UnjunkFullPath $p) $r} $file.Path $root)){throw 'A scanned leftover file is outside its package directory.'}}
    if(-not $PSCmdlet.ShouldProcess($root,'Remove only scanned leftovers of an app absent from every account')){
        return [pscustomobject]@{ExitCode=0;Output='Preview only.';RestartRequired=$false;ItemsCleaned=0L;FileResult=$empty;RemovedProvisionedPackages=0L}
    }
    $results=New-Object 'System.Collections.Generic.List[object]'
    foreach($file in $App.Files) {
        if($State.Cancel){break}
        try {
            if(-not (Test-UnjunkAppStillAbsent $App.Name)){throw 'App registration appeared during cleanup; remaining data preserved.'}
        } catch {
            $results.Add([pscustomobject]@{Path=$root;Bytes=0L;Status='Skipped';Reason=$_.Exception.Message}); break
        }
        $results.Add((& $core {param($file,$root,$exclusions) Remove-UnjunkCandidate -File $file -Root $root -Exclusions $exclusions -Confirm:$false} $file $root $App.Exclusions))
    }
    # Leave empty directory shells intact. Never sweep new/unscanned data with recursive deletion.
    $files=@($results | Where-Object Status -eq 'Deleted').Count; $bytes=[long](($results | Measure-Object Bytes -Sum).Sum)
    $fileResult=[pscustomobject]@{DeletedFiles=$files;DeletedBytes=$bytes;SkippedFiles=@($results | Where-Object Status -eq 'Skipped').Count;QueuedFiles=0;QueuedBytes=0L;Results=$results.ToArray()}
    [pscustomobject]@{ExitCode=0;Output="Removed $files leftover files / $(Format-UnjunkSize $bytes). Registered apps are preserved.";RestartRequired=$false;ItemsCleaned=[long]$files;FileResult=$fileResult;RemovedProvisionedPackages=0L}
}

Export-ModuleMember -Function Get-UnjunkAppLeftovers,Invoke-UnjunkAppLeftoverCleanup,Test-UnjunkProtectedAppName
