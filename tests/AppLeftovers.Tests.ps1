#requires -Version 5.1
# Real disposable file fixtures, mocked Appx/CIM providers. Never removes real packages.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.AppLeftovers.psm1') -Force
$module=Get-Module Unjunk.AppLeftovers
$passed=0
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
function Reject([scriptblock]$Action,[string]$Message){$failed=$false;try{& $Action | Out-Null}catch{$failed=$true};Assert $failed $Message}
function Package($Name,$State='Installed',[bool]$Framework=$false){[pscustomobject]@{Name=$Name;NonRemovable=$false;IsFramework=$Framework;IsResourcePackage=$false;IsPartiallyStaged=$false;PackageUserInformation=@([pscustomobject]@{InstallState=$State})}}
$fixture=Join-Path $PSScriptRoot ('AppLeftoversFixture-'+[Guid]::NewGuid().ToString('N'))
$profile1=Join-Path $fixture 'UserOne';$profile2=Join-Path $fixture 'UserTwo'
$family='Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe'
$data1=Join-Path $profile1 ('AppData\Local\Packages\'+$family)
$data2=Join-Path $profile2 ('AppData\Local\Packages\'+$family)
$activeData=Join-Path $profile1 'AppData\Local\Packages\Contoso.Active_8wekyb3d8bbwe'
$terminalData=Join-Path $profile1 'AppData\Local\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe'
$chatData=Join-Path $profile1 'AppData\Local\Packages\OpenAI.ChatGPT-Desktop_2p2nqsd0c76g0'
$codexData=Join-Path $profile1 'AppData\Local\Packages\OpenAI.Codex_2p2nqsd0c76g0'
$null=New-Item -ItemType Directory -Path $data1,$data2,$activeData,$terminalData,$chatData,$codexData -Force
foreach($directory in @($data1,$data2,$activeData,$terminalData,$chatData,$codexData)){Set-Content -LiteralPath (Join-Path $directory 'data.bin') -Value 'fixture'}
$packages=@((Package 'Microsoft.WindowsStore'),(Package 'Contoso.Active'),(Package 'Microsoft.MicrosoftStickyNotes' 'Staged'),(Package 'Contoso.Framework' 'Staged' $true))
& $module {
    param($packages,$profile1,$profile2)
    $script:packages=$packages; $script:profile1=$profile1; $script:profile2=$profile2
    $script:failInventory=$false;$script:emptyInventory=$false;$script:reinstall=$false;$script:targetedReads=0;$script:reinstallAfter=0
    $script:removed=New-Object 'System.Collections.Generic.List[string]'
    function script:Get-AppxPackage {
        param([switch]$AllUsers,$Name,$PackageTypeFilter,$ErrorAction)
        if(-not $AllUsers){throw 'Fixture requires machine-wide inventory'}
        if($script:failInventory){throw 'Fixture app inventory failure'}
        if($script:emptyInventory){return}
        $packages=@($script:packages)
        if($Name){$script:targetedReads++}
        if($script:reinstall -or ($script:reinstallAfter -gt 0 -and $script:targetedReads -ge $script:reinstallAfter)){
            $packages+=@([pscustomobject]@{Name='Microsoft.MicrosoftStickyNotes';PackageUserInformation=@([pscustomobject]@{InstallState='Installed'})})
        }
        if($Name){$packages | Where-Object Name -eq $Name}else{$packages}
    }
    function script:Get-AppxProvisionedPackage {
        param([switch]$Online,$ErrorAction)
        foreach($name in @('Microsoft.MicrosoftStickyNotes','Microsoft.WindowsTerminal','OpenAI.ChatGPT-Desktop','Contoso.Active','Contoso.Framework')){
            [pscustomobject]@{DisplayName=$name;Version='1.0.0.0';PackageName=$name+'_1.0.0.0_neutral_~_8wekyb3d8bbwe'}
        }
    }
    function script:Get-CimInstance {
        param($ClassName,$ErrorAction)
        @([pscustomobject]@{LocalPath=$script:profile1;SID='S-1-5-21-100-1001';Special=$false},[pscustomobject]@{LocalPath=$script:profile2;SID='S-1-5-21-100-1002';Special=$false})
    }
    function script:Remove-AppxProvisionedPackage {param([switch]$Online,$PackageName,$ErrorAction) $script:removed.Add($PackageName);[pscustomobject]@{RestartNeeded=$false}}
    function script:Remove-AppxPackage {throw 'Installed app uninstallation must never be called'}
} $packages $profile1 $profile2
try {
    $apps=@(Get-UnjunkAppLeftovers)
    Assert ($apps.Count -eq 3) 'Inventory offers two local-profile leftovers and one unused machine provisioned copy'
    Assert (@($apps | Where-Object Name -ne 'Microsoft.MicrosoftStickyNotes').Count -eq 0) 'Active apps, shared frameworks, Terminal and ChatGPT are preserved'
    Assert (@($apps | Where-Object Kind -eq 'DataLeftover').Count -eq 2) 'Machine-wide leftover discovery covers all local user profiles'
    Assert (@($apps | Where-Object Kind -eq 'ProvisionedLeftover').Count -eq 1) 'A staged-only app can have its unused provisioning cleaned'
    foreach($name in @('OpenAI.ChatGPT-Desktop','OpenAI.Codex','9E2F88E3.ChatGPT','Microsoft.WindowsTerminal','Microsoft.WindowsTerminalPreview')){Assert (Test-UnjunkProtectedAppName $name) ('Always protect '+$name)}
    $app=@($apps | Where-Object Root -eq $data1)[0]
    $preview=Invoke-UnjunkAppLeftoverCleanup -App $app -WhatIf
    Assert ((Test-Path -LiteralPath (Join-Path $data1 'data.bin')) -and $preview.ItemsCleaned -eq 0) 'WhatIf does not remove app data'
    $provisioned=@($apps | Where-Object Kind -eq 'ProvisionedLeftover')[0]
    $null=Invoke-UnjunkAppLeftoverCleanup -App $provisioned -WhatIf
    Assert ((& $module {$script:removed.Count}) -eq 0) 'WhatIf never removes provisioning'
    $installed=$app.PSObject.Copy();$installed.Kind='Installed'
    Reject {Invoke-UnjunkAppLeftoverCleanup -App $installed -Confirm:$false} 'Old installed-app selections are rejected by the engine'
    $escaped=$app.PSObject.Copy();$escaped.Root=$fixture
    Reject {Invoke-UnjunkAppLeftoverCleanup -App $escaped -Confirm:$false} 'A broad or changed profile root is rejected'
    $other=$app.PSObject.Copy();$other.UserSid='Other'
    Reject {Invoke-UnjunkAppLeftoverCleanup -App $other -Confirm:$false} 'Cross-account replay is rejected'
    & $module {$script:reinstall=$true}
    Reject {Invoke-UnjunkAppLeftoverCleanup -App $app -Confirm:$false} 'Reinstalling an app for any account preserves its data'
    Reject {Invoke-UnjunkAppLeftoverCleanup -App $provisioned -Confirm:$false} 'Reinstalled apps cannot have their provisioning removed'
    & $module {$script:reinstall=$false;$script:failInventory=$true}
    Reject {Get-UnjunkAppLeftovers} 'Failed machine app inventory cannot become an empty uninstall list'
    Reject {Invoke-UnjunkAppLeftoverCleanup -App $app -Confirm:$false} 'Inventory failure before deletion preserves scanned data'
    & $module {$script:failInventory=$false;$script:emptyInventory=$true}
    Reject {Get-UnjunkAppLeftovers} 'An unexpectedly empty machine inventory fails closed'
    & $module {$script:emptyInventory=$false}
    $path=Join-Path $data1 'data.bin';$app.Exclusions=@($path)
    $result=Invoke-UnjunkAppLeftoverCleanup -App $app -Confirm:$false
    Assert ($result.FileResult.SkippedFiles -eq 1 -and (Test-Path -LiteralPath $path)) 'File exclusions are rechecked at execution'
    $app.Exclusions=@()
    Set-Content -LiteralPath $path -Value 'changed after scan'
    $result=Invoke-UnjunkAppLeftoverCleanup -App $app -Confirm:$false
    Assert ($result.FileResult.DeletedFiles -eq 0 -and (Test-Path -LiteralPath $path)) 'Files changed after scanning are preserved'
    Set-Content -LiteralPath (Join-Path $data1 'second.bin') -Value 'second'
    $app=@(Get-UnjunkAppLeftovers | Where-Object Root -eq $data1)[0]
    Set-Content -LiteralPath (Join-Path $data1 'unscanned.bin') -Value 'new'
    & $module {$script:targetedReads=0;$script:reinstallAfter=2}
    $result=Invoke-UnjunkAppLeftoverCleanup -App $app -Confirm:$false
    Assert ($result.FileResult.DeletedFiles -eq 1 -and $result.FileResult.SkippedFiles -eq 1) 'A registration appearing between files stops leftover cleanup'
    Assert (Test-Path -LiteralPath (Join-Path $data1 'unscanned.bin')) 'New unscanned data is never swept away'
    & $module {$script:reinstallAfter=0}
    $second=@(Get-UnjunkAppLeftovers | Where-Object Root -eq $data2)[0]
    $result=Invoke-UnjunkAppLeftoverCleanup -App $second -Confirm:$false
    Assert ($result.ItemsCleaned -eq 1 -and $result.FileResult.DeletedFiles -eq 1) 'Eligible leftover files in another profile can be cleaned machine-wide'
    $result=Invoke-UnjunkAppLeftoverCleanup -App $provisioned -Confirm:$false
    Assert ($result.RemovedProvisionedPackages -eq 1 -and (& $module {$script:removed.Count}) -eq 1) 'Only the exact unused provisioned copy is removed'
    Assert ((Test-Path -LiteralPath (Join-Path $activeData 'data.bin')) -and (Test-Path -LiteralPath (Join-Path $terminalData 'data.bin')) -and (Test-Path -LiteralPath (Join-Path $chatData 'data.bin'))) 'Every protected or active app data fixture survives all cleanup cases'
} finally {
    $resolved=[IO.Path]::GetFullPath($fixture);$parent=[IO.Path]::GetFullPath($PSScriptRoot)+'\'
    if(-not $resolved.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'AppLeftoversFixture-*'){throw 'Unsafe fixture cleanup path'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
Write-Host "All $passed app-leftover checks passed."
