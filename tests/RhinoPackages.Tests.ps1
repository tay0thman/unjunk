#requires -Version 5.1
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Core.psm1') -Force
$module=Get-Module Unjunk.Core
$root=Join-Path $env:TEMP ('Unjunk-package-tests-'+[Guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $root
$passed=0
function Assert([bool]$Condition,[string]$Message) { if(-not $Condition){throw "FAIL: $Message"}; $script:passed++; Write-Host "PASS: $Message" }
function Package([string]$Name,[string[]]$Versions,[string]$Active) {
    $path=Join-Path $root $Name
    $null=New-Item -ItemType Directory -Path $path
    foreach($version in $Versions) {
        $folder=Join-Path $path $version; $null=New-Item -ItemType Directory -Path $folder
        Set-Content -LiteralPath (Join-Path $folder 'manifest.yml') -Value @("name: $Name","version: $version")
        Set-Content -LiteralPath (Join-Path $folder 'plugin.dll') -Value 'disposable test fixture'
    }
    if($Active) { Set-Content -LiteralPath (Join-Path $path 'manifest.txt') -Value $Active }
}
function Targets { @(& $module {param($r) Get-UnjunkRhinoPackageTargets -Root $r} $root) }
function ScanRow($Target) { & $module {param($t) Get-UnjunkTargetScan $t 7 @() @{Cancel=$false} $false} $Target }
function Envelope($Row) { [pscustomobject]@{Cancelled=$false;UserSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value;Computer=$env:COMPUTERNAME;ScannedAt=[DateTime]::UtcNow.ToString('o');Targets=@($Row);Exclusions=@();CrowdStrikeSafe=$false} }
$originalCatalog=& $module {(Get-Item Function:Get-UnjunkTargets).ScriptBlock}
try {
    Package Numeric @('1.2.0','1.9.0','1.10.0') '1.10.0'
    Package Downgraded @('1.0.0','2.0.0','3.0.0') '1.0.0'
    Package Prerelease @('2.0.0','3.0.0-beta.2','3.0.0-beta.10') '3.0.0-beta.10'
    Package FourPart @('2025.5.5.0','2025.5.10.0') '2025.5.10.0'
    Package Builds @('1.0.0+one','1.0.0+two') '1.0.0+two'
    Package Missing @('1.0.0','2.0.0') ''
    Package BrokenPointer @('1.0.0','2.0.0') '99.0.0'
    Package Unknown @('backup','1.0.0') '1.0.0'
    $all=Targets; $old=@($all | Where-Object {-not $_.ScanOnly})
    Assert (@($old | Where-Object Root -like '*\Numeric\*').Count -eq 2) 'Numeric version ordering treats 1.10 as newer than 1.9'
    $downgraded=@($old | Where-Object Root -like '*\Downgraded\*')
    Assert ($downgraded.Count -eq 1 -and $downgraded[0].RhinoPackage.Version -eq '2.0.0') 'Active downgrade and newest version are both preserved'
    $pre=@($old | Where-Object Root -like '*\Prerelease\*')
    Assert ($pre.Count -eq 1 -and $pre[0].RhinoPackage.Version -eq '3.0.0-beta.2') 'Prerelease numeric identifiers sort correctly and newest stable is preserved'
    Assert (@($old | Where-Object Root -like '*\FourPart\2025.5.5.0').Count -eq 1) 'Four-part Yak versions are supported'
    Assert (@($old | Where-Object Root -like '*\Builds\*').Count -eq 0) 'Equal versions differing only in build metadata are preserved'
    Assert (@($old | Where-Object {$_.Root -match '\\(Missing|BrokenPointer|Unknown)\\'}).Count -eq 0) 'Missing pointers, broken active versions and unknown folders are preserved'
    Assert (@($all | Where-Object ScanOnly).Count -eq 1) 'An informational package inventory is always returned'
    Assert (@($old | Where-Object Recommended).Count -eq 0) 'Old plug-in versions require explicit selection'
    $selected=@($old | Where-Object Root -like '*\Numeric\1.2.0')[0]
    $row=ScanRow $selected; $envelope=Envelope $row
    & $module {param($r) $script:packageTestRoot=$r; function script:Get-UnjunkTargets {Get-UnjunkRhinoPackageTargets -Root $script:packageTestRoot}} $root
    Set-Content -LiteralPath (Join-Path $root 'Numeric\manifest.txt') -Value '1.2.0'
    $caught=$false; try { Invoke-UnjunkCleanup -Scan $envelope -TargetIds @($selected.Id) -Confirm:$false } catch {$caught=$true}
    Assert ($caught -and (Test-Path -LiteralPath (Join-Path $selected.Root 'plugin.dll'))) 'An active-version change after scanning blocks cleanup'
    Set-Content -LiteralPath (Join-Path $root 'Numeric\manifest.txt') -Value '1.10.0'
    $preview=Invoke-UnjunkCleanup -Scan $envelope -TargetIds @($selected.Id) -WhatIf
    Assert ((Test-Path -LiteralPath $selected.Root) -and $preview.DeletedFiles -eq 0) 'WhatIf preserves plug-in files and directories'
    $result=Invoke-UnjunkCleanup -Scan $envelope -TargetIds @($selected.Id) -Confirm:$false
    Assert ($result.DeletedFiles -eq 2 -and $result.RemovedDirectories -eq 1 -and -not (Test-Path -LiteralPath $selected.Root)) 'Selected older version files and its empty folder are removed'
    Assert ((Test-Path -LiteralPath (Join-Path $root 'Numeric\1.10.0\plugin.dll')) -and (Get-Content -LiteralPath (Join-Path $root 'Numeric\manifest.txt')).Trim() -eq '1.10.0') 'Active plug-in payload and manifest pointer remain unchanged'
    $selected=@(Targets | Where-Object Root -like '*\Numeric\1.9.0')[0]; $row=ScanRow $selected; $envelope=Envelope $row
    $newFile=Join-Path $selected.Root 'new-after-scan.txt'; Set-Content -LiteralPath $newFile -Value 'keep'
    $result=Invoke-UnjunkCleanup -Scan $envelope -TargetIds @($selected.Id) -Confirm:$false
    Assert ((Test-Path -LiteralPath $newFile) -and (Test-Path -LiteralPath $selected.Root)) 'New unscanned files prevent removal of the old-version directory'
    Write-Host "All $passed Rhino package checks passed."
} finally {
    & $module {param($original) Set-Item Function:script:Get-UnjunkTargets -Value $original; Remove-Variable packageTestRoot -Scope Script -ErrorAction SilentlyContinue} $originalCatalog
    $resolved=[IO.Path]::GetFullPath($root); $prefix=[IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')+'\Unjunk-package-tests-'
    if(-not $resolved.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase)) {throw 'Test cleanup path check failed.'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
