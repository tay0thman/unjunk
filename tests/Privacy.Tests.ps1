#requires -Version 5.1
# Disposable Recent folder and in-memory registry fixture; never clears real user history.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Privacy.psm1') -Force
$module=Get-Module Unjunk.Privacy
$passed=0
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
$fixture=Join-Path $PSScriptRoot ('PrivacyFixture-'+[Guid]::NewGuid().ToString('N'))
$previousAppData=$env:APPDATA;$env:APPDATA=$fixture
$recent=Join-Path $fixture 'Microsoft\Windows\Recent'
$automatic=Join-Path $recent 'AutomaticDestinations';$custom=Join-Path $recent 'CustomDestinations'
$null=New-Item -ItemType Directory -Path $automatic,$custom -Force
foreach($name in @('private-document.lnk','second.LNK','keep.lnk','desktop.ini')){Set-Content -LiteralPath (Join-Path $recent $name) -Value 'fixture'}
$autoFile=Join-Path $automatic 'pinned.automaticDestinations-ms';$customFile=Join-Path $custom 'pinned.customDestinations-ms';$nestedLink=Join-Path $custom 'nested.lnk'
foreach($path in @($autoFile,$customFile,$nestedLink)){Set-Content -LiteralPath $path -Value 'pinned jump list fixture'}
$autoHash=(Get-FileHash -LiteralPath $autoFile).Hash;$customHash=(Get-FileHash -LiteralPath $customFile).Hash
& $module {
    param($recent)
    $script:recentFixture=$recent;$script:lateFile=$false;$script:changeFile=$false;$script:changeRegistry=$false;$script:deletedValues=0
    function script:Open-UnjunkPrivacyRegistryKey {param($Path,$Writable) return $null}
    function script:Get-UnjunkPrivacyWindowsRoot {Join-Path $env:APPDATA 'WindowsFixture'}
    function script:Get-UnjunkActivityInventory {param($Exclusions) [pscustomobject]@{Databases=@();Issues=@()}}
    $script:registry=@{History='secret-document-path';Usage='secret-app-path';StartTrackDocs=1;StartTrackProgs=1}
    function script:Get-UnjunkPrivacyEntries {
        $root='Software\Microsoft\Windows\CurrentVersion\Explorer'
        $entries=@([pscustomobject]@{Key=$root+'\RecentDocs';Name='History';Hash=(Get-UnjunkPrivacyHash $script:registry.History 'String')},
            [pscustomobject]@{Key=$root+'\UserAssist\{11111111-1111-1111-1111-111111111111}\Count';Name='Usage';Hash=(Get-UnjunkPrivacyHash $script:registry.Usage 'String')})
        if($script:lateFile){Set-Content -LiteralPath (Join-Path $script:recentFixture 'new-after-scan.lnk') -Value 'new'}
        if($script:changeFile){Set-Content -LiteralPath (Join-Path $script:recentFixture 'private-document.lnk') -Value 'changed after scan'}
        if($script:changeRegistry){$script:registry.History='new-value-after-scan'}
        $entries
    }
    function script:Remove-UnjunkPrivacyEntry($Entry){
        if(-not (Test-UnjunkPrivacyKey $Entry.Key)){throw 'Unexpected registry write'}
        if((Get-UnjunkPrivacyHash $script:registry[$Entry.Name] 'String') -ne $Entry.Hash){return $false}
        $script:registry.Remove($Entry.Name);$script:deletedValues++;return $true
    }
} $recent
try {
    $preview=Get-UnjunkPrivacyPreview -Exclusions @((Join-Path $recent 'keep.lnk'))
    Assert ($preview.RecentLinks -eq 2 -and $preview.RegistryRecords -eq 2 -and $preview.JumpListsPreserved) 'Preview counts only direct recent shortcuts and allowed registry history'
    Assert (($preview | ConvertTo-Json) -notmatch 'private-document|secret-document|secret-app') 'Privacy preview does not copy sensitive path or usage contents'
    $result=Invoke-UnjunkPrivacyCleanup -Exclusions @((Join-Path $recent 'keep.lnk')) -WhatIf
    Assert ((Test-Path -LiteralPath (Join-Path $recent 'private-document.lnk')) -and (& $module {$script:deletedValues}) -eq 0) 'WhatIf preserves file and registry history'
    foreach($path in @('Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced','Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage','Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist','Software\Microsoft\Windows\CurrentVersion\Run')){
        Assert (-not (& $module {param($path) Test-UnjunkPrivacyKey $path} $path)) ('Privacy allowlist excludes '+(Split-Path $path -Leaf))
    }
    Assert (& $module {Test-UnjunkPrivacyKey 'Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSavePidlMRU\txt'}) 'File-dialog history subkeys are eligible'
    & $module {$script:lateFile=$true;$script:changeFile=$true;$script:changeRegistry=$true}
    $result=Invoke-UnjunkPrivacyCleanup -Exclusions @((Join-Path $recent 'keep.lnk')) -Confirm:$false
    Assert ($result.FileResult.DeletedFiles -eq 1 -and $result.RemovedPrivacyRecords -eq 1 -and $result.ItemsCleaned -eq 2) 'Only unchanged recent links and registry records are counted removed'
    Assert ((Test-Path -LiteralPath (Join-Path $recent 'private-document.lnk')) -and (Test-Path -LiteralPath (Join-Path $recent 'new-after-scan.lnk'))) 'Changed and newly created shortcuts survive'
    Assert (Test-Path -LiteralPath (Join-Path $recent 'keep.lnk')) 'Excluded recent-file shortcuts survive'
    Assert ((Get-FileHash -LiteralPath $autoFile).Hash -eq $autoHash -and (Get-FileHash -LiteralPath $customFile).Hash -eq $customHash -and (Test-Path -LiteralPath $nestedLink)) 'Automatic, custom and nested Jump List files remain byte-for-byte unchanged'
    Assert ((& $module {$script:registry.StartTrackDocs}) -eq 1 -and (& $module {$script:registry.StartTrackProgs}) -eq 1) 'Jump List and application tracking preferences are preserved'
    Assert (($result | ConvertTo-Json -Depth 8) -notmatch 'private-document|secret-document|secret-app|new-value-after-scan' -and $result.FileResult.Results.Count -eq 0) 'Cleanup report does not recreate the erased privacy traces'
    Assert (Test-Path -LiteralPath (Join-Path $recent 'desktop.ini')) 'Non-shortcut Recent folder files are preserved'
} finally {
    $env:APPDATA=$previousAppData;$path=[IO.Path]::GetFullPath($fixture);$parent=[IO.Path]::GetFullPath($PSScriptRoot)+'\'
    if(-not $path.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $path -Leaf) -notlike 'PrivacyFixture-*'){throw 'Unsafe privacy fixture path'}
    Remove-Item -LiteralPath $path -Recurse -Force
}
Write-Host "All $passed privacy checks passed."
