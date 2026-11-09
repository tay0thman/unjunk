#requires -Version 5.1
# Mock native launches only. Never starts servicing or Disk Cleanup.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Windows.psm1') -Force
$module=Get-Module Unjunk.Windows;$passed=0
$fixture=Join-Path $PSScriptRoot ('ForegroundFixture-'+[Guid]::NewGuid().ToString('N'))
$previousLocal=$env:LOCALAPPDATA;$env:LOCALAPPDATA=$fixture
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
function Reject([scriptblock]$Action,[string]$Message){$failed=$false;try{& $Action | Out-Null}catch{$failed=$true};Assert $failed $Message}
try {
    & $module {
        $script:maintenanceMutexName='Local\Unjunk-ForegroundFixture-'+[Guid]::NewGuid().ToString('N')
        $script:calls=New-Object 'System.Collections.Generic.List[object]';$script:code=0;$script:missingResult=$false;$script:mismatch=$false
        function script:Test-UnjunkWindowsAdmin {$true}
        function script:Start-Process {
            param($FilePath,$ArgumentList,$WindowStyle,[switch]$PassThru,$RedirectStandardOutput,$RedirectStandardError,$ErrorAction)
            $script:calls.Add([pscustomobject]@{FilePath=$FilePath;Arguments=@($ArgumentList);Style=$WindowStyle;Redirected=[bool]($RedirectStandardOutput -or $RedirectStandardError)})
            if($ArgumentList -contains '-ResultPath'){
                $resultPath=$ArgumentList[([array]::IndexOf($ArgumentList,'-ResultPath')+1)].Trim('"')
                $outPath=$ArgumentList[([array]::IndexOf($ArgumentList,'-OutputPath')+1)].Trim('"')
                Set-Content -LiteralPath $outPath -Value 'DISM progress fixture'
                if(-not $script:missingResult){@{ExitCode=if($script:mismatch){5}else{$script:code}} | ConvertTo-Json | Set-Content -LiteralPath $resultPath}
            }
            if($RedirectStandardOutput){Set-Content -LiteralPath $RedirectStandardOutput -Value 'Native output fixture'}
            if($RedirectStandardError){Set-Content -LiteralPath $RedirectStandardError -Value ''}
            $p=[pscustomobject]@{Handle=123;Id=456;ExitCode=$script:code;Polls=0;Disposed=$false}
            $p | Add-Member ScriptMethod WaitForExit {param($Milliseconds) if($null -eq $Milliseconds){return};$this.Polls++;return $this.Polls -ge 2}
            $p | Add-Member ScriptMethod Dispose {$this.Disposed=$true}
            $script:lastProcess=$p
            return $p
        }
    }
    $state=@{Cancel=$false;Message='';StatusPath=(Join-Path $fixture 'status.json')}
    $result=& $module {param($s) Invoke-UnjunkWindowsNative Dism.exe @('/Online','/Cleanup-Image','/StartComponentCleanup','/NoRestart','/English') $s} $state
    $call=& $module {$script:calls[-1]}
    Assert ($call.Style -eq 'Normal' -and -not $call.Redirected -and 'StartComponentCleanup' -in $call.Arguments -and $call.FilePath -like '*powershell.exe') 'DISM opens a normal visible progress console without redirecting the console itself'
    Assert ($result.ExitCode -eq 0 -and $result.Output -match 'progress fixture' -and $result.Foreground -and (& $module {$script:lastProcess.Disposed})) 'DISM output, verified completion and disposal are retained'
    Assert ($state.Message -match 'foreground window' -and (Get-Content -LiteralPath $state.StatusPath -Raw) -match 'foreground window') 'UI and detached-worker status identify the foreground task'
    $result=& $module {param($s) Invoke-UnjunkWindowsNative cleanmgr.exe @('/sagerun:9000') $s} $state
    $call=& $module {$script:calls[-1]}
    Assert ($call.Style -eq 'Normal' -and -not $call.Redirected -and $call.FilePath -like '*cleanmgr.exe' -and $call.Arguments[0] -eq '/sagerun:9000') 'Disk Cleanup uses its visible native dialog and keeps the selected cleanup profile'
    Assert ($result.ExitCode -eq 0 -and $result.Foreground) 'Foreground Disk Cleanup completion is returned to the sequence'
    $result=& $module {param($s) Invoke-UnjunkWindowsNative vssadmin.exe @('delete','shadows','/Shadow={fixture}','/Quiet') $s} $state
    $call=& $module {$script:calls[-1]}
    Assert ($call.Style -eq 'Hidden' -and $call.Redirected -and -not $result.Foreground) 'Short snapshot commands keep their existing noninteractive behavior'
    & $module {$script:code=3010}
    $result=Invoke-UnjunkComponentAnalysis -State $state
    $call=& $module {$script:calls[-1]}
    Assert ('AnalyzeComponentStore' -in $call.Arguments -and $result.RestartRequired -and $result.CompletedAt) 'Read-only WinSxS analysis uses the same visible console and preserves native exit codes'
    & $module {$script:code=0;$script:missingResult=$true}
    Reject {& $module {Invoke-UnjunkWindowsNative Dism.exe @('/StartComponentCleanup')}} 'Closing the DISM window without a result cannot be reported as success'
    & $module {$script:missingResult=$false;$script:mismatch=$true}
    Reject {& $module {Invoke-UnjunkWindowsNative Dism.exe @('/StartComponentCleanup')}} 'Conflicting wrapper/native exit codes are rejected'
    & $module {$script:mismatch=$false;$script:code=$null}
    Reject {& $module {Invoke-UnjunkWindowsNative cleanmgr.exe @('/sagerun:9000')}} 'A missing native exit code is never converted to success'
    Write-Host "All $passed foreground checks passed."
} finally {
    $env:LOCALAPPDATA=$previousLocal
    $path=[IO.Path]::GetFullPath($fixture);$parent=[IO.Path]::GetFullPath($PSScriptRoot)+'\'
    if(-not $path.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $path -Leaf) -notlike 'ForegroundFixture-*'){throw 'Unsafe fixture path'}
    if(Test-Path -LiteralPath $path){Remove-Item -LiteralPath $path -Recurse -Force}
}
