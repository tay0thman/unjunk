#requires -Version 5.1
# Synthetic services/processes only. Never stops a real service or application.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Core.psm1') -Force
$module=Get-Module Unjunk.Core;$passed=0
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
& $module {
    $script:stopCalls=0;$script:deny=$false;$script:processLookups=0
    function script:Get-Service {
        param($Name,$ErrorAction)
        $script:serviceNames=@($Name)
        $service=[pscustomobject]@{ServiceName='Autodesk Access Service Host';Status='Running';StartType='Automatic';Disposed=$false}
        $service | Add-Member ScriptMethod Stop {$script:stopCalls++;if($script:deny){throw 'Fixture access denied'};$this.Status='Stopped'}
        $service | Add-Member ScriptMethod WaitForStatus {param($Status,$Timeout) if($this.Status -ne 'Stopped'){throw 'Fixture still running'}}
        $service | Add-Member ScriptMethod Dispose {$this.Disposed=$true}
        $script:service=$service;$service
    }
    function script:Get-Process {
        param($Name,$ErrorAction)
        $script:processLookups++;$script:processName=$Name
        $p=[pscustomobject]@{Id=50001;ProcessName='AdskAccessServiceHost';Handle=123;HasExited=$false;Killed=$false;Disposed=$false;ServiceWasStopped=$false}
        $p | Add-Member ScriptMethod Kill {$this.ServiceWasStopped=($script:service.Status -eq 'Stopped');$this.Killed=$true;$this.HasExited=$true}
        $p | Add-Member ScriptMethod WaitForExit {param($Timeout) return $this.HasExited}
        $p | Add-Member ScriptMethod Dispose {$this.Disposed=$true}
        $script:hostProcess=$p;$p
    }
}
$result=@(& $module {Stop-UnjunkAutodeskAccess -State @{Cancel=$false;Message=''}})
Assert ((& $module {$script:serviceNames -contains 'Autodesk Access Service Host' -and $script:serviceNames -notcontains 'AdskLicensingService'}) -and (& $module {$script:processName}) -eq 'AdskAccessServiceHost') 'Only Autodesk Access service names and its exact host process are targeted'
Assert ((& $module {$script:service.Status}) -eq 'Stopped' -and (& $module {$script:service.StartType}) -eq 'Automatic') 'Stopping Access leaves the startup setting unchanged'
Assert ((& $module {$script:hostProcess.Killed -and $script:hostProcess.ServiceWasStopped}) -and @($result | Where-Object Status -eq 'ServiceStopped').Count -eq 1) 'A lingering Access host can be terminated only after the service stopped'
Assert ((& $module {$script:service.Disposed -and $script:hostProcess.Disposed})) 'Service and process handles are released'
& $module {$script:deny=$true;$script:processLookups=0}
$failed=@(& $module {Stop-UnjunkAutodeskAccess -State @{Cancel=$false;Message=''}})
Assert ($failed.Count -eq 1 -and $failed[0].Status -eq 'Skipped' -and $failed[0].Reason -match 'administrator' -and (& $module {$script:processLookups}) -eq 0) 'Permission failures are reported without killing a still-running service'
$before=& $module {$script:stopCalls}
$cancel=@(& $module {Stop-UnjunkAutodeskAccess -State @{Cancel=$true;Message=''}})
Assert ($cancel.Count -eq 0 -and (& $module {$script:stopCalls}) -eq $before) 'Cancellation prevents service or process actions'
Write-Host "All $passed Autodesk service checks passed."
