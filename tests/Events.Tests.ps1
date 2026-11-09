#requires -Version 5.1
# All event providers are mocked. No real Event Viewer channel is read or cleared.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Events.psm1') -Force
$module=Get-Module Unjunk.Events;$passed=0
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
& $module {
    $script:names=@('Application','Security','Microsoft-Windows-Windows Defender/Operational','CrowdStrike/Operational','Denied/Analytic','Empty','Unreadable')
    $script:cleared=New-Object 'System.Collections.Generic.List[string]'
    function script:Get-UnjunkEventLogNames {$script:names}
    function script:Get-UnjunkEventLogRecordCount([string]$Name){if($Name -eq 'Unreadable'){throw 'Read denied'};if($Name -eq 'Empty'){return 0L};return 15L}
    function script:Clear-UnjunkEventLog([string]$Name){if($Name -eq 'Denied/Analytic'){throw 'Windows refused this channel'};$script:cleared.Add($Name)}
}
$logs=@(Get-UnjunkEventLogs)
Assert ($logs.Count -eq 7 -and @($logs | Where-Object Name -eq 'Security').Count -eq 1 -and @($logs | Where-Object Name -eq 'CrowdStrike/Operational').Count -eq 1) 'Inventory includes security, audit and security-product channels'
Assert (@($logs | Where-Object {$_.Name -eq 'Unreadable' -and $_.ReadError}).Count -eq 1 -and (& $module {$script:cleared.Count}) -eq 0) 'Preview retains unreadable channel identities without clearing anything'
$preview=Invoke-UnjunkEventLogCleanup -Logs $logs -WhatIf
Assert ($preview.ClearedLogs -eq 0 -and (& $module {$script:cleared.Count}) -eq 0) 'WhatIf never clears event logs'
& $module {$script:names+=@('CreatedAfterScan');$script:names=@($script:names | Where-Object {$_ -ne 'Application'})}
$result=Invoke-UnjunkEventLogCleanup -Logs ($logs+$logs[1]) -Confirm:$false
$calls=@(& $module {$script:cleared})
Assert ('Security' -in $calls -and 'Microsoft-Windows-Windows Defender/Operational' -in $calls -and 'CrowdStrike/Operational' -in $calls) 'Requested security and security-product channels use the same clearing API'
Assert ('Application' -notin $calls -and 'CreatedAfterScan' -notin $calls -and 'Empty' -notin $calls) 'Missing, newly created and empty channels are preserved'
Assert (@($calls | Where-Object {$_ -eq 'Security'}).Count -eq 1) 'Duplicate channel selections are cleared once'
Assert ($result.ClearedLogs -eq 4 -and $result.ItemsCleaned -eq 4 -and $result.SkippedLogs -eq 2) 'Results count confirmed cleared channels and report denied channels separately'
Assert ($result.Output -match 'Windows refused' -and $result.Output -match 'Logging stays enabled') 'Access failures remain visible without disabling logging or bypassing protection'
$before=$calls.Count
$result=Invoke-UnjunkEventLogCleanup -Logs $logs -State @{Cancel=$true;Message=''} -Confirm:$false
Assert ($result.ItemsCleaned -eq 0 -and (& $module {$script:cleared.Count}) -eq $before) 'Cancellation prevents later channel clears'
Write-Host "All $passed Event Viewer checks passed."
