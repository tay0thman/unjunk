#requires -Version 5.1
[CmdletBinding()]
param([switch]$History,[string]$ExpectedUserSid)
$ErrorActionPreference='Stop'
try {
    Import-Module (Join-Path $PSScriptRoot 'Unjunk.OneClick.psm1')
    $sid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    if($ExpectedUserSid -and $ExpectedUserSid -ne $sid) { throw 'Launch Unjunk with administrator rights for the same Windows account so the correct user caches and history are used.' }
    $principal=[Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA' -or (-not $History -and -not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
        $null=Start-UnjunkOneClick -History:$History
        return
    }
    $run=$null
    if(-not $History) {
        $exclusions=@(); $closeApps=$true; $settingsPath=Join-Path $env:LOCALAPPDATA 'Unjunk\settings.json'
        if(Test-Path -LiteralPath $settingsPath) {
            $settings=Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json
            if($null -ne $settings.PSObject.Properties['Exclusions']) { $exclusions=@($settings.Exclusions) }
            if($null -ne $settings.PSObject.Properties['CloseApps']) { $closeApps=[bool]$settings.CloseApps }
        }
        $run=Invoke-UnjunkOneClick -Exclusions $exclusions -CloseApps $closeApps
    }
    Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
    . (Join-Path $PSScriptRoot 'Unjunk.History.Gui.ps1')
    Show-UnjunkHistoryWindow -Runs @(Get-UnjunkHistory) -SelectedRunId $(if($run){$run.Id}else{''})
} catch {
    Add-Type -AssemblyName PresentationFramework
    $null=[Windows.MessageBox]::Show($_.Exception.Message,'Unjunk one-click cleanup')
    exit 1
}
