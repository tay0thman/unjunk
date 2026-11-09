#requires -Version 5.1
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition -eq 'Core') {
    & "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath
    if($LASTEXITCODE -ne 0) { throw 'Windows PowerShell build failed.' }
    return
}
$output=Join-Path $PSScriptRoot 'Unjunk.exe'
if(Test-Path -LiteralPath $output) { throw 'Unjunk.exe already exists. Close Unjunk, rename the old build, then build again.' }
$compiler=New-Object Microsoft.CSharp.CSharpCodeProvider
$parameters=New-Object CodeDom.Compiler.CompilerParameters
$parameters.GenerateExecutable=$true; $parameters.GenerateInMemory=$false; $parameters.OutputAssembly=$output
$parameters.CompilerOptions='/target:winexe /win32icon:"'+(Join-Path $PSScriptRoot 'unjunk.ico')+'"'
foreach($reference in @([Management.Automation.PowerShell].Assembly.Location,'System.Windows.Forms.dll','System.Core.dll','System.dll')) { $null=$parameters.ReferencedAssemblies.Add($reference) }
try {
    $result=$compiler.CompileAssemblyFromSource($parameters,(Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Unjunk.Launcher.cs') -Raw))
    if($result.Errors.HasErrors) { throw (($result.Errors | ForEach-Object {$_.ToString()}) -join [Environment]::NewLine) }
} finally { $compiler.Dispose() }
Write-Host "Built $output"
