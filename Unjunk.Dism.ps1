#requires -Version 5.1
# Visible console for native progress; the parent reads the durable result before continuing.
param(
    [Parameter(Mandatory)][ValidateSet('AnalyzeComponentStore','StartComponentCleanup','Help')][string]$Operation,
    [Parameter(Mandatory)][string]$OutputPath,
    [Parameter(Mandatory)][string]$ResultPath
)
$ErrorActionPreference='Stop'
$code=1
try {
    $Host.UI.RawUI.WindowTitle='Unjunk - DISM '+$Operation
    Write-Host ('Unjunk / DISM '+$Operation) -ForegroundColor Cyan
    Write-Host 'Live Windows servicing progress. The next cleanup stage starts when DISM finishes.'
    $arguments=if($Operation -eq 'Help'){@('/?')}else{@('/Online','/Cleanup-Image',('/'+$Operation),'/NoRestart','/English')}
    # Native stderr may produce PowerShell ErrorRecords; preserve it without masking the exit code.
    $ErrorActionPreference='Continue'
    & "$env:SystemRoot\System32\Dism.exe" @arguments 2>&1 | Tee-Object -FilePath $OutputPath | ForEach-Object { Write-Host $_ }
    $nativeCode=$LASTEXITCODE
    $ErrorActionPreference='Stop'
    if($null -eq $nativeCode){throw 'DISM did not return an exit code.'}
    $code=[int]$nativeCode
} catch {
    $text=$_.Exception.Message
    Write-Host $text -ForegroundColor Red
    Add-Content -LiteralPath $OutputPath -Value $text -ErrorAction SilentlyContinue
} finally {
    @{ExitCode=$code;CompletedAt=[DateTime]::UtcNow.ToString('o')} | ConvertTo-Json | Set-Content -LiteralPath $ResultPath -Encoding UTF8 -ErrorAction Stop
}
exit $code
