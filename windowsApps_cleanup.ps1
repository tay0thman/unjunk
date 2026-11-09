#requires -Version 5.1
# Legacy entry point now opens the protected GUI. It never uninstalls registered apps.
& (Join-Path $PSScriptRoot 'unjunk.ps1') -Gui
