@echo off
if exist "%~dp0Unjunk.exe" (
    start "" "%~dp0Unjunk.exe"
) else (
    start "" powershell.exe -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0unjunk.ps1" -Gui
)