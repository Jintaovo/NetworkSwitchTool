@echo off
rem ====================================================================
rem  NetworkSwitchTool launcher
rem  Starts the PowerShell GUI without an extra console window.
rem  The script itself asks for administrator rights (UAC) when needed.
rem ====================================================================
setlocal
set "SCRIPT=%~dp0NetworkSwitchTool.ps1"
if not exist "%SCRIPT%" (
  echo [ERROR] NetworkSwitchTool.ps1 not found next to this launcher.
  echo         Expected: "%SCRIPT%"
  pause
  exit /b 1
)
start "" powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%SCRIPT%"
exit /b 0
