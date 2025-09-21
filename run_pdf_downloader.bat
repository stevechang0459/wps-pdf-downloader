:: SPDX-License-Identifier: MIT
:: Copyright (c) 2025 Steve Chang

REM run_pdf_downloader.bat

@echo off
REM Run a PowerShell script next to this BAT with one-time ExecutionPolicy Bypass.
REM Default target script name:
set "PS1=wps_pdf_downloader.ps1"

setlocal
pushd "%~dp0"

if not exist "%PS1%" (
  echo [ERROR] Cannot find "%PS1%" next to this BAT.
  echo        Edit the BAT and change the PS1 variable to your script name if needed.
  goto :end
)

REM Prefer PowerShell 7 (pwsh), fallback to Windows PowerShell
where pwsh >nul 2>nul
if %ERRORLEVEL%==0 (
  pwsh -NoProfile -Command "Unblock-File -LiteralPath '%~dp0%PS1%' ; exit 0" >nul 2>nul
  pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0%PS1%"
  set "RC=%ERRORLEVEL%"
  goto :report
)

where powershell >nul 2>nul
if %ERRORLEVEL%==0 (
  powershell -NoProfile -Command "Unblock-File -LiteralPath '%~dp0%PS1%' ; exit 0" >nul 2>nul
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0%PS1%"
  set "RC=%ERRORLEVEL%"
  goto :report
)

echo [ERROR] Neither PowerShell 7 (pwsh) nor Windows PowerShell is available.
set "RC=1"
goto :report

:report
echo.
echo Exit code: %RC%
echo.
:end
pause
popd
endlocal
exit /b %RC%
