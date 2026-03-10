@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%tds-local-streets-auto-setup.ps1"

if not exist "%PS_SCRIPT%" (
  echo.
  echo ERROR: Could not find "%PS_SCRIPT%"
  echo.
  pause
  exit /b 1
)

echo Running one-click TDS streets setup...
echo This will auto-detect your Texas ZIP, convert, index, start backend,
echo install/update the Street Backend Manager app, and open TDS PAK.
echo.

powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%PS_SCRIPT%" -OpenBackendManagerAfterSetup -OpenTdsPakAfterSetup
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
  echo Setup exited with error code %EXIT_CODE%.
) else (
  echo Setup finished.
)
echo.
pause
exit /b %EXIT_CODE%
