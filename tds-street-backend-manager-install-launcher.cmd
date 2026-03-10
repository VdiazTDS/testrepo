@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%tds-street-backend-manager-install.ps1"

if not exist "%PS_SCRIPT%" (
  echo.
  echo ERROR: Could not find "%PS_SCRIPT%"
  echo.
  pause
  exit /b 1
)

echo Installing/Updating TDS PAK Street Backend Manager...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%PS_SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
  echo Installer exited with error code %EXIT_CODE%.
) else (
  echo Installer finished.
)
echo.
pause
exit /b %EXIT_CODE%
