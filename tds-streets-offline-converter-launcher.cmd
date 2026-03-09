@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%tds-streets-offline-converter.ps1"

if not exist "%PS_SCRIPT%" (
  echo.
  echo ERROR: Could not find "%PS_SCRIPT%"
  echo Make sure this launcher is in the same folder as tds-streets-offline-converter.ps1
  echo.
  pause
  exit /b 1
)

echo Launching TDS Streets Offline Converter...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%PS_SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
  echo Converter exited with error code %EXIT_CODE%.
) else (
  echo Converter finished.
)
echo.
pause
exit /b %EXIT_CODE%
