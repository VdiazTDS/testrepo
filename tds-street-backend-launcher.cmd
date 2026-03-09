@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "BACKEND_SCRIPT=%SCRIPT_DIR%tds-street-backend.py"
set "DEFAULT_DB=%SCRIPT_DIR%tds-streets.sqlite"
set "DEFAULT_DATA_DIR=%USERPROFILE%\Documents\TDS-Pak-Street-Data"
set "DEFAULT_HOST=127.0.0.1"
set "DEFAULT_PORT=8787"

if exist "%DEFAULT_DATA_DIR%\tds-streets.sqlite" (
  set "DEFAULT_DB=%DEFAULT_DATA_DIR%\tds-streets.sqlite"
)

if not "%~1"=="" (
  set "DEFAULT_DB=%~1"
)

if not exist "%BACKEND_SCRIPT%" (
  echo.
  echo ERROR: Could not find "%BACKEND_SCRIPT%"
  echo.
  pause
  exit /b 1
)

set "PYTHON_CMD="
where py >nul 2>nul
if %ERRORLEVEL%==0 (
  set "PYTHON_CMD=py -3"
) else (
  where python >nul 2>nul
  if %ERRORLEVEL%==0 (
    set "PYTHON_CMD=python"
  )
)

if not defined PYTHON_CMD (
  echo.
  echo ERROR: Python was not found.
  echo Install Python from https://www.python.org/downloads/
  echo.
  pause
  exit /b 1
)

if not exist "%DEFAULT_DB%" (
  echo.
  echo WARNING: "%DEFAULT_DB%" was not found.
  echo Run tds-street-indexer-launcher.cmd first to build the index.
  echo.
)

echo.
echo Starting local streets backend...
echo URL: http://%DEFAULT_HOST%:%DEFAULT_PORT%
echo Press Ctrl+C in this window to stop the backend.
echo.
echo Note: If you are using VS Code Live Server, store the SQLite DB outside the served repo folder to avoid auto-reload on file watches.
echo.

%PYTHON_CMD% "%BACKEND_SCRIPT%" --db "%DEFAULT_DB%" --host %DEFAULT_HOST% --port %DEFAULT_PORT%
set "EXIT_CODE=%ERRORLEVEL%"
echo.
if not "%EXIT_CODE%"=="0" (
  echo Backend exited with error code %EXIT_CODE%.
) else (
  echo Backend stopped.
)
echo.
pause
exit /b %EXIT_CODE%
