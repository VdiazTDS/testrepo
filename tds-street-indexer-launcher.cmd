@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "INDEXER_SCRIPT=%SCRIPT_DIR%tds-street-indexer.py"
set "DEFAULT_DATA_DIR=%USERPROFILE%\Documents\TDS-Pak-Street-Data"
if not exist "%DEFAULT_DATA_DIR%" (
  mkdir "%DEFAULT_DATA_DIR%" >nul 2>nul
)
set "DEFAULT_DB=%DEFAULT_DATA_DIR%\tds-streets.sqlite"

if not exist "%INDEXER_SCRIPT%" (
  echo.
  echo ERROR: Could not find "%INDEXER_SCRIPT%"
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

echo.
echo Indexing streets into SQLite...
echo Output DB: "%DEFAULT_DB%"
echo.

if "%~1"=="" (
  %PYTHON_CMD% "%INDEXER_SCRIPT%" --db "%DEFAULT_DB%"
) else (
  %PYTHON_CMD% "%INDEXER_SCRIPT%" "%~1" --db "%DEFAULT_DB%"
)

set "EXIT_CODE=%ERRORLEVEL%"
echo.
if not "%EXIT_CODE%"=="0" (
  echo Indexer exited with error code %EXIT_CODE%.
) else (
  echo Indexing complete.
)
echo.
pause
exit /b %EXIT_CODE%
