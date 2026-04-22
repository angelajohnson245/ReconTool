@echo off
setlocal EnableExtensions
title Finance Recon Tool

REM Launcher directory (no hardcoded paths — follows this .cmd file)
set "SCRIPT_DIR=%~dp0"
set "RECON_APP="

if exist "%SCRIPT_DIR%recon_streamlit_app.py" (
  set "RECON_APP=%SCRIPT_DIR%recon_streamlit_app.py"
  goto :have_recon
)

for /d %%D in ("%SCRIPT_DIR%*") do (
  if exist "%%D\recon_streamlit_app.py" (
    set "RECON_APP=%%D\recon_streamlit_app.py"
    goto :have_recon
  )
)

echo Could not find recon_streamlit_app.py. Please keep all files in the same folder.
pause
exit /b 1

:have_recon
for %%I in ("%RECON_APP%") do set "RECON_DIR=%%~dpI"
for %%I in ("%RECON_APP%") do set "RECON_NAME=%%~nxI"

set "PY_CMD="
where py >nul 2>nul && set "PY_CMD=py"
if not defined PY_CMD where python >nul 2>nul && set "PY_CMD=python"
if not defined PY_CMD where python3 >nul 2>nul && set "PY_CMD=python3"

if not defined PY_CMD (
  echo Python was not found. Install Python from https://www.python.org/downloads/
  echo Be sure to check "Add python.exe to PATH" during setup, then double-click this launcher again.
  echo.
  pause
  exit /b 1
)

echo Checking dependencies...
for /f "delims=" %%M in ('%PY_CMD% -c "import importlib.util; m=('streamlit','pandas','openpyxl'); print(' '.join(x for x in m if importlib.util.find_spec(x) is None))"') do set "MISSING=%%M"

if "%MISSING%"=="" (
  echo Everything is already installed. Opening Finance Recon Tool...
) else (
  echo Installing missing packages...
  %PY_CMD% -m pip install --upgrade %MISSING%
  if errorlevel 1 (
    echo.
    echo Package install failed. Check the messages above.
    pause
    exit /b 1
  )
)

echo.
echo Launching Finance Recon Tool...
echo.

pushd "%RECON_DIR%"
%PY_CMD% -m streamlit run "%RECON_NAME%"
if errorlevel 1 (
  echo.
  echo The app stopped with an error. See messages above.
  popd
  pause
  exit /b 1
)
popd
