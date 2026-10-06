@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
if exist "%SCRIPT_DIR%.venv\Scripts\python.exe" (
  "%SCRIPT_DIR%.venv\Scripts\python.exe" "%SCRIPT_DIR%classify_bill_light.py" %*
  exit /b
)
where python >nul 2>&1
if not errorlevel 1 (
  python "%SCRIPT_DIR%classify_bill_light.py" %*
  exit /b
)
where py >nul 2>&1
if not errorlevel 1 (
  py -3 "%SCRIPT_DIR%classify_bill_light.py" %*
  exit /b
)
echo Python was not found. Install Python 3.10 or newer and follow README.md.
exit /b 1
