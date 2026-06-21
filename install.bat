@echo off
setlocal
cd /d %~dp0

where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python was not found. Install Python 3.11 or newer first.
  pause
  exit /b 1
)

if not exist .venv\Scripts\python.exe (
  python -m venv .venv
  if errorlevel 1 goto :failed
)

call .venv\Scripts\activate.bat
python -m pip install -U pip
if errorlevel 1 goto :failed
python -m pip install -e .[dev]
if errorlevel 1 goto :failed

echo.
echo Installation completed. Double-click start.bat to open the local web tool.
pause
exit /b 0

:failed
echo.
echo [ERROR] Installation failed. Review the messages above.
pause
exit /b 1
