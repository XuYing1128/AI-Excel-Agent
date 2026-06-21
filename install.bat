@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python 3.11 or newer is required.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  python -m venv .venv
  if errorlevel 1 goto :failed
)

".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -e ".[dev]"
if errorlevel 1 goto :failed

echo.
echo Installation completed. Double-click start.bat to launch the local app.
pause
exit /b 0

:failed
echo.
echo Installation failed. Review the messages above.
pause
exit /b 1
