@echo off
setlocal
cd /d "%~dp0"

rem Keep this batch file ASCII-only. Some Windows cmd.exe code pages parse
rem UTF-8 Chinese text inside IF blocks as broken commands.
set "APP_PY=.venv\Scripts\python.exe"
set "SYS_PY="

if exist "%APP_PY%" goto :venv_ready

where python >nul 2>nul
if not errorlevel 1 (
  set "SYS_PY=python"
) else (
  where py >nul 2>nul
  if not errorlevel 1 set "SYS_PY=py -3"
)

if "%SYS_PY%"=="" (
  echo Python 3.11 or newer is required. Install Python and enable Add Python to PATH.
  pause
  exit /b 1
)

echo Creating local runtime. Please wait...
%SYS_PY% -m venv .venv
if errorlevel 1 goto :failed

:venv_ready
if not exist ".venv\.installed" (
  echo Installing dependencies. First run may take several minutes...
  "%APP_PY%" -m pip install --disable-pip-version-check -e .
  if errorlevel 1 goto :failed
  echo ok> ".venv\.installed"
)

"%APP_PY%" scripts\app_preflight.py
if errorlevel 2 exit /b 0
if errorlevel 1 goto :failed

echo Launching local Excel assistant...
"%APP_PY%" -m streamlit run app.py --server.address 127.0.0.1 --server.port 8501 --server.headless false
exit /b 0

:failed
echo.
echo Startup failed. Run install.bat once, then try start.bat again.
pause
exit /b 1
