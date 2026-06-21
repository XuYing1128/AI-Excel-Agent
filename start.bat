@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Local environment not found. Run install.bat first.
  pause
  exit /b 1
)

".venv\Scripts\python.exe" scripts\app_preflight.py
if errorlevel 2 exit /b 0
if errorlevel 1 (
  echo Startup check failed. Run install.bat, then try again.
  pause
  exit /b 1
)

".venv\Scripts\python.exe" -m streamlit run app.py --server.address 127.0.0.1 --server.port 8501 --server.headless false
