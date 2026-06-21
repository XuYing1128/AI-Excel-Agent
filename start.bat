@echo off
setlocal
cd /d %~dp0

if not exist .venv\Scripts\python.exe (
  echo [ERROR] .venv was not found. Run install.bat first.
  pause
  exit /b 1
)

call .venv\Scripts\activate.bat
streamlit run app.py --server.address 127.0.0.1 --server.port 8501
pause
