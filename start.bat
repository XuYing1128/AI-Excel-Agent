@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Local environment not found. Run install.bat first.
  pause
  exit /b 1
)

".venv\Scripts\python.exe" -m streamlit run app.py --server.address 127.0.0.1 --server.port 8501 --server.headless false
