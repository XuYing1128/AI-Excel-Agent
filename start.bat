@echo off
setlocal
chcp 65001 >nul
cd /d %~dp0

if not exist .venv\Scripts\python.exe (
  echo [错误] 未找到本地运行环境，请先双击 install.bat 完成安装。
  pause
  exit /b 1
)

call .venv\Scripts\activate.bat
start "" powershell.exe -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 2; Start-Process 'http://127.0.0.1:8501'"
streamlit run app.py --server.address 127.0.0.1 --server.port 8501
pause
