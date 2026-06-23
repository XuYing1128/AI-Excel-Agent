@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo 未找到 Python。请先安装 Python 3.11 或更新版本，并勾选 Add Python to PATH。
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo 正在创建本地运行环境，请稍候...
  python -m venv .venv
  if errorlevel 1 goto :failed
)

if not exist ".venv\.installed" (
  echo 正在安装依赖，首次运行会稍慢...
  ".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -e .
  if errorlevel 1 goto :failed
  echo ok> ".venv\.installed"
)

".venv\Scripts\python.exe" scripts\app_preflight.py
if errorlevel 2 exit /b 0
if errorlevel 1 goto :failed

echo 正在启动本地表格助手...
".venv\Scripts\python.exe" -m streamlit run app.py --server.address 127.0.0.1 --server.port 8501 --server.headless false
exit /b 0

:failed
echo.
echo 启动失败。请检查上面的提示，或先双击 install.bat 重新安装依赖。
pause
exit /b 1
