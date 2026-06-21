@echo off
setlocal
chcp 65001 >nul
cd /d %~dp0

where python >nul 2>nul
if errorlevel 1 (
  echo [错误] 未找到 Python，请先安装 Python 3.11 或更高版本。
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
echo 安装完成。以后双击 start.bat 即可打开本地表格助手。
pause
exit /b 0

:failed
echo.
echo [错误] 安装失败，请查看上方提示。
pause
exit /b 1
