#!/bin/sh
set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  echo "未找到 python3。请先安装 Python 3.11 或更新版本。"
  exit 1
fi

if [ ! -x ".venv/bin/python" ]; then
  echo "正在创建本地运行环境，请稍候..."
  python3 -m venv .venv
fi

if [ ! -f ".venv/.installed" ]; then
  echo "正在安装依赖，首次运行会稍慢..."
  .venv/bin/python -m pip install --disable-pip-version-check -e .
  echo ok > .venv/.installed
fi

.venv/bin/python scripts/app_preflight.py || code=$?
if [ "${code:-0}" = "2" ]; then
  exit 0
fi
if [ "${code:-0}" = "1" ]; then
  echo "启动检查失败。"
  exit 1
fi

echo "正在启动本地表格助手..."
.venv/bin/python -m streamlit run app.py --server.address 127.0.0.1 --server.port 8501 --server.headless false

