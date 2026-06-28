# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包规格：AI-Excel-Agent Windows onedir。

打包：.venv\Scripts\python.exe -m PyInstaller AI-Excel-Agent.spec --noconfirm
产物：dist\AI-Excel-Agent\AI-Excel-Agent.exe（文件夹形式，双击即用）

设计要点：
- onedir（不是 onefile）：Streamlit 在 onefile 下兼容性差、启动慢，onedir 更稳。
- run_app.py 作主入口（用 streamlit.web.bootstrap 拉 app.py）。
- collect-all 收齐 streamlit/altair/pyarrow 等动态加载的隐藏依赖。
- 只打包运行时资源(app.py/skills/config/.streamlit/templates)，**绝不打包 data/private(outputs/examples（排除 API key 与使用记录）。
"""
from PyInstaller.utils.hooks import collect_all, collect_data_files

block_cipher = None

# ---- 收集动态依赖 ----
datas = []
binaries = []
hiddenimports = [
    "streamlit.web.bootstrap",
    "streamlit.runtime.scriptrunner",
    "pandas._libs",
    "pandas._libs.tslibs",
    "openpyxl",
    "xlsxwriter",
    "xlrd",
    "streamlit_antd_components",
    "skills.registry",  # skills 是顶层 namespace 包（无 __init__），需显式声明
]

for pkg in ["streamlit", "altair", "pyarrow", "streamlit_antd_components"]:
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h

# ---- 运行时资源（只读） ----
datas += [
    ("app.py", "."),
    ("run_app.py", "."),
    ("skills", "skills"),
    ("config", "config"),
    (".streamlit", ".streamlit"),
]
# templates 作为示例模板素材，可选随包
datas += [("templates", "templates")]
# skills 是 Python 包（含 import），确保 src 下的 excel_agent 也进来
datas += collect_data_files("excel_agent")

a = Analysis(
    ["run_app.py"],
    pathex=["src", "."],  # src 给 excel_agent，"." 给 skills 顶层包
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 排除测试/开发依赖，减小体积
        "pytest",
        "tests",
        "matplotlib.tests",
        "pandas.tests",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="AI-Excel-Agent",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,  # GUI 应用，不弹黑框
    disable_windowed_traceback=False,
    icon=None,  # 如有 icon.ico 可改为 ("icon.ico",)
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="AI-Excel-Agent",
)
