# 安全与杀毒软件说明

AI-Excel-Agent 是本地 Python 项目，不包含打包的可执行文件、驱动、系统服务、持久化程序或隐藏下载器。

## 启动脚本

`start.bat` 只执行一条本地命令：

```bat
".venv\Scripts\python.exe" -m streamlit run app.py --server.address 127.0.0.1 --server.port 8501 --server.headless false
```

它不会：

- 调用 PowerShell；
- 隐藏窗口；
- 使用编码命令；
- 下载或执行远程脚本；
- 修改注册表或系统启动项；
- 监听公网地址。

浏览器由 Streamlit 自己以普通方式打开，服务仅监听 `127.0.0.1`。

## 安装脚本

`install.bat` 只负责：

1. 检查 Python；
2. 创建项目自己的 `.venv`；
3. 使用 pip 安装 `pyproject.toml` 中公开声明的依赖。

安装过程需要联网下载 Python 包，这是依赖安装，不是运行时下载器。安装完成后，`start.bat` 不会执行任何 pip 或下载命令。

## 自定义接口密钥

自定义接口配置保存在：

```text
data/private/api_settings.json
```

该目录已加入 `.gitignore`，不会提交到 Git。密钥不会写入任务报告或 Excel。

## 误报处理

如果安全软件仍然告警：

1. 先确认告警文件路径和项目 Git 提交；
2. 对比 GitHub 中对应脚本的明文内容；
3. 不要直接关闭杀毒软件；
4. 将样本提交给安全软件厂商进行误报复核。

项目不会要求用户添加整个目录到杀毒软件白名单。
