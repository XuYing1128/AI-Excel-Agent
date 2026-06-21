# AI-Excel-Agent V1

AI-Excel-Agent 是一个面向 Windows 的本地 Excel / CSV 自动化工具。网页支持输入自然语言需求、上传本地数据、确认生成内容、在线预览、自动检查、继续修改和下载结果。

所有文件默认在本机处理。页面已隐藏 Streamlit 的部署按钮和开发者菜单。

## 核心原则

- **大模型不碰具体单元格。** 模型只允许做意图理解、一次性澄清、结果解释和建议性审查。
- **客观正确性不花 token。** 完整性、重复、公式错误、时间冲突、字段异常和数据缺失由确定性 Python 校验器检查。
- **先确认生成内容，再生成。** 页面会显示标题、文件名、数据来源、要求列和自动计算项。
- **明确需求不再硬套模板。** 用户给出列、数据或计算规则时，会按内容方案生成并做一致性检查。
- **主观审查不阻塞下载。** 当前默认关闭；即使失败，也不影响下载 Excel 和校验报告。
- **自定义接口完全可选。** 不配置接口也能直接使用本地规则和表格内核。
- **生成后可以继续修改。** 审查建议可直接带入下一次修改，每次修改都保留上一版本。

## V1 能做什么

现有内核支持 13 类表格：

1. `personal_budget`
2. `family_budget`
3. `quotation`
4. `invoice_draft`
5. `inventory`
6. `sales_report`
7. `ecommerce_analysis`
8. `project_plan`
9. `schedule`
10. `attendance`
11. `finance_model`
12. `dashboard`
13. `generic_table`

其中：

- `sales_report` 在提供可识别的 CSV/XLSX/XLSM 时，会调用现有销售分析器生成真实数据报表。
- 上传其他类型的数据时，会按文件的实际字段生成原始数据、清洗数据和报告，不再套用无关示例模板。
- 需求文字中明确提供列和数据时，会生成自定义内容表格。
- 生成、公式、样式、图表和校验继续复用现有 `workbook_builder.py`、公式库、样式库和校验器。

## 安装

先安装 [Python 3.11 或更高版本](https://www.python.org/downloads/windows/)，安装时勾选 **Add Python to PATH**。

### 双击安装

1. 双击 `install.bat`。
2. 脚本会创建 `.venv` 并安装项目及测试依赖。

### 命令行手动安装

```bat
cd D:\AI-Excel-Agent
python -m venv .venv
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -e ".[dev]"
```

## 启动网页

双击：

```text
start.bat
```

或手动运行：

```bat
".venv\Scripts\python.exe" -m streamlit run app.py --server.address 127.0.0.1 --server.port 8501 --server.headless false
```

浏览器打开：

```text
http://127.0.0.1:8501
```

页面为全中文单页，不需要登录。工具栏、部署入口、开发菜单和统计上报已关闭。

## 自定义模型接口

点击页面顶部的“接口设置”进入独立设置页。设置只有点击“保存设置”后才会生效，需要填写：

- 接口名称
- 接口地址，例如 `https://你的接口地址/v1`
- 接口密钥
- 模型名称
- 超时时间

可分别控制：

- 是否用于辅助理解用户需求
- 是否用于生成后的建议审查

接口配置保存在本机：

```text
data/private/api_settings.json
```

该目录已加入 `.gitignore`，不会提交到 Git。模型只接收用户文字、输入文件名、生成方案和结构化摘要，不接收完整工作簿行数据，也不能修改具体单元格。接口连接失败时自动回退到本地规则，不影响表格生成和下载。

## 网页使用流程

1. 输入完整需求，或上传 `.csv`、`.xlsx`、`.xlsm`。
2. 点击“分析需求”。
3. 系统调用本地规则分类器识别表格类型。
4. 必要时只追问一轮信息。
5. 查看并确认标题、文件名、数据来源、列、公式和汇总内容。
6. 点击“确认并生成”。
7. 本地确定性内核生成 Excel。
8. 自动运行确定性校验器。
9. 在页面中直接预览表格、生成方案、检查报告和审查建议。
10. 将审查建议带入“继续修改”，生成保留旧版的新版本。
11. 通过“最近文件”重新打开历史任务并继续修改。
12. 需要时下载 Excel 和相关报告。

## 输出目录

每个网页任务有独立 `task_id`：

```text
outputs/
  manifest.json
  tasks/
    20260621_153012_sales_report/
      input/
      output/
        根据内容生成或用户指定的文件名.xlsx
      reports/
        validation.json
        subjective_review.json
      task_spec.json
      run_log.json
  legacy/
```

`outputs/manifest.json` 保存任务历史。旧版平铺输出可运行以下脚本安全归档，脚本只移动、不删除：

```bat
".venv\Scripts\python.exe" scripts\archive_old_outputs.py
```

## 测试与验收

```bat
".venv\Scripts\python.exe" -m pytest -q
".venv\Scripts\python.exe" scripts\v1_smoke_test.py
".venv\Scripts\python.exe" -m streamlit run app.py --server.address 127.0.0.1 --server.port 8501 --server.headless false
```

Smoke Test 会：

- 构造销售月报需求；
- 生成结构化任务方案；
- 创建独立任务目录；
- 调用现有销售分析内核；
- 运行确定性校验；
- 写入 manifest；
- 输出最终文件路径或明确失败原因。

## CLI

旧 CLI 仍然保留，用于开发和兼容：

```powershell
python -m excel_agent.cli create --type personal_budget --output temp\budget.xlsx
python -m excel_agent.cli analyze --input examples\input\sales.csv --output temp\sales_report.xlsx
python -m excel_agent.cli clean --input examples\input\messy_sales.csv --output temp\cleaned.xlsx
python -m excel_agent.cli validate --input temp\sales_report.xlsx
```

日常使用建议通过网页，以确保生成方案、任务目录、检查报告和历史清单完整生成。

## V1 能力边界

- 不做真实 Excel 引擎重算，不安装或依赖 LibreOffice、Excel COM、OfficeCLI。
- openpyxl 会写入公式，但不会计算公式缓存；打开 Excel 后由 Excel 自行重算。
- 不做账号、鉴权、多用户、云端、队列或 EXE 打包。
- 不接入 `config/*.yaml` 或 `templates/*/template.yaml` 的配置化重构。
- 不做自进化，不允许 Agent 自动修改项目代码。
- 自定义接口为可选功能；建议审查仅是建议，不是客观校验证明。
- 财务、税务、工资、审计、绩效、合同报价等高风险结果必须人工复核。

## 安全说明

- `start.bat` 不调用 PowerShell、不隐藏窗口、不下载远程脚本，只以前台方式启动本地 Python 和 Streamlit。
- `install.bat` 只创建 `.venv` 并通过 pip 安装 `pyproject.toml` 中明文声明的依赖。
- `.env` 已加入 `.gitignore`，不得提交。
- `data/private/` 已加入 `.gitignore`，自定义接口密钥只保存在本机。
- 上传文件复制到当前任务的 `input/`，原文件不会被覆盖。
- 任务结果保存在本地 `outputs/tasks/`。
- 默认不向外部模型发送完整工作簿、行数据或单元格内容。
- 即使两个主观模型都认为通过，也只能表述为“未发现明显主观问题”，不能替代确定性校验和用户确认。
- 详细说明见 `SECURITY.md`。如果安全软件仍然告警，请先核对告警路径并提交厂商误报复核，不建议关闭杀毒软件或粗暴添加整个目录白名单。

## 项目规则

- 项目级规则：`AGENTS.md`
- 表格技能规则：`skills/xlsx_pro/SKILL.md`
