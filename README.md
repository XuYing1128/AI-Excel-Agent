# AI-Excel-Agent V1

AI-Excel-Agent 是一个面向 Windows 的本地 Excel / CSV 自动化工具。V1 提供运行在 `127.0.0.1` 的 Streamlit 网页，用于输入自然语言需求、上传本地数据、确认结构化任务计划、生成 Excel、查看确定性校验报告并下载结果。

所有文件默认在本机处理。它不是托管服务，不包含账号、鉴权、多租户、云端或任务队列。

## 核心原则

- **大模型不碰具体单元格。** 模型只允许做意图理解、一次性澄清、结果解释和建议性审查。
- **客观正确性不花 token。** 完整性、重复、公式错误、时间冲突、字段异常和数据缺失由确定性 Python 校验器检查。
- **先确认 TaskSpec，再生成。** 用户确认结构化执行计划之前，网页不会调用生成内核。
- **主观审查不阻塞下载。** 当前默认关闭；即使失败，也不影响下载 Excel 和校验报告。

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
- 其他类型在 V1 中主要生成标准模板/demo，不保证自动完成复杂真实数据分析。
- 生成、公式、样式、图表和校验继续复用现有 `workbook_builder.py`、公式库、样式库和校验器。

## 安装

先安装 [Python 3.11 或更高版本](https://www.python.org/downloads/windows/)，安装时勾选 **Add Python to PATH**。

### 双击安装

1. 双击 `install.bat`。
2. 脚本会创建 `.venv` 并安装项目及测试依赖。

### PowerShell 手动安装

```powershell
cd D:\AI-Excel-Agent
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install -e .[dev]
```

## 启动网页

双击：

```text
start.bat
```

或手动运行：

```powershell
.\.venv\Scripts\Activate.ps1
streamlit run app.py --server.address 127.0.0.1 --server.port 8501
```

浏览器打开：

```text
http://127.0.0.1:8501
```

## 网页使用流程

1. 输入一句话需求，或上传 `.csv`、`.xlsx`、`.xlsm`。
2. 点击“分析需求”。
3. 系统调用本地规则分类器识别表格类型。
4. 必要时只追问一轮信息。
5. 查看并确认 `TaskSpec`：类型、目标、输入、图表、汇总、说明页和系统假设。
6. 点击“确认并生成”。
7. 本地确定性内核生成 Excel。
8. 自动运行确定性校验器。
9. 查看 `pass` / `warn` / `fail`、问题和建议。
10. 下载 Excel、`validation.json`、`task_spec.json` 和可选主观审查报告。

## 输出目录

每个网页任务有独立 `task_id`：

```text
outputs/
  manifest.json
  tasks/
    20260621_153012_sales_report/
      input/
      output/
        result.xlsx
      reports/
        validation.json
        subjective_review.json
      task_spec.json
      run_log.json
  legacy/
```

`outputs/manifest.json` 保存任务历史。旧版平铺输出可运行以下脚本安全归档，脚本只移动、不删除：

```powershell
.\.venv\Scripts\python.exe scripts\archive_old_outputs.py
```

## 测试与验收

```powershell
.\.venv\Scripts\Activate.ps1
pytest -q
python scripts\v1_smoke_test.py
streamlit run app.py --server.address 127.0.0.1 --server.port 8501
```

Smoke Test 会：

- 构造销售月报需求；
- 生成 TaskSpec；
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

日常使用建议通过网页，以确保 TaskSpec、任务目录、校验报告和 manifest 完整生成。

## V1 能力边界

- 不做真实 Excel 引擎重算，不安装或依赖 LibreOffice、Excel COM、OfficeCLI。
- openpyxl 会写入公式，但不会计算公式缓存；打开 Excel 后由 Excel 自行重算。
- 不做账号、鉴权、多用户、云端、队列或 EXE 打包。
- 不接入 `config/*.yaml` 或 `templates/*/template.yaml` 的配置化重构。
- 不做自进化，不允许 Agent 自动修改项目代码。
- 主观模型审查默认关闭，仅是建议，不是客观校验证明。
- 财务、税务、工资、审计、绩效、合同报价等高风险结果必须人工复核。

## 安全说明

- `.env` 已加入 `.gitignore`，不得提交。
- 上传文件复制到当前任务的 `input/`，原文件不会被覆盖。
- 任务结果保存在本地 `outputs/tasks/`。
- 默认不向外部模型发送完整工作簿、行数据或单元格内容。
- 即使两个主观模型都认为通过，也只能表述为“未发现明显主观问题”，不能替代确定性校验和用户确认。

## 项目规则

- 项目级规则：`AGENTS.md`
- 表格技能规则：`skills/xlsx_pro/SKILL.md`
