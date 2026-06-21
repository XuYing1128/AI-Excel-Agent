# AI-Excel-Agent

AI-Excel-Agent 是一个本地 AI 表格自动化工作台。你可以把自然语言需求交给 Codex，再由本项目生成、修改、清洗、分析和校验 Excel / CSV 文件。

## 能做什么

- 生成常见办公模板：个人预算、报价单、库存进销存、销售报表、电商分析、项目计划、课程/排班、考勤、财务测算、综合 Dashboard。
- 把 CSV / XLSX 数据清洗成规范表格，并生成 `clean_report`。
- 从销售数据生成带公式、图表和汇总页的 Excel 报表。
- 校验工作簿是否能打开、是否有错误值、是否缺少说明页、公式覆盖是否可疑。
- 优先把计算逻辑写成 Excel 公式，方便你后续在 Excel 里继续编辑。

## 不能保证什么

- 不能替代财务、税务、审计、法务或绩效审批的人工复核。
- 如果本机没有 Excel 或 LibreOffice，程序只能做静态校验，不能保证所有公式已经由真实 Excel 引擎重算。
- 自动分类和模板生成是 MVP 版本，复杂业务规则需要继续补充模板和校验器。

## 安装

建议使用 Python 3.10 或更高版本。

```powershell
cd D:\AI-Excel-Agent
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install -e .[dev]
```

如果只想装依赖，也可以运行：

```powershell
python -m pip install -r requirements.txt
```

## 基本命令

```powershell
python -m excel_agent.cli create --type budget --output outputs/demo_budget.xlsx
python -m excel_agent.cli create --type quotation --output outputs/quotation.xlsx
python -m excel_agent.cli analyze --input examples/input/sales.csv --output outputs/demo_sales_report.xlsx
python -m excel_agent.cli clean --input examples/input/messy_sales.csv --output outputs/cleaned.xlsx
python -m excel_agent.cli validate --input outputs/demo_sales_report.xlsx
```

也可以先让分类器判断需求：

```powershell
python -m excel_agent.cli classify "帮我做一个电商 SKU GMV 和退款分析表"
```

## 自然语言任务文件

可以把中文需求写入 `examples/tasks/*.txt`，再用 `run-task` 读取任务、判断类型、生成表格并自动校验：

```powershell
python -m excel_agent.cli run-task --task examples/tasks/task_sales.txt --output outputs/task_sales.xlsx
```

任务文件里如果写了本地输入路径，例如 `examples/input/sales.csv`，程序会先读取该文件；当前 MVP 对销售月报会调用分析器生成报表，其他类型会先校验输入可读并生成对应模板 demo。没有输入文件时会使用示例数据生成 demo。

## 如何让 Codex 使用这个项目

1. 在本目录工作。
2. 表格任务必须遵守 `AGENTS.md`，并使用 `skills/xlsx_pro/SKILL.md` 中的 `xlsx_pro` 技能。
3. 描述你想要的表格，例如“根据这个 CSV 做销售月报，加图表和异常检查”。
4. 要求输出文件放在 `outputs/`，并运行 `validate`。

## 添加新模板

1. 在 `templates/<新类型>/` 下写模板说明。
2. 在 `config/table_types.yaml` 注册类型、别名和关键词。
3. 在 `src/excel_agent/workbook_builder.py` 增加模板规格或专用构建函数。
4. 在 `tests/` 中增加一个生成测试。
5. 用 `python -m excel_agent.cli validate --input outputs/xxx.xlsx` 校验。

## 校验结果

校验命令会输出 JSON。`status` 为 `pass` 表示没有发现错误；`warn` 表示存在需要人工查看的风险；`fail` 表示文件无法打开或有明显错误。

```powershell
python -m excel_agent.cli validate --input outputs/demo_budget.xlsx --json outputs/demo_budget.validation.json
python skills/xlsx_pro/scripts/validate_workbook.py outputs/demo_budget.xlsx
```

报告字段含义：

- `summary`：文件大小、sheet 数量、可见 sheet 数、公式数量、是否存在说明页和数据页等概览。
- `errors`：必须修复的问题，例如文件打不开、没有可见 sheet、公式缓存存在错误值。
- `warnings`：建议人工检查的问题，例如缺少筛选、表头异常、格式风险、重复订单号、负库存。
- `suggestions`：对应错误或警告的修复建议。

如果 `status` 不是 `pass`，应先修复 `errors`，再处理 `warnings`，最后重新运行校验。

## 常见问题

**为什么打开后公式结果为空或没有刷新？**  
openpyxl 会写入公式，但不会计算公式。安装 Excel 或 LibreOffice 后可运行 `skills/xlsx_pro/scripts/recalc.py` 触发重算。

**为什么模板里有示例数据？**  
MVP 默认放少量示例数据，便于看清公式和图表。你可以直接替换输入区数据。

**能处理中文列名吗？**  
可以。清洗器会保留中文列名，并尝试识别日期、金额、数量等常见字段。
