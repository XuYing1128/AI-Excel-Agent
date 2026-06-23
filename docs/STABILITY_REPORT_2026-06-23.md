# AI-Excel-Agent 稳定性改造报告 (2026-06-23)

> 面向 AI 编码代理 (Codex 等) 的快速交接文档。目标：让"本地接入大模型生成复杂表格"从**经常失败/输出不沾边**变为**稳定、高质量、模型主导**。所有改动在工作区，**未提交**。`pytest` 53 项全过。

## Codex 后续增强（同日）

- 新增多张内嵌表格解析：制表符、Markdown、CSV 风格数据可直接从需求文字进入 `TaskSpec`。
- 意图分类改为加权语义匹配；“绩效评估/薪酬调整”不再因出现“考勤”而误判为考勤模板。
- blueprint 支持 workbook-level `sheets`，允许参数表、明细表和跨表公式。
- 新增 `json-repair`，普通 JSON 和工具参数出现可修复语法问题时自动修复；截断响应仍重新请求。
- 模型不可用时先调用本地业务编译器；已实现绩效薪酬多表编译器，生成参数表、明细表、82 个公式、分组汇总和总计。
- 需求一致性校验不再因 `content_plan` 缺失而跳过；明确要求图表/公式但结果缺失时直接判定失败。
- Streamlit 生成与修改流程增加阶段进度条、运行中状态和持续计时。
- 回归测试更新为 `57 passed`。

## 本轮补充（模板忠实生成 + 考试编排 + 两个崩溃修复）

- **模板忠实生成**：新增 `template_filler.py`，当用户同时上传模板和数据并选 flexible/strict 时，**克隆模板原文件**（保留所有工作表/表头/样式/标题行），按列名匹配把数据写入；不再用通用生成器另起炉灶导致"模板对不上无法导入"。在 `_generate_with_model` 和 `_local_generate` 两处都加了 fast-path：识别到"模板+数据"即跳过慢模型，直接进入此路径。
- **考试编排**：新增 `schedule_planner.py`。从需求自动解析考场名（"二教E504/505/506"）、容量（"65个位置"）、时段（"八点到十点…"）、天数；贪心分配，保证同一学生不在同一时段两个考场。模板里出现 考点/考场地址/考试时间 列时自动启用（无需关键词）。
- **派生工作表**：`导出计数_姓名` 这类含 姓名/计数/占比 列的表会自动按姓名聚合并写入计数/占比。
- **崩溃修复**：
  - `_check_merged_filter_area` 在 `auto_filter.ref` 是单元格（如 "A1"）时抛 `TypeError: 'Cell' object is not iterable`；新增 `_cells_in_ref` 同时处理单 Cell 与 tuple-of-rows。
  - 模型多表 blueprint 某张表缺 `columns` 时整任务崩 `ValueError: 工作簿方案必须包含 columns`；`normalize_workbook_blueprint` 现在跳过坏 sheet。
- **实测**（用户真实数据 `川外264次省考名单.xls` + 老师模板）：**0.96s 完成**，模式 `template_faithful`；4 张工作表全保留，总信息 371 行齐全，易查分导入 371 行带考场/时间，导出计数 250 条；3 个考场 × 6 场次全用上，**0 个同人同场冲突**。

## 本轮补充（图表全类型 + UI 改版 + 复杂表验收）

- **图表引擎扩到 9 种类型**（`rich_workbook_builder.py`）：column/bar/line/area/pie/doughnut/radar/scatter/combo。新增 `_build_category_chart` / `_build_scatter_chart` / `_build_combo_chart` 分发；`ALLOWED_CHART_TYPES`、`ROUND_CHART_TYPES` 扩展；`_default_chart` → 公开 `build_default_chart_spec(columns, chart_type)`。
- **新增 `chart_spec.py`**：`CHART_TYPE_LABELS`（UI 友好标签）、`chart_type_from_text()`（大白话→类型，如 占比/构成→pie、趋势→line、对比→column、相关性→scatter、双轴→combo）、`wants_chart()`。
- **模型图表词汇**：`llm_workbook_agent` 工具 schema 的 chart `enum` 扩到 9 种；system prompt 增加“按用语选类型”指引；`_execute_build_tool` 在用户指定 `options.chart_types` 且模型未给图时，注入首选类型的默认图。
- **UI 改版（`app.py`）**：全新 CSS 设计系统（渐变 hero、步骤指示 `render_steps`、卡片/按钮/输入/Tab 美化）；确认页新增“图表样式”多选（用 `CHART_TYPE_LABELS`，存入 `spec.options.chart_types`）；文案更口语化。纯展示层，逻辑未动。
- **复杂表验收**：用真实 `deepseek-v4-pro` 跑绩效薪酬题目 → 模型失败(本次为公式含禁止字符)→自动兜底到 `domain_compiler:performance_compensation` → 生成「说明/参数表/明细表」、明细表 82 公式。**结构正确（非考勤模板）**。
- **已知权衡**：推理模型很慢，该题端到端 ~378s（先试模型再兜底）。强烈建议把接口模型换成 `deepseek-chat`（非推理，支持工具，秒级、JSON 更完整），可同时解决慢和截断两个问题。
- 之前 003333/011519/023142 三次失败是**旧版代码**所致；当前代码本地路径已能正确生成（`task_type=finance_model`、`inline_tables=5`、domain 编译器触发）。

---

## TL;DR（30 秒）

修了 4 类问题：

1. **生成会硬失败、连文件都不出**（用户反馈"现在做不了表格了"）→ 改为 **3 级自动兜底**：工具调用 → JSON 方案模式 → 本地确定性生成。始终产出可用表格，且兜底有明确提示（非静默）。
2. **代码过拟合 2~3 个演示样例**（销售表/天气表），导致换个需求"乱做" → 删掉上百行硬编码关键词判断，改为通用校验；并修复一个**列识别正则 bug**（把"季度合计用公式…"当成列名，反复打回正确方案）。
3. **图表错误**（柱状对比图只有 1 个系列）→ 改为多系列簇状竖向柱图。
4. **用户的模型"测试连接 OK 但生成总走本地兜底"** → 实测发现 `deepseek-v4-pro` 是**推理(thinking)模型**，慢(~100s)且思考吃光 token；而保存超时 45s、需求分析仅 1800 token → 超时/截断 → 退本地。**提高超时与 token 预算 + 截断检测 + 限制示例行数**后，用其真实模型 26s 端到端跑通（mode=`llm_tool_agent`，非兜底）。

外加：xls 模板无 LibreOffice 时不再报错（xlrd 兜底转换）、UI 措辞去术语化、文档与代码行为对齐。

---

## 架构与关键入口（先看这几处）

生成主链路：`app.py` → `services/generation_service.generate_from_task_spec()`

```
generate_from_task_spec
├─ settings.configured & use_for_generation:
│   └─ _generate_with_model()            # 新增：3 级兜底编排
│      1) llm_workbook_agent.generate_with_llm_agent()      # 工具调用循环
│      2) llm_workbook_agent.generate_blueprint_via_json()  # 新增：纯 JSON 方案（不依赖工具调用）
│      3) _local_generate()              # 兜底：从 content_plan / 上传数据确定性生成（带明确提示）
└─ 否则: _local_generate()
```

需求分析链路：`task_spec_builder.build_task_spec_draft()`（本地规则）+ `services/api_task_planner.enhance_task_spec_draft()`（模型理解，无工具）。
方案执行内核：`rich_workbook_builder.build_rich_workbook()`（模型 blueprint → openpyxl 实际写盘）。

---

## 逐文件改动

### 1. `src/excel_agent/services/generation_service.py`  (+241/-…，结构性重写)
- **删除**：模型失败即 `raise RuntimeError` 的硬失败逻辑。
- **新增** `_generate_with_model(task_spec, task_paths, settings, progress, notices)`：三级兜底编排。模型两条路径都失败时，调用 `_local_generate` 并向 `notices` 追加**明确**说明（含模型反馈原因），`result.mode` 标记为 `local_fallback:<inner_mode>`；写 `generation_model_fallback`(status=warning) 运行日志。
- **新增** `_local_generate(...)`：把原本内联的本地生成分支抽成函数；**放宽** explicit-structure 判定（不再要求 `policy=="custom_content"`，只要 `content_plan.explicit_structure && columns && 无上传文件` 就走 `build_custom_workbook`）。仅当确实无法产出文件时才抛错（被外层 `except` 捕获→失败结果）。
- import 增加 `generate_blueprint_via_json`。
- 删除一条术语化 notice。
- **行为变更**：模型失败不再让整个任务失败（除非本地也无法生成，如缺输入文件）。无测试断言旧硬失败行为，故安全。

### 2. `src/excel_agent/services/llm_workbook_agent.py`  (+341/-…，最大改动)
- **去过拟合**：`_blueprint_requirement_issues()` 从 ~140 行针对单一销售样例的硬编码（区域/省份/季度/绿点/⚠/固定列序…）缩为**通用检查**：空列、重复列、用户显式列序、聚合列(合计/小计/总计/平均)必须带 formula、内联数据行数匹配。新增常量 `_AGGREGATE_LABEL_HINTS`。
- **移除**：`_execute_build_tool()` 里"要求图表但 charts 为空就 built=False 打回"的硬卡点；改由 `build_rich_workbook(require_charts=True)` 自动补默认图。
- **修复正则 bug**：`_extract_requested_column_order()` 捕获组从 `[^\r\n。]+` 改为 `[^。；;\r\n]+`，并按 `_COLUMN_CLAUSE_STOP_WORDS`（公式/排序/生成/图表…）+ 长度截断过滤。此前会把"列为：A、B、C；C用公式…"里的后半句当成列名，导致正确 blueprint 被反复判错。
- **新增** `generate_blueprint_via_json(task_spec, task_paths, settings, *, progress=None) -> AgentGenerationResult`：不使用工具调用，要求模型只返回 JSON 方案，复用 `_execute_build_tool` 构建+校验。写 `llm_json_mode_started/completed/failed` 运行日志。
- **早退**：工具循环新增 `stalled` 计数，连续 2 轮"无工具调用且内容无法解析"则 break（让上层走 JSON/本地兜底，避免空烧 6 轮）。
- **超时**：`generate_with_llm_agent` 与 `generate_blueprint_via_json` 的 `agent_settings.timeout_seconds` 下限由 120→**240**。
- **system prompt**：新增"减少冗长推理 + 无上传文件时示例数据≤8行"（降低推理模型截断概率）。
- 成功 notice 去术语化。
- import 增加 `chat_completion`。

### 3. `src/excel_agent/services/custom_api_service.py`  (+29)
- 新增常量 `TRUNCATION_HINT`（"输出被截断…调大等待时间或换非推理模型 deepseek-chat"）。
- `ApiCallResult` 增加字段 `finish_reason: str | None = None`。
- `chat_completion()`：捕获 `finish_reason`；当 `content` 为空且 `finish_reason=="length"` 时返回 `success=False` + `TRUNCATION_HINT`（而非泛化的"无法识别"）。
- `chat_completion_with_tools()`：当无 tool_calls、content 为空且 `finish_reason=="length"` 时返回失败 + `TRUNCATION_HINT`。

### 4. `src/excel_agent/services/api_task_planner.py`  (+30)
- `enhance_task_spec_draft()`：为需求分析调用克隆 `planning_settings`，超时下限 **120s**，`max_tokens` 1800→**8000**（推理模型思考即超 1800→截断→静默退本地，是"分析不精细"的根因）。
- **放宽列采纳**：模型给的 `content_plan.columns` 现在在**未上传数据文件**时一律采纳（原来仅 generic_table 才采纳）。让确认页预览=最终产出，并把列结构带给生成步骤 → 缓解"需求与输出不沾边"。
- system prompt：鼓励给出合理默认列结构 + 澄清问题用于"请用户补充"。

### 5. `src/excel_agent/rich_workbook_builder.py`  (+44)
- `_add_chart()`：`bar`=横向、`column`=竖向(`type="col"`)；BarChart 设 `grouping="clustered"`、`overlap=-10`、`gapWidth=80`、`style=10`。
- `_default_chart()`：从"只取最后一个数值列"改为**取最多 4 个输入数值列**（跳过 formula/合计列）→ 真正的多系列对比。

### 6. `src/excel_agent/custom_workbook_builder.py`  (+38)
- `_add_custom_chart()`：同样改为多系列簇状竖向柱图（本地路径与模型路径图表质量对齐）。

### 7. `src/excel_agent/io_utils.py`  (+58)
- `convert_legacy_xls()`：拆出 `_convert_xls_with_soffice()`；soffice 不存在或失败时回退 `_convert_xls_basic()`（xlrd 读值 + openpyxl 写 xlsx，保留表名/表头/值/布局）。**xls 模板不再因缺 LibreOffice 硬失败**。注意 `_convert_xls_basic` 内含 `destination_dir.mkdir(...)`。
- 数据读取路径（`read_table` 的 `.xls` 分支）本就用 xlrd，正常（已用用户真实 .xls 验证 14列×371行）。

### 8. `tests/test_generation_fallback.py`  (新增)
4 个回归测试：模型失败→本地兜底出文件；JSON 模式构建；截断→`TRUNCATION_HINT`；默认图多系列。

### 9. `README.md` / `AGENTS.md`  (+4/+4)
把"模型失败必须硬报错"的旧表述改为"按三级兜底用已确认方案在本地生成可用表格 + 明确提示原因，不静默套无关模板"，使文档与新行为一致。

---

## 关键诊断证据（针对用户的 deepseek-v4-pro）

用真实接口 3 次实测：
- 小请求：工具调用 OK / JSON OK，~3s，`reasoning_tokens>0`（**确认是推理模型**）。
- 复杂表(20人考勤)：`finish_reason=tool_calls`、blueprint 合法可解析，但**耗时 103s、completion 11415 tokens(含 reasoning 8445)**。
- 结论：模型有效且会用工具；失败是**超时(saved 45s) + 分析步骤 1800 token 截断**所致，**与模型名无关**（先前"模型名无效"的判断是错的）。
- 修复后端到端实测：复杂销售表 **26s** 经模型完成（`mode=llm_tool_agent`），含两级表头/公式/总计/图表，列与需求一致。

---

## 验证方式

```bat
".venv\Scripts\python.exe" -m pytest -q          # 53 passed
".venv\Scripts\python.exe" scripts\v1_smoke_test.py
".venv\Scripts\python.exe" -m streamlit run app.py --server.address 127.0.0.1 --server.port 8501
```
- 兜底/JSON/图表均有 mock 单测覆盖（`tests/test_generation_fallback.py`）。
- 端到端走真实模型需在「接口设置」配置可用接口。

---

## 待办 / 给用户的建议（非代码阻塞项）

1. **模型选择**：`deepseek-v4-pro`（推理）可用但慢(20~100s)；日常建议在「接口设置」换 **`deepseek-chat`**（非推理，支持工具，快得多）。两者现在都稳定。
2. 若坚持用推理模型，建议把「接口设置」的"等待时间"从 45 调到 ≥120（代码已有内部下限兜底，但保存值仍影响主观审查等调用）。
3. `data/private/api_settings.json` 是用户私密配置（含 key），**未改动**；不要在未询问下修改。
4. 主观审查调用(`subjective_review_service`)超时仍用原始 saved 值——非阻塞（失败=不出审查建议），未改；如需可同样加超时下限。
5. 改动未提交；建议 review 后 `git commit`（可参考本报告作为 commit 说明）。
6. 行尾：仓库文件被改后 Git 提示 LF→CRLF，属正常（Windows），如有 `.gitattributes` 规范可统一。
