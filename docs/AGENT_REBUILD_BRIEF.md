# AI-Excel-Agent 智能体化改造交接文档（给 Codex）

> 这是一份**实现规格 + 工作指令**。目标读者是负责实现的编码代理（Codex）。
> 写作日期 2026-06-24。请在动手前先读：本文件 → `AGENTS.md` → `docs/STABILITY_REPORT_2026-06-23.md` → 下面"现状速读"列出的源码。
> **本项目与用户均为中文，所有用户可见文案、注释、提交信息用简体中文。**

---

## 0. 一句话目标

把现有的"分支判断 + 专用 builder"表格工具，升级为**一个多步智能体（agent）来调度现有能力**：网页收集并澄清需求 → 打包给 agent → agent 按 skill、调用本地工具（含沙箱 Python）、多模型分工协作 → 产出符合要求的复杂表格，并能记住用户偏好。

**定位：本地单机工具，朋友圈使用，各自在设置页填自己的 API key，无登录、无云端、无多租户、不开放公网端口。** 这是"路线 A"，不是 SaaS。

---

## 1. 硬约束（违反任何一条都算未完成）

1. **不推倒重来。** 现有 Excel 生成内核是最大资产，必须保留并被 agent 当工具调用：`template_filler.py`、`schedule_planner.py`、`domain_builders.py`、`rich_workbook_builder.py`、`custom_workbook_builder.py`、`validators.py`、`chart_spec.py`、`inline_table_parser.py`。
2. **不引入重型框架。** 不装 Hermes / LangGraph / TabClaw 等框架，**借鉴其思想、自己用 Python 实现** agent 主循环（预计 300–600 行）。理由：中文友好、不被外部协议牵着走、升级不冲突、可控可调试。
3. **不接 MCP（本期）。** 本地单进程直接调函数比走 MCP 更快更稳。把"对外暴露成 MCP server"留作后续可选层，预留接口但不实现。
4. **保留已验证的快路径，不要让 agent 拖慢它们。** 当前"模板+数据"任务 0.9 秒完成（`generation_service._generate_with_model` 里的 template-faithful fast-path）。这类确定场景**继续走快路径，不进 agent 循环**。agent 用于自由描述 / 未覆盖的长尾需求。
5. **测试常绿。** `".venv\Scripts\python.exe" -m pytest -q` 当前 57 项全过。每个阶段结束必须仍全过，并为新模块补测试。
6. **不破坏现有验收用例。** 见 §8，两个真实任务（排考、绩效薪酬）必须始终正确产出。
7. **沙箱 `run_python` 必须安全**（见 §5.3），并提供总开关（默认开，可在设置页关）。
8. **遵守 `AGENTS.md` 既有规则**（公式不硬编码、不覆盖用户原文件、输出绑定 task_id、高风险结果标注人工复核等）。
9. **增量提交**：每个阶段（§4 的一节）独立提交，提交信息说明做了什么 + 测试状态。

---

## 2. 现状速读（基于真实代码，不是旧快照）

技术栈：Python 3.11、Streamlit 单页、openpyxl/pandas/xlsxwriter、requests。OpenAI 兼容 HTTP 客户端。无数据库（用 JSON 文件）。

**入口与编排**
- `app.py` — Streamlit 单页。三页：制作表格 / 接口设置 / 最近文件。已有渐变 hero、步骤指示 `render_steps`、图表类型选择器、继续修改 tab。
- `src/excel_agent/services/generation_service.py` — `generate_from_task_spec()` 主入口。`_generate_with_model()` 含三级兜底（工具 agent → JSON 模式 → 本地）+ 两个 fast-path（内嵌表、模板+数据）；`_local_generate()` 按特征路由到各 builder；`_wrap_result()` 收尾。
- `src/excel_agent/services/llm_workbook_agent.py` — **当前的单次 blueprint 模型路径**（工具循环 + JSON 模式）。这是本次要**降级/替换**的部分：它让模型"一次性吐完整 blueprint"，对大请求会被推理模型截断。保留其确定性校验思想，但主 LLM 路径改为新的多步 agent。
- `src/excel_agent/services/api_task_planner.py` — `enhance_task_spec_draft()` 需求理解/澄清（无工具，单模型）。
- `src/excel_agent/services/custom_api_service.py` — `chat_completion()` / `chat_completion_with_tools()`，OpenAI 兼容客户端，`ApiCallResult`/`ToolChatResult`，`TRUNCATION_HINT`，`completion_endpoint()`，自带截断检测与有限重试。
- `src/excel_agent/services/subjective_review_service.py` — 审查（单模型，非阻塞）。
- `src/excel_agent/services/revision_service.py` — `build_revision_task_spec()` 继续修改。

**模型配置（要改造的痛点）**
- `src/excel_agent/api_settings.py` — `ApiSettings` 数据类，**只支持单个模型**，存 `data/private/api_settings.json`。字段：enabled/base_url/api_key/model/provider_name/timeout_seconds/use_for_intent/use_for_review/use_for_generation。
- 用户实测在用：base_url=`https://api.deepseek.com/v1`，model 在 `deepseek-v4-pro` / `deepseek-v4-flash` 间换过；都是推理模型，对大请求会截断。

**Excel 内核（agent 要调用的工具）**
- `template_filler.py` — `fill_template(template, data_paths, output, prompt, ...)`：克隆模板原文件→按表头定位→清空样例行→按列名填数据→自动派生计数表(姓名/计数/占比)→需要时调排考→裁掉多余空行→自动改标题里的考次号。**这是"模板忠实"的核心，已验证。**
- `schedule_planner.py` — `schedule_registrations()` 贪心排考；`parse_rooms/parse_capacity/parse_slots/parse_day_count` 从自然语言解析考场/容量/时段/天数。
- `domain_builders.py` — `build_performance_compensation_workbook()` + `can_build_performance_compensation()`（绩效薪酬专用，含 82 公式、排名、调薪、图表）。
- `rich_workbook_builder.py` — `build_rich_workbook(blueprint)` 把结构化 blueprint 写成工作簿；`normalize_blueprint` / `normalize_workbook_blueprint`（单表/多表）；9 种图表类型；`build_default_chart_spec()`；`_normalize_key_token` 列名归一（兼容 name/title 别名 + 记录键重映射）。
- `custom_workbook_builder.py` — `build_custom_workbook` / `build_dataset_workbook` / `build_inline_tables_workbook`。
- `chart_spec.py` — `CHART_TYPE_LABELS`（9 种中文标签）、`chart_type_from_text()`（大白话→图表类型）。
- `inline_table_parser.py` — `extract_inline_tables()` 从需求文字解析 tab/markdown/csv 表。
- `validators.py` — `validate_workbook()` / `inspect_workbook()` 确定性校验。
- 其它：`content_plan.py`、`task_spec.py`（`TaskSpec`/`TaskSpecDraft`）、`task_spec_builder.py`、`task_paths.py`（task_id、目录、staging）、`manifest.py`、`preview.py`、`io_utils.py`（`read_table` 支持 xls）、`intent_classifier.py`。

**skills**
- `skills/xlsx_pro/SKILL.md` + `references/*.md` + `scripts/*.py`（inspect/recalc/render_preview/validate）。

**测试**：`tests/` ~57 项；`pyproject.toml` 配置 `pythonpath=["src"]`、`testpaths=["tests"]`。

---

## 3. 目标架构（最终形态，本期落地前 6 块）

```
浏览器（Streamlit，沿用 app.py，逐步重构）
  │ 1) 描述需求 + 上传数据/模板 + 选模板模式
  ▼
需求澄清器（沿用 api_task_planner，改为可配置 planner 模型）
  │ 2) 产出结构化 TaskSpec + 澄清问题
  ▼
【新】Agent 编排器  services/agent/orchestrator.py
  │  多步循环：读上下文/偏好/skill → 选工具 → 执行 → 看结果 → 决定下一步 → finish
  │  确定场景命中快路径时直接调 generation_service，不空转
  ▼
【新】工具注册表  services/agent/tools/*.py
  ├ 复用：fill_template / schedule_exam / build_performance / build_rich_workbook /
  │        build_dataset / validate_workbook / inspect_workbook / read_table_summary
  └ 新增：run_python（沙箱）/ render_preview
  ▼
【新】模型网关 + 角色分工  model_registry.py（多 provider + 角色映射）
  ├ planner（理解/澄清）  builder（设计/选工具）  reviewer（审查）
  └ fast（分类/摘要/起名）  coder（写 run_python 代码，可选）
  ▼
【新】记忆层  memory_store.py（SQLite, data/private/memory.db）
  ├ 用户偏好（自动学，谨慎）  历史任务（可回放）  技能版本（可回滚）
  ▼
Excel 文件 + 结构化审查报告（沿用 outputs/tasks/<task_id>/ 结构）
```

> 注意：这与现状的差别是**多了 agent 编排层和模型角色层**；Excel 内核、任务目录、校验、manifest 全部沿用。

---

## 4. 要做的事（按此顺序，持续推进，每节一个提交）

### 阶段一：多模型 + 角色分工（先做，因为后面都依赖它）

**目标**：设置页能配置任意多个 OpenAI 兼容模型，并把每个"角色"指派给某个模型。

1. 新增 `src/excel_agent/model_registry.py`：
   - 新数据结构（存 `data/private/model_settings.json`）：
     ```json
     {
       "providers": [
         {"id":"deepseek-chat","name":"DeepSeek","base_url":"https://api.deepseek.com/v1",
          "api_key":"...","model":"deepseek-chat","timeout_seconds":120,"enabled":true},
         {"id":"doubao","name":"豆包",...}, {"id":"kimi",...}, {"id":"mimo",...},
         {"id":"glm",...}, {"id":"qwen",...}
       ],
       "roles": {
         "planner":"deepseek-chat", "builder":"deepseek-chat", "reviewer":"glm",
         "fast":"qwen", "coder":"deepseek-chat"
       },
       "agent_enabled": true,
       "run_python_enabled": true
     }
     ```
   - API：`load_model_settings()` / `save_model_settings()` / `get_provider(role) -> ProviderConfig | None` / `list_providers()` / `test_provider(cfg)`。
   - **向后兼容**：若只存在旧的 `data/private/api_settings.json`，首次加载时自动迁移成一个 provider，并把所有角色都指向它。旧文件保留不删。
   - 角色取不到时回退顺序：指定角色 → 任一 enabled provider → 报"未配置模型"。
2. 复用 `custom_api_service.chat_completion()`（它已是 OpenAI 兼容，6 家厂商都支持）。**不强制引入 LiteLLM**——除非后续要接非 OpenAI 格式或要成本统计，再考虑。把"按角色取 provider 然后调 chat_completion"封装成 `model_registry.chat(role, system, user, **kw)`。
3. 改造设置页（`app.py` 的 `render_settings_page`）：
   - 模型列表：增/删/改/测试连接（复用 `test_api_connection` 思路，但针对每个 provider）。
   - 角色指派：每个角色一个下拉框，选已配置的 provider。给出中文说明（planner=理解需求；builder=设计表格；reviewer=审查；fast=分类摘要起名；coder=写代码，可不配）。
   - 两个开关：`agent_enabled`（启用智能体编排）、`run_python_enabled`（允许沙箱执行模型写的 Python）。
   - 默认值要合理：未配置时 UI 明确提示"请先在此添加至少一个模型"。
4. 其它调用点（`api_task_planner`、`subjective_review_service`、新 agent）改为通过 `model_registry.chat(role=...)` 取模型，不再读单一 `ApiSettings`。保留 `ApiSettings` 类供迁移与兼容。
5. 测试：迁移逻辑、角色回退、save/load round-trip。

**验收**：能在设置页配 ≥2 个模型并分别指派角色；老配置自动迁移不报错；`pytest` 全过。

---

### 阶段二：通用 Agent 编排器（核心）

**目标**：自建多步 agent 主循环，替代 `llm_workbook_agent` 的"单次 blueprint"作为主 LLM 路径。

1. 新增 `src/excel_agent/services/agent/orchestrator.py`：
   - 入口：`run_agent(task_spec, task_paths, *, progress=None, max_steps=12) -> AgentResult`。
   - 循环（借鉴 Hermes/Claude Code，自己实现）：
     ```
     组装 system prompt（角色 = builder）：包含可用工具清单(JSON schema) + 命中的 skill 正文 + 用户偏好摘要 + 任务约束
     组装 user：original_request + confirmed TaskSpec + 已暂存文件清单(只给文件名/列名/行数摘要，不灌全量数据)
     loop step in 1..max_steps:
        model = model_registry.chat(role="builder", tools=tool_schemas, ...)
        if 模型调用工具: 执行工具 → 把结构化结果回灌 messages
        elif 模型调用 finish_task(output_path): 校验文件存在 → break
        else（纯文字）: 提示"请调用工具或 finish_task" → 累计 stalled，≥2 次跳出
     若未 finish 或步数耗尽: 回退到 generation_service 的确定性生成（绝不让用户空手）
     每一步写入 run_log.json（透明决策链：第几步、调了什么工具、参数摘要、结果摘要）
     ```
   - 工具调用走 OpenAI function-calling（`chat_completion_with_tools`）；模型不支持/截断时，退化为"JSON 指令"模式（让模型回 `{"tool":"...","args":{...}}`，自己解析）——复用现有截断检测。
   - **不让模型一次性吐全量数据**：数据通过工具按引用读取（`read_table_summary` 给列名+样例，真要全量让 `run_python` 直接读暂存文件）。这是避免推理模型截断的关键。
2. 接入 `generation_service.generate_from_task_spec`：
   - 顺序：① 命中快路径（模板+数据 / 内嵌表）→ 直接确定性生成（沿用，最快最稳）；② 否则若 `agent_enabled` 且配了 builder 模型 → `run_agent`；③ agent 失败 → 现有三级兜底。
   - 即"确定场景走专用路径，长尾走 agent，agent 兜底走本地"。
3. 把旧 `llm_workbook_agent` 标记为兼容保留（仍可作为 agent 的一个"build_workbook 工具"），但不再是主路径。
4. 测试：用 mock 模型（返回预设工具调用序列）驱动 orchestrator，断言它正确调用工具、正确 finish、超时正确兜底。

**验收**：mock 驱动下 agent 能多步调用工具并产出文件；真实模型下一个"自由描述"任务能跑通；失败必有兜底文件；`pytest` 全过。

---

### 阶段三：工具注册表 + 沙箱 run_python

1. 新增 `src/excel_agent/services/agent/tools/` 包。每个工具统一形态：
   ```python
   @dataclass
   class ToolResult: ok: bool; summary: str; data: dict; artifacts: list[str]
   # 每个工具: name, description(中文), json_schema(参数), handler(args, ctx)->ToolResult
   ```
   - 复用类工具（薄封装现有函数）：`read_table_summary`、`inspect_workbook`、`fill_template`、`schedule_exam`、`build_performance_compensation`、`build_rich_workbook`、`build_dataset`、`validate_workbook`、`render_preview`。
   - `read_table_summary(path)`：返回列名、行数、前 5 行样例、每列类型推断。**不返回全量数据。**
2. **`run_python`（杀手锏，谨慎实现）**：`src/excel_agent/services/agent/tools/run_python.py`
   - 作用：agent 现场写 pandas/openpyxl 代码处理我们没写过的需求（合并去重、按宿舍排座、跨表对账……），无需新增专用 builder。
   - 执行方式：**独立子进程**（`subprocess`，全新解释器，不继承父进程内存），不是父进程 `exec`。
   - 安全细则（全部必须）：
     - 超时（默认 60s，可配）；超时杀进程。
     - 工作目录限制：`cwd` 设为当前 task 的临时目录；代码只允许读 task 的 `input/` 暂存数据、写 task 的 `output/`。通过传入的路径变量约束，并在运行器里校验所有打开的路径都在白名单目录内。
     - import 白名单：`pandas, numpy, openpyxl, datetime, re, math, json, collections, itertools, statistics, decimal, csv`。运行器顶部安装一个 import 钩子，import 白名单外的模块直接抛错。
     - 禁网络：运行器顶部 monkeypatch `socket.socket` 抛错（本地工具，模型代码无需联网）。
     - 资源：可设最大内存（POSIX 用 `resource`；Windows 退化为仅超时，注释说明）。
     - 返回：stdout、stderr、新产生的文件列表、是否成功。失败把 traceback 摘要回灌给 agent 让它自己改。
   - **总开关**：`run_python_enabled=false` 时该工具不注册，agent 看不到它。设置页可关。
   - 文档：在工具描述和 README 写明"会在你本机执行模型生成的 Python 代码，已限时/限目录/限 import/禁网络；介意可在设置关闭"。
3. 测试：run_python 正常执行返回 stdout；越权 import 被拦；超时被杀；路径越界被拒；关闭开关后工具不出现。

**验收**：给 agent 一个"把两个名单合并、按姓名去重、统计每人报考科目数"的任务，它用 `run_python` 写代码解决，产出正确表格；安全测试全过。

---

### 阶段四：Skills 体系

1. 把"任务套路"沉淀为 skill（markdown playbook，agent 命中后注入 system prompt）：
   ```
   skills/xlsx_pro/SKILL.md        （已有，通用 Excel 操作）
   skills/template_fill/SKILL.md   （模板+数据 → fill_template 流程）
   skills/schedule/SKILL.md        （排考：解析约束→schedule_exam→填模板→校验）
   skills/perf_review/SKILL.md     （绩效薪酬：参数表+明细表+公式+排名+调薪）
   skills/data_clean/SKILL.md      （清洗：缺失/重复/异常→报告）
   skills/report_chart/SKILL.md    （报表+图表：选 chart_spec 类型）
   ```
2. 新增 `skills/registry.py`：`match_skills(task_spec) -> list[Skill]`（关键词匹配即可，无需向量）。每个 SKILL.md 带 frontmatter（name/triggers/tools），agent 启动时把命中的 1–2 个 skill 正文注入。
3. 每个 SKILL.md 写清：何时用、步骤、该调哪些工具、常见坑、验收要点。中文。
4. 测试：match_skills 对排考/绩效/清洗 prompt 命中正确 skill。

**验收**：agent 处理排考时确实注入 schedule skill 并按其流程调工具；`pytest` 全过。

---

### 阶段五：SQLite 记忆层（偏好 + 历史 + 技能版本）

1. 新增 `src/excel_agent/memory_store.py`，SQLite 在 `data/private/memory.db`（已在 .gitignore 的 data/private 下）。三张表：
   - `preferences(key, value, updated_at)` — 用户偏好。
   - `task_history(task_id, prompt, task_type, output_file, status, created_at)` — 可回放。
   - `skills(name, version, content, enabled, created_at)` — 技能版本化，可回滚。
2. **偏好学习（保守）**：任务完成后，用 fast 模型从"需求+最终结构"抽取**有限**偏好（如：sheet 名用中文 / 表头颜色 / 不要说明页 / 不要图表 / 列名别名映射 / 常做表型）。下次任务启动时把偏好摘要注入 agent。**只学这些低风险项**，不学业务逻辑。
3. **技能进化（更保守，本期只做存储+回滚，不做自动改生产）**：把表存好、UI 能查看/启用/禁用/回滚某版本即可。自动生成新技能留到以后，且必须"生成→跑测试样例→人工确认→启用"，绝不自动改线上逻辑。
4. UI：历史任务页接 task_history（可重开）；设置页加"我的偏好"查看/清空。
5. 测试：偏好读写、历史落库、技能版本回滚。

**验收**：连做两次"蓝色表头的考勤表"后，第三次新任务默认带蓝色表头偏好；历史任务可重开；`pytest` 全过。

---

### 阶段六：结构化多模型审查

1. 新增 `src/excel_agent/services/review/`，**固定 2 个审查器**（别做成模型互聊，费钱空转）：
   - `requirement_review`：是否完成用户要求（列、数据条数、计算、排序、模板一致性）。
   - `excel_usability_review`：公式/格式/冻结/筛选/图表/可导入性/有无多余 sheet。
   - 各用 `reviewer` 角色模型，输出**结构化 JSON**：`{"status":"pass|warn|fail","issues":[],"suggestions":[],"requires_user_confirmation":bool}`。**非阻塞**：审查失败不影响下载。
2. 把审查建议接到现有"继续修改"（`revision_service`）：一键把建议带入下一次修改（现状已有按钮，复用）。
3. 测试：mock reviewer 返回 JSON，断言解析与展示。

**验收**：生成后能看到两类结构化审查结果；建议可一键带入修改；`pytest` 全过。

---

### 阶段七：UI 整合（沿用 Streamlit，逐步收口）

1. 制作表格主流程串起来：上传 → 澄清问答 → 确认 TaskSpec/计划 → 生成（命中快路径秒出 / 否则 agent，过程显示决策步骤进度）→ 预览（含图表）→ 校验报告 → 审查 → 下载 → 继续修改。
2. agent 运行时把"第 N 步：调用了 XX 工具"实时显示（用现有 `st.status`/进度条），让用户看得到在做事（解决之前"点了没反应"的体验问题）。
3. 文案继续口语化、去术语（不出现 blueprint/agent/MCP 这些词，对用户说"正在分析/正在生成/正在检查"）。
4. 不拆前后端、不引入 Next.js（路线 A 不需要）。

**验收**：朋友能在不懂任何术语的情况下走完整流程；长任务有可见进度。

---

### 阶段八：打包分发（朋友双击即用）

1. `start.bat` 一键启动（已有，完善：自动建 venv/装依赖/起 Streamlit/开浏览器，幂等）。
2. 可选 PyInstaller 打 `.exe`（Windows）；Mac 给 `.command` 脚本。打包说明写进 README。
3. 首次启动引导：提示"请到接口设置添加你的模型 key"。
4. 验收：在一台干净的 Windows 机器上，按 README 双击即可启动并生成一张表。

---

## 5. 关键设计补充

### 5.1 模型分工建议（默认值，用户可在设置改）
- `fast`（便宜快）：分类、字段识别、起文件名、抽偏好 → Qwen / GLM-flash 之类。
- `planner`：理解需求、澄清提问、生成计划 → DeepSeek-chat / Kimi。
- `builder`：agent 主循环、选工具、设计结构 → 较强模型。
- `reviewer`：审查（换一个不同厂商增加视角多样性）→ GLM / 豆包。
- `coder`：写 run_python 代码 → 擅长代码的模型（可与 builder 同一个）。
- **不要每步都用最强模型**，否则朋友们的 token 账单很快上来。

### 5.2 避免推理模型截断（重要教训）
现状失败根因：让模型"一次性吐含全部数据的 blueprint"，推理模型把输出额度耗在思考上→截断。新 agent **必须**：数据按引用读取（工具返回摘要/让 run_python 直接读文件），模型只产出**结构与决策**，不复述数据。这样 token 占用小、不截断。

### 5.3 run_python 安全（再强调）
本地工具、用户自己的机器、模型写的代码（不是远程下发）。子进程 + 超时 + import 白名单 + 目录白名单 + 禁网络 + 总开关。Windows 无 `resource` 限内存时仅靠超时，注释写清。README 明确告知风险与关闭方式。

### 5.4 始终有兜底
任何 agent/模型失败，最终都要落到 `generation_service` 的确定性生成，**绝不让用户空手**，并用中文说明这次为何走了本地。

---

## 6. 目录与新增文件一览（建议）

```
src/excel_agent/
  model_registry.py                 # 阶段一
  memory_store.py                   # 阶段五
  services/agent/
    orchestrator.py                 # 阶段二
    tools/__init__.py               # 工具注册
    tools/excel_tools.py            # 封装现有 builder
    tools/run_python.py             # 沙箱（阶段三）
    tools/preview.py
  services/review/
    requirement_review.py           # 阶段六
    excel_usability_review.py
skills/
  registry.py                       # 阶段四
  template_fill/SKILL.md
  schedule/SKILL.md
  perf_review/SKILL.md
  data_clean/SKILL.md
  report_chart/SKILL.md
data/private/
  model_settings.json               # 新（多模型+角色）
  memory.db                         # 新（SQLite）
tests/
  test_model_registry.py
  test_agent_orchestrator.py
  test_run_python_sandbox.py
  test_skills_registry.py
  test_memory_store.py
  test_reviews.py
```

---

## 7. 不要做的事（明确排除）

- 不做登录/鉴权/多租户/计费/限流（本地工具，朋友圈自填 key）。
- 不上云、不开公网端口（只 127.0.0.1）。
- 不接 MCP（本期）、不装 Hermes/LangGraph/TabClaw 框架。
- 不做 Agent 自动改生产代码（技能进化本期只存储+回滚+人工确认）。
- 不引入 Next.js/React（继续 Streamlit）。
- 不删除/覆盖用户原始文件；输出仍进 `outputs/tasks/<task_id>/`。

---

## 8. 总验收（全部通过才算完成本期）

1. `".venv\Scripts\python.exe" -m pytest -q` 全过（含新增测试）。
2. **排考任务**（真实数据 `川外264次省考名单.xls` + 老师模板，模板模式=灵活）：仍 1 秒级产出 4 张表的模板忠实结果，易查分导入按考场/时段填满、单场≤65、同人不同场零冲突、标题考次正确。
3. **绩效薪酬任务**：仍正确产出 参数表+明细表+82 公式+图表。
4. **新长尾任务**（如"合并两个名单去重并统计每人科目数"，无专用 builder）：agent 用 run_python 解决并产出正确表格。
5. **多模型**：能配 ≥3 个 provider 并分别指派角色，各角色调用各自模型；老配置自动迁移。
6. 沙箱安全测试全过（越权 import/路径/超时被拦）。
7. 全程中文、长任务有可见进度、失败有中文兜底说明。

---

## 9. 给 Codex 的工作方式建议

- 先读本文件 + `AGENTS.md` + `docs/STABILITY_REPORT_2026-06-23.md` + §2 列的源码，确认现状再动手。
- 严格按 §4 阶段顺序，一阶段一提交，每次提交前跑 `pytest -q`。
- 复用优先：能薄封装现有函数就不要重写。
- 任何"让模型自由发挥改 Excel"的冲动都要压住——模型负责**选工具/填参/写受控代码/审查/解释**，确定性产出交给工具函数。
- 不确定的产品决策（如默认模型分工、是否默认开 run_python）按本文件默认值来，并在 README 写清如何改。
```
