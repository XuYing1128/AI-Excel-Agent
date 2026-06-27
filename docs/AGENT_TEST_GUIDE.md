# AI-Excel-Agent 自动化压测指南（给测试智能体读）

你是本项目的**自动化压测智能体**。你的任务：长时间循环、自行出题，模拟真实业务场景，把这套
“自然语言 → 复杂 Excel 表格”的工具压到出问题，并把每个问题**归因到 模型/工具/代码/网络 哪一层**，
便于维护改进。**目标是暴露问题，不是凑通过率**——专挑容易翻车的复杂场景。

---

## 1. 这个项目是什么

- 一个**本地 AI Excel 智能体**：用户用中文描述需求，它调大模型（builder=豆包写 `run_python` 代码、
  planner=豆包理解需求、reviewer=deepseek 审查/诊断）生成复杂 `.xlsx`。
- **内核是“AI 当主角”**：豆包用 Python（openpyxl/pandas）直接写代码生成/就地编辑工作簿；本地规则只当后备。
- **真算闭环保证公式正确**：交卷前用本机 LibreOffice 把公式真算一遍，发现 `#VALUE!`/循环引用就退回让模型修。
- 详细背景可读 `README.md`、`AGENTS.md`。

## 2. 怎么跑一道题（用统一 harness，别自己重写生成/校验）

入口：`scripts/agent_test_harness.py`。它封装了「生成 → LibreOffice 真算 → 校验 → 结构提取 →
按期望判定 →（失败时）根因诊断」，并把每条结果追加到 `outputs/_agent_test/results.jsonl`。

命令行：
```bash
python scripts/agent_test_harness.py --prompt "做一张……" --expect '{"min_sheets":2,"min_charts":1}' --slug case_001
# 带上传文件（就地编辑/模板填充类）：
python scripts/agent_test_harness.py --prompt "把表A工资按姓名填到表B花名册的月工资列，其它别动" \
  --input demo_fill_table/表A_工资数据.xlsx demo_fill_table/表B_花名册.xlsx \
  --expect '{"keep_sheets_from_input":true,"must_contain":["月工资"]}' --slug case_002
```
或在 Python 里 `import`：
```python
import sys; sys.path.insert(0, "scripts")
from agent_test_harness import run_one_case
r = run_one_case(prompt, input_files=[...], expect={...}, slug="case_003")
```

`--expect` 支持的硬指标（都可选，按题目挑着给）：
- `min_sheets`：至少几个工作表
- `sheet_names_include`：必须含哪些工作表名（子串匹配）
- `min_charts` / `chart_types`：至少几个图表 / 含哪些类型（`bar`/`line`/`pie`/`area`/`radar`/`scatter`/`doughnut`）
- `min_formula_count`：至少几个公式（验证“用公式而非写死结果”）
- `must_contain`：产物里必须出现的文本（如 `合计`、`及格率`、`环比`）
- `keep_sheets_from_input`：就地编辑必须保留输入文件的**全部**工作表（防止 AI 重造丢表）
- `max_time_s`：生成耗时上限

## 3. 判定标准（harness 返回 `passed`）

`passed = 生成成功 且 真算无错(LibreOffice 可用时) 且 所有 expect 硬指标通过`。

- **真算金标准最重要**：`recalc.ok=false` 表示产物里有公式算出 `#VALUE!`/`#DIV/0!`/`#REF!`/循环引用
  ——这是**实打实的错表**（普通 `validate` 抓不到，因为它只看公式长相和缓存）。
- `checks` 里每条 `passed=false` 都是没满足你给的期望。
- 注意区分：`validation` 的 `warnings` 有时是**校验器自己误判**（例如把横向求和 `=SUM(B2:E2)` 当成少覆盖）——
  warning 不直接算失败，但值得你在诊断里留意是不是“工具层”问题。

## 4. 要测哪些方面（覆盖这些维度，每轮换着来）

1. **公式正确性（最重点）**：跨表引用（`成绩表!C2:C13`，别漏工作表名前缀）、纵向 vs 横向求和、
   VLOOKUP/INDEX-MATCH、SUMIF/SUMIFS/COUNTIFS、RANK 排名、占比/百分比、IF 多条件、日期/工龄计算、
   IFERROR/除零/空值边界。**重点看真算是否 0 错误。**
2. **多工作表与复杂结构**：参数表+明细+汇总联动、两级（合并）表头、小计/总计行、交叉汇总（透视味道）、
   多表跨表统计。
3. **图表**：柱/条/折线/面积/饼/环形/雷达/散点/组合；**按语义选型**（占比→饼、趋势→折线、对比→柱、
   多维→雷达、相关→散点）；用户**点名某类型时必须用对**；图表数据源要指向有真实数值的单元格。
4. **就地编辑 / 模板保真**：上传文件改某列/加一列、其它别动（`keep_sheets_from_input`）；上传空模板+数据按模板填。
5. **数据规模与边界**：几百行数据、空值/缺失、负数、0（除零）、超长文本、重复键、特殊字符、混合类型。
6. **真实业务领域**（轮着出，贴近真实语气）：教务（成绩统计/排考/课表/学分绩点）、销售（区域产品交叉/环比同比）、
   财务（预算执行/现金流/账龄）、人事（薪酬绩效/考勤/花名册）、库存（出入库/安全库存/周转）、
   项目（进度甘特味道/工时）、问卷统计、评分汇总等。
7. **稳定性 / 复现**：**同一道题连跑 2–3 次**（不同 slug），看结果是否一致——模型有随机性，
   “同题一次对一次错”本身就是要暴露的稳定性问题。
8. **需求贴合度（软指标，你来判断）**：harness 只查客观硬指标；产物是否**真的满足需求**（算的对不对、
   有没有答非所问、有没有过度设计/偷工减料）要你打开产物或读轨迹自己判断，发现就记为失败并归因。

## 5. 出题原则

- 每轮**只出一道**，覆盖一个或几个维度的组合；由易到难，逐步加复杂度。
- 用**真实业务语气**写需求（像教务老师/财务/HR 会怎么说），别写成测试用例腔。
- 给**合理的 `expect`**（你预期一个正确产物该满足什么），但别过苛刻到误杀。
- **不要重复**已出过的题；记录出过的维度/场景，保证长跑下覆盖面广。
- 优先设计**容易暴露问题**的题：跨表统计、就地编辑别动其它、点名图表类型、边界数据、多表联动。

## 6. 失败了怎么分析（核心价值）

每条失败结果里：
- `fail_reason`：为什么没过（真算报错 / 哪条期望没满足）。
- `task_dir`：那次任务目录，里面有两份关键文件——
  - `agent_trace.json`：模型**每一步写的代码** + 工具完整输出 + 真算关卡的退回/放行。
  - `diagnostic_report.md`：AI 把问题**归因到 模型/工具/代码/网络 层** + 证据 + 可能原因 + 改进方向。
- `diagnosis.problems`：诊断给出的分层归因（`layer` ∈ model/tool/code/network）。

把失败**按层级归类统计**：
- `model` 多 → 模型能力/提示词问题（写错公式、误解需求）；
- `tool` 多 → 工具问题（run_python 报错、真算/校验误判）；
- `code` 多 → 编排/校验/兜底逻辑缺陷（最值得反馈给开发者修）；
- `network` 多 → 连接/超时/依赖。

## 7. 循环模式（长时间自动跑）

```
loop:
  1) 设计 1 道新题（换维度/场景/难度，参照 §4 §5）
  2) run_one_case(prompt, input_files?, expect, slug=递增)
  3) 看 passed：
       - 通过 → 记一笔（维度、是否首次该场景）
       - 失败 → 读 diagnostic_report.md + agent_trace.json，把根因归到某层，记录典型案例
  4) 每 10 题：输出一次阶段汇总
       - 总通过率；按维度的通过率；按层级(model/tool/code/network)的失败分布；
       - 3–5 个最有价值的失败案例（题目 + 根因层 + 一句话原因 + task_dir）
  5) 继续，直到用户喊停
```

阶段汇总建议直接读 `outputs/_agent_test/results.jsonl` 统计（每行一条完整记录）。

## 8. 注意事项

- **每题都真实调豆包/deepseek**（慢 + 耗 token）：复杂题生成 ~60–180s，失败诊断再加 ~30–120s。
  按这个节奏安排，别同时并发太多（会触发限流）。
- **真算需要本机 LibreOffice**；没装时 harness 会跳过真算（`recalc.available=false`），判定会变宽，要留意。
- 所有产物/结果在 `outputs/_agent_test/`（已 gitignore，不污染仓库）。长跑可定期清理通过用例、只留失败案例。
- harness 单题异常会被捕获并记为失败（`fail_reason` 带 `harness异常`），不会中断你的循环。
