---
name: data_clean
triggers: 清洗,去重,缺失,异常,合并,标准化,脏数据
tools: read_table_summary,run_python,build_dataset,validate_workbook
---

# data_clean

用于数据清洗、合并、去重、异常标记和 clean_report 任务。

## 流程

1. 先用 `read_table_summary` 读取列名、类型和样例，不向模型灌全量数据。
2. 规则明确时优先用确定性 pandas 处理；长尾转换可用 `run_python`。
3. 保留原始数据页，另建清洗数据页和 clean_report。
4. 记录清洗动作：空行空列、日期金额标准化、去重、缺失、异常。
5. 输出后运行 `validate_workbook`。

## 常见坑

- 字段口径不清楚时，不要猜；生成 clean_report 并提示用户确认。
- 中文列名要保留，不随意翻译。

