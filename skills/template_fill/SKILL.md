---
name: template_fill
triggers: 模板,套用,严格,导入,参考格式,按模板
tools: fill_template,validate_workbook,inspect_workbook
---

# template_fill

用于“上传模板 + 上传数据”的任务。模板可能只是视觉参考，也可能是业务系统导入格式。

## 流程

1. 明确模板模式：reference / flexible / strict。
2. 模板示例数据默认不作为业务数据，除非用户明确选择。
3. strict：字段、工作表结构和顺序冲突时停止并说明差异。
4. flexible：保留模板形式，但需求优先，可增删必要内容。
5. 优先调用 `fill_template`，不要重新设计无关工作簿。
6. 生成后运行 `validate_workbook`，检查可打开、表头、公式、筛选和导入相关字段。

## 常见坑

- 不要把旧模板标题、旧考次、旧日期带入新结果。
- 不要把模板里的样例行当成真实业务数据。
- 数据文件与模板文件必须分开处理。

