---
name: perf_review
triggers: 绩效,薪酬,工资,奖金,调薪,考核,缺勤扣分
tools: build_performance_compensation,build_rich_workbook,validate_workbook
---

# perf_review

用于绩效考核、工资薪酬、奖金调整和缺勤扣分等高风险表格。

## 流程

1. 识别参数表：指标权重、等级区间、调薪比例、扣分规则。
2. 识别明细表：部门、编号、姓名、当前薪资、缺勤天数和各指标评分。
3. 优先调用 `build_performance_compensation`，生成参数表 + 明细表。
4. 公式必须写入 Excel：加权总分、扣分、等级、排名、调薪比例、调整后薪资。
5. 生成后校验公式数量、表头、汇总行、条件格式和人工复核提醒。

## 风险提醒

绩效、工资、奖金和调薪结果必须提示“需要人工复核”，不能作为最终审批依据。

