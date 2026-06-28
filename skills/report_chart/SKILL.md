---
name: report_chart
triggers: 图表,仪表盘,dashboard,趋势,占比,top,排名,报表
tools: build_rich_workbook,build_dataset,render_preview,validate_workbook
---

# report_chart

用于销售报表、电商分析、经营看板、趋势图、占比图和 Top N 排名。

## 流程

1. 先确认图表是否是用户硬性要求；确认后不得审查建议删除。
2. 区分数据源 sheet 和展示 sheet；Dashboard 必须分离。
3. 图表数据源使用稳定的值区域，避免只引用未重算隐藏公式导致空图。
4. 根据语义选择图表：趋势用折线，排名/对比用柱状，占比用饼/环，双指标可组合图。
5. 图表生成后用 `render_preview` 或结构摘要检查是否存在、范围是否清楚。
6. 最终运行 `validate_workbook`。

## 常见坑

- 图表不能只剩空坐标轴。
- 小计/总计是否进入图表必须按用户要求；默认图表排除小计和总计。

