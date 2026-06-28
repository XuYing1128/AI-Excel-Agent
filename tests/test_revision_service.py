from openpyxl import load_workbook

from excel_agent.api_settings import ApiSettings
from excel_agent.services.generation_service import generate_from_task_spec
from excel_agent.services.revision_service import build_revision_task_spec
from excel_agent.services.validation_service import validate_generated_workbook
from excel_agent.task_paths import create_task_paths
from excel_agent.task_spec_builder import build_task_spec_draft


def test_revision_creates_new_version_name_and_keeps_original_goal():
    current = build_task_spec_draft("帮我生成个人月度收支预算表", []).task_spec
    current.options["revision_index"] = 1
    revised = build_revision_task_spec(
        current,
        "删除图表，并把备注列放到最后。",
        ApiSettings(),
    )
    assert "本次修改要求" in revised.user_goal
    assert revised.output_name.endswith("_修改版2.xlsx")
    assert revised.options["revision_index"] == 2


def test_revision_can_remove_summary_and_rename_custom_table():
    prompt = """主题：天气安排
表格需包含以下列：日期、城市、最高气温、最低气温、日均气温
日均气温取最高气温和最低气温平均值。
表格最后增加周平均。
数据如下：
6月15日,洛杉矶,29,18,
"""
    current = build_task_spec_draft(prompt, []).task_spec
    current.options["revision_index"] = 1
    revised = build_revision_task_spec(
        current,
        "删除周平均，标题改为西海岸天气安排，并且只保留一张工作表。",
        ApiSettings(),
    )
    plan = revised.options["content_plan"]
    assert plan["title"] == "西海岸天气安排"
    assert plan["summary_rules"] == []
    assert plan["layout"] == "single_sheet"
    assert revised.include_summary is False


def test_revision_request_for_chart_sets_chart_requirement():
    current = build_task_spec_draft("生成销售汇总表。", []).task_spec
    current.include_charts = False
    revised = build_revision_task_spec(
        current,
        "重新审查生成图表，增加柱状对比图。",
        ApiSettings(),
    )
    assert revised.include_charts is True
    assert revised.options["chart_requirements"]["required"] is True
    assert "column" in revised.options["chart_types"]


def test_revision_generates_a_new_valid_workbook_without_overwriting(tmp_path):
    prompt = """主题：天气安排
表格需包含以下列：日期、城市、最高气温、最低气温、日均气温
日均气温取最高气温和最低气温平均值。
数据如下：
2026-06-15,洛杉矶,29,18,
2026-06-16,旧金山,21,13,
"""
    original = build_task_spec_draft(prompt, []).task_spec
    original_paths = create_task_paths(
        original.task_type,
        tmp_path,
        output_name=original.output_name,
    )
    assert generate_from_task_spec(original, original_paths).success

    revised = build_revision_task_spec(
        original,
        "标题改为天气安排修改版，只保留一张工作表。",
        ApiSettings(),
    )
    revised_paths = create_task_paths(
        revised.task_type,
        tmp_path,
        output_name=revised.output_name,
    )
    assert generate_from_task_spec(revised, revised_paths).success
    validation = validate_generated_workbook(
        revised_paths.output_file,
        revised,
        revised_paths,
    )

    assert original_paths.output_file.exists()
    assert revised_paths.output_file.exists()
    assert original_paths.output_file != revised_paths.output_file
    assert validation.status == "pass"
    assert load_workbook(revised_paths.output_file).active["A1"].value == "天气安排修改版"
