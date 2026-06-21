from excel_agent.task_spec_builder import (
    build_task_spec_draft,
    merge_user_answers_into_task_spec,
)


def test_vague_request_asks_for_business_and_columns():
    draft = build_task_spec_draft("做个表格", [])

    assert "这张表主要用于什么业务场景？" in draft.clarifying_questions
    assert "明细区必须包含哪些列？请按最终顺序填写。" in draft.clarifying_questions
    assert draft.task_spec.options["requirement_gaps"]


def test_known_template_does_not_force_unnecessary_questions():
    draft = build_task_spec_draft("帮我生成个人月度收支预算表", [])

    assert draft.clarifying_questions == []


def test_answers_rebuild_content_plan_for_confirmation():
    draft = build_task_spec_draft("做个表格", [])
    merged = merge_user_answers_into_task_spec(
        draft.task_spec,
        {
            "task_type": "sales_report",
            "clarifications": {
                "用途": "销售月报",
                "字段": "表格需包含以下列：日期、销售员、数量、单价、销售额",
            },
            "goal_detail": "销售额自动计算，并按销售员汇总。",
        },
    )

    plan = merged.options["content_plan"]
    assert [item["name"] for item in plan["columns"]] == [
        "日期",
        "销售员",
        "数量",
        "单价",
        "销售额",
    ]
    assert plan["explicit_structure"] is True
