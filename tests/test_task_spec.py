from excel_agent.task_spec import load_task_spec, save_task_spec
from excel_agent.task_spec_builder import (
    build_task_spec_draft,
    merge_user_answers_into_task_spec,
)


def test_task_spec_draft_and_roundtrip(tmp_path):
    draft = build_task_spec_draft(
        "根据订单数据做销售月报和图表",
        ["orders.csv"],
    )
    assert draft.task_spec.task_type == "sales_report"
    assert draft.task_spec.include_charts is True
    assert len(draft.clarifying_questions) == 1
    assert "销售报表和电商订单分析" in draft.clarifying_questions[0]

    path = tmp_path / "task_spec.json"
    save_task_spec(draft.task_spec, path)
    loaded = load_task_spec(path)
    assert loaded.to_dict() == draft.task_spec.to_dict()


def test_task_spec_requires_at_most_one_clarification_round():
    draft = build_task_spec_draft("做个表格", [])
    assert 1 <= len(draft.clarifying_questions) <= 5
    merged = merge_user_answers_into_task_spec(
        draft.task_spec,
        {
            "task_type": "inventory",
            "goal_detail": "需要库存预警和汇总",
            "data_mode": "template",
        },
    )
    assert merged.task_type == "inventory"
    assert merged.options["clarification_rounds"] == 1
    assert "库存预警" in merged.user_goal
