from excel_agent.memory_store import (
    clear_preferences,
    get_preference,
    learn_preferences_from_task,
    list_preferences,
    list_skill_versions,
    list_task_history,
    record_task_history,
    rollback_skill_version,
    save_skill_version,
    set_preference,
)
from excel_agent.task_spec import TaskSpec


def test_preferences_roundtrip_and_clear(tmp_path):
    db = tmp_path / "memory.db"
    set_preference("sheet_name_language", "中文", db)
    assert get_preference("sheet_name_language", path=db) == "中文"
    assert list_preferences(db)["sheet_name_language"] == "中文"
    clear_preferences(db)
    assert list_preferences(db) == {}


def test_task_history_roundtrip(tmp_path):
    db = tmp_path / "memory.db"
    record_task_history(
        task_id="task1",
        prompt="做销售表",
        task_type="sales_report",
        output_file="outputs/task1.xlsx",
        status="pass",
        path=db,
    )
    rows = list_task_history(path=db)
    assert rows[0]["task_id"] == "task1"
    assert rows[0]["task_type"] == "sales_report"


def test_skill_version_and_rollback(tmp_path):
    db = tmp_path / "memory.db"
    v1 = save_skill_version("demo", "第一版", path=db)
    v2 = save_skill_version("demo", "第二版", path=db)
    assert (v1, v2) == (1, 2)
    content = rollback_skill_version("demo", 1, path=db)
    versions = list_skill_versions("demo", path=db)
    enabled = [item for item in versions if item["enabled"]]
    assert content == "第一版"
    assert enabled[0]["version"] == 1


def test_learn_preferences_from_task_is_low_risk(tmp_path):
    db = tmp_path / "memory.db"
    spec = TaskSpec(
        task_type="dashboard",
        user_goal="做图表",
        include_charts=True,
        preserve_template_style=True,
    )
    learned = learn_preferences_from_task(
        spec,
        {"sheets": [{"name": "说明"}, {"name": "汇总"}]},
        path=db,
    )
    assert learned["prefer_charts_when_requested"] is True
    assert learned["prefer_template_style_when_uploaded"] is True
    assert get_preference("sheet_name_language", path=db) == "中文"

