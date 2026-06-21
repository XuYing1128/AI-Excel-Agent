import json

from excel_agent.services.generation_service import generate_from_task_spec
from excel_agent.task_paths import create_task_paths
from excel_agent.task_spec import TaskSpec


def test_generation_service_uses_existing_builder(tmp_path):
    paths = create_task_paths("personal_budget", tmp_path / "tasks")
    spec = TaskSpec(
        task_type="personal_budget",
        user_goal="生成个人预算表",
        include_charts=False,
    )
    result = generate_from_task_spec(spec, paths)

    assert result.success is True
    assert paths.output_file.exists()
    assert paths.task_spec_file.exists()
    log = json.loads(paths.run_log_file.read_text(encoding="utf-8"))
    assert log["latest_status"] == "success"
    assert any(event["event"] == "generation_completed" for event in log["events"])


def test_generation_failure_is_logged_and_task_spec_is_preserved(tmp_path):
    paths = create_task_paths("sales_report", tmp_path / "tasks")
    spec = TaskSpec(
        task_type="sales_report",
        user_goal="根据缺失文件生成销售月报",
        input_files=[str(tmp_path / "missing.csv")],
    )
    result = generate_from_task_spec(spec, paths)

    assert result.success is False
    assert paths.task_spec_file.exists()
    assert paths.run_log_file.exists()
    log = json.loads(paths.run_log_file.read_text(encoding="utf-8"))
    assert log["latest_status"] == "error"
