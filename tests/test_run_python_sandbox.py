from excel_agent.services.agent.tools import ToolContext, registered_tools
from excel_agent.services.agent.tools.run_python import run_python_code, run_python_tool
from excel_agent.task_paths import create_task_paths
from excel_agent.task_spec import TaskSpec


def _ctx(tmp_path, *, enabled=True):
    paths = create_task_paths("generic_table", tmp_path / "tasks")
    return ToolContext(
        task_spec=TaskSpec(task_type="generic_table", user_goal="测试脚本"),
        task_paths=paths,
        run_python_enabled=enabled,
    )


def test_run_python_writes_only_task_output(tmp_path):
    ctx = _ctx(tmp_path)
    result = run_python_code(
        """
import json
with open(OUTPUT_DIR + "/result.json", "w", encoding="utf-8") as f:
    json.dump({"ok": True}, f)
print("done")
""",
        ctx,
        timeout_seconds=10,
    )

    assert result.ok is True
    assert (ctx.task_paths.output_dir / "result.json").exists()
    assert "done" in result.data["stdout"]
    assert any(path.endswith("result.json") for path in result.artifacts)


def test_run_python_blocks_unapproved_import(tmp_path):
    ctx = _ctx(tmp_path)
    result = run_python_tool().handler({"code": "import os\nprint(os.getcwd())"}, ctx)

    assert result.ok is False
    assert result.error == "import_not_allowed"
    assert "os" in result.data["invalid_imports"]


def test_run_python_rejects_path_escape(tmp_path):
    ctx = _ctx(tmp_path)
    outside = tmp_path / "outside.txt"
    result = run_python_code(
        f'open(r"{outside}", "w", encoding="utf-8").write("bad")',
        ctx,
        timeout_seconds=10,
    )

    assert result.ok is False
    assert not outside.exists()
    assert "path_escape" in result.data["stderr"]


def test_run_python_timeout_is_killed(tmp_path):
    ctx = _ctx(tmp_path)
    result = run_python_code("while True:\n    pass", ctx, timeout_seconds=1)

    assert result.ok is False
    assert result.error == "timeout"


def test_run_python_tool_hidden_when_disabled(tmp_path):
    ctx = _ctx(tmp_path, enabled=False)
    names = [tool.name for tool in registered_tools(ctx)]
    assert "run_python" not in names
