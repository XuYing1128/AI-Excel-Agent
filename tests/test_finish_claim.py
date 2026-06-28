"""锁住 P0-2 修复：finish_task 在默认文件名不存在时，认领 output 目录里的别名产物；真没产物时仍如实报缺失。"""

import importlib

from openpyxl import Workbook

from excel_agent.services.agent.tools.base import ToolContext
from excel_agent.task_paths import create_task_paths
from excel_agent.task_spec import TaskSpec


def _finish_tool():
    et = importlib.import_module("excel_agent.services.agent.tools.excel_tools")
    return {t.name: t for t in et.excel_tools()}["finish_task"]


def _ctx(tmp_path):
    paths = create_task_paths("generic_table", tmp_path)
    paths.output_dir.mkdir(parents=True, exist_ok=True)
    ctx = ToolContext(task_spec=TaskSpec(task_type="generic_table", user_goal="x"), task_paths=paths)
    return ctx, paths


def test_finish_claims_alternate_named_output(tmp_path):
    ctx, paths = _ctx(tmp_path)
    alt = paths.output_dir / "员工名单补全结果.xlsx"  # run_python 自选的别名
    Workbook().save(alt)
    assert not paths.output_file.exists()

    result = _finish_tool().handler({"summary": "done"}, ctx)

    assert result.ok is True  # 认领了别名产物，不再误报 output_missing
    assert paths.output_file.exists()  # 已回填到标准 output_file


def test_finish_reports_missing_when_no_output(tmp_path):
    ctx, paths = _ctx(tmp_path)

    result = _finish_tool().handler({"summary": "done"}, ctx)

    assert result.ok is False  # 真没产物时仍如实报缺失，没把判定关死
    assert result.error == "output_missing"
