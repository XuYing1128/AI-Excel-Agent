"""根因诊断模块测试：组装上下文(含轨迹)、解析归因、降级——均用 mock，不依赖网络。"""

import json

from excel_agent.api_settings import ApiSettings
from excel_agent.services import diagnostic_report as dr
from excel_agent.services.custom_api_service import ApiCallResult
from excel_agent.task_paths import create_task_paths
from excel_agent.task_spec import TaskSpec


def _cfg_settings():
    return ApiSettings(enabled=True, base_url="https://e/v1", api_key="k", model="m", use_for_review=True)


def _spec():
    return TaskSpec(task_type="generic_table", user_goal="做一个各科平均分统计表")


def _common(tmp_path):
    paths = create_task_paths("generic_table", tmp_path)
    validation = {"status": "warn", "warnings": [{"check": "x", "message": "公式范围 B2:E2 可能少覆盖"}], "errors": []}
    recalc = {"available": True, "ok": True, "error_cells": []}
    wbsum = {"sheet_count": 1, "sheets": [{"name": "科目统计"}]}
    return paths, validation, recalc, wbsum


def test_diagnosis_context_includes_model_code_from_trace(tmp_path, monkeypatch):
    paths, validation, recalc, wbsum = _common(tmp_path)
    # 落一份轨迹：模型写了自引用公式
    (paths.task_dir / "agent_trace.json").write_text(
        json.dumps({"steps": [{"step": 1, "kind": "tool", "tool": "run_python",
                               "arguments": {"code": "ws['B2']='=AVERAGE(B2:B3)'"}}]}, ensure_ascii=False),
        encoding="utf-8",
    )
    captured = {}

    def fake_chat(settings, *, system_prompt, user_prompt, **kw):
        captured["user_prompt"] = user_prompt
        return ApiCallResult(True, json.dumps({"overall": "ok", "problems": []}), None, 200, 10)

    monkeypatch.setattr(dr, "_reviewer_settings", lambda api: _cfg_settings())
    dr.generate_diagnostic_report(
        _spec(), paths, validation_report=validation, recalc_result=recalc,
        workbook_summary=wbsum, chat_func=fake_chat,
    )
    # 轨迹里模型写的代码进了喂给诊断模型的上下文
    assert "执行轨迹" in captured["user_prompt"]
    assert "AVERAGE(B2:B3)" in captured["user_prompt"]


def test_diagnosis_normalizes_problems(tmp_path, monkeypatch):
    paths, validation, recalc, wbsum = _common(tmp_path)

    def fake_chat(settings, *, system_prompt, user_prompt, **kw):
        payload = {
            "overall": "发现一个模型层问题",
            "problems": [
                {"title": "自引用公式", "layer": "model", "evidence": "step1 B2",
                 "likely_cause": "平均分引用了含自身的范围", "suggestion": "引用数据列", "confidence": "high"},
                {"title": "存疑项", "layer": "胡乱写的层", "confidence": "超高"},  # 非法→规整
            ],
        }
        return ApiCallResult(True, json.dumps(payload, ensure_ascii=False), None, 200, 10)

    monkeypatch.setattr(dr, "_reviewer_settings", lambda api: _cfg_settings())
    report = dr.generate_diagnostic_report(
        _spec(), paths, validation_report=validation, recalc_result=recalc,
        workbook_summary=wbsum, chat_func=fake_chat,
    )
    problems = report["problems"]
    assert problems[0]["layer"] == "model" and problems[0]["confidence"] == "high"
    assert problems[1]["layer"] == "unknown" and problems[1]["confidence"] == "low"
    # markdown 也落盘
    assert (paths.task_dir / "diagnostic_report.md").exists()


def test_diagnosis_disabled_without_reviewer(tmp_path, monkeypatch):
    paths, validation, recalc, wbsum = _common(tmp_path)
    monkeypatch.setattr(dr, "_reviewer_settings", lambda api: ApiSettings())  # 未配置
    report = dr.generate_diagnostic_report(
        _spec(), paths, validation_report=validation, recalc_result=recalc,
        workbook_summary=wbsum, chat_func=lambda *a, **k: ApiCallResult(True, "{}", None, 200, 10),
    )
    assert report["enabled"] is False
    assert report["problems"] == []


def test_run_diagnostic_async_writes_report(tmp_path, monkeypatch):
    from openpyxl import Workbook

    from excel_agent.services import recalc as recalc_mod

    # 真算/网络都打桩，验证的是“后台线程能独立跑完并落盘”，不依赖 LibreOffice/网络。
    monkeypatch.setattr(recalc_mod, "recalc_workbook", lambda p, **k: {"available": True, "ok": True, "error_cells": []})
    monkeypatch.setattr(dr, "_reviewer_settings", lambda api: ApiSettings())  # 降级，不调网络
    paths = create_task_paths("generic_table", tmp_path)
    paths.output_file.parent.mkdir(parents=True, exist_ok=True)
    Workbook().save(paths.output_file)

    thread = dr.run_diagnostic_async(_spec(), paths)
    thread.join(timeout=30)
    assert not thread.is_alive()  # 后台线程独立跑完
    assert (paths.task_dir / "diagnostic_report.json").exists()  # 报告已落盘
