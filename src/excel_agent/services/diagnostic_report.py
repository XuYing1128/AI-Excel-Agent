"""根因诊断：把问题归因到 模型 / 工具 / 代码 / 网络 哪一层，便于维护改进。

与 subjective_review 的区别：那个只看“结构摘要”给主观建议；这个吃**细粒度执行轨迹**
（模型每步写的代码、工具完整输出、真算关卡裁决）+ 校验 + 真算 + 产物结构，目标是
“指出问题在哪一层、给证据和可能原因、给把握程度”——不只是罗列现象，而是帮开发者定位。
没配诊断模型时安全降级（enabled=False，不阻塞）。
"""

from __future__ import annotations

import json
import threading
from pathlib import Path
from typing import Any, Callable

from ..api_settings import ApiSettings
from ..model_registry import get_role_api_settings
from ..task_paths import TaskPaths, append_run_log_event
from ..task_spec import TaskSpec
from .custom_api_service import ApiCallResult, chat_completion, parse_json_object

ChatFunc = Callable[..., ApiCallResult]

_VALID_LAYERS = {"model", "tool", "code", "network"}
_VALID_CONFIDENCE = {"high", "medium", "low"}
_LAYER_CN = {"model": "模型", "tool": "工具", "code": "代码", "network": "网络/环境", "unknown": "待定"}
_CONF_CN = {"high": "高", "medium": "中", "low": "低"}

DIAGNOSIS_SYSTEM_PROMPT = (
    "你是这套本地 Excel 智能体的根因诊断器，面向开发者维护，不是给最终用户的客气建议。"
    "给你的是这次任务的执行轨迹（模型每步写的代码、工具的完整输出、真算关卡的裁决）、"
    "确定性校验与 LibreOffice 真算结果、产物结构摘要、用户原始需求。"
    "请找出这次生成里『已经发生或可能存在』的问题，并把每个问题归因到具体一层：\n"
    "  · model（模型）：理解错需求、写错公式/代码、漏步骤、过度设计、反复改不对；\n"
    "  · tool（工具）：run_python 报错、真算/校验工具本身误判或漏报、工具返回异常；\n"
    "  · code（代码）：编排/校验/兜底逻辑缺陷、走错分支、信号缺失导致误判；\n"
    "  · network（网络/环境）：连接失败、超时、依赖缺失（如未装 LibreOffice）。\n"
    "硬要求：① 每个问题必须给证据，指明在轨迹第几步/哪个工具/哪个单元格/哪条校验；"
    "② 给『可能原因』和『改进方向』；③ 给把握程度 confidence(high/medium/low)，"
    "证据不足就标 low、绝不硬下结论；④ 注意分辨：校验/真算报的‘问题’也可能是工具自己误判"
    "（要结合轨迹里模型实际写的公式判断）；⑤ 没有明显问题就让 problems 为空数组，别硬凑。\n"
    "只返回 JSON：{\"overall\": \"一句话总评\", \"problems\": [{\"title\": \"简短问题名\", "
    "\"layer\": \"model|tool|code|network\", \"evidence\": \"证据(引用轨迹/单元格/校验)\", "
    "\"likely_cause\": \"可能原因\", \"suggestion\": \"改进方向\", \"confidence\": \"high|medium|low\"}]}"
)


def generate_diagnostic_report(
    task_spec: TaskSpec,
    task_paths: TaskPaths,
    *,
    validation_report: dict[str, Any],
    recalc_result: dict[str, Any],
    workbook_summary: dict[str, Any],
    generation_summary: dict[str, Any] | None = None,
    api_settings: ApiSettings | None = None,
    chat_func: ChatFunc = chat_completion,
) -> dict[str, Any]:
    """读轨迹+校验+真算，调诊断模型产出分层归因报告；写 diagnostic_report.json / .md。"""

    context = _build_context(
        task_spec, task_paths, validation_report, recalc_result, workbook_summary, generation_summary
    )
    settings = _reviewer_settings(api_settings)
    if not settings.configured or not settings.use_for_review:
        report = {
            "enabled": False,
            "reason": "未配置诊断（reviewer）模型，已跳过根因诊断。",
            "overall": "",
            "problems": [],
        }
        _save(report, task_paths)
        append_run_log_event(task_paths, event="diagnostic_report_skipped", status="skipped",
                             details={"reason": report["reason"]})
        return report

    response = chat_func(
        settings,
        system_prompt=DIAGNOSIS_SYSTEM_PROMPT,
        user_prompt=json.dumps(context, ensure_ascii=False),
        temperature=0.1,
        max_tokens=4000,
        json_mode=True,
    )
    if not response.success:
        report = {"enabled": True, "error": response.error, "overall": f"诊断调用失败：{response.error}", "problems": []}
    else:
        try:
            payload = parse_json_object(response.content)
            report = {
                "enabled": True,
                "model": settings.provider_name,
                "overall": str(payload.get("overall", "")),
                "problems": _normalize_problems(payload.get("problems", [])),
            }
        except (ValueError, TypeError) as exc:
            report = {"enabled": True, "error": str(exc), "overall": "诊断结果无法解析。", "problems": []}

    _save(report, task_paths)
    _save_markdown(report, task_paths)
    append_run_log_event(
        task_paths,
        event="diagnostic_report",
        status="success" if "error" not in report else "warning",
        details={"problem_count": len(report.get("problems", [])), "model": report.get("model")},
    )
    return report


def run_diagnostic_async(
    task_spec: TaskSpec,
    task_paths: TaskPaths,
    *,
    generation_summary: dict[str, Any] | None = None,
    api_settings: ApiSettings | None = None,
) -> threading.Thread:
    """在后台 daemon 线程里跑根因诊断：自己算 校验 + 真算 + 结构再归因，全程不阻塞主流程/出表。

    诊断只读产物、调模型、落盘报告，不触碰 UI，因此放后台线程安全。任何异常都被吞掉，
    绝不影响表格的生成与交付。
    """

    def _worker() -> None:
        try:
            from ..validators import inspect_workbook, validate_workbook
            from .recalc import recalc_workbook

            output = Path(task_paths.output_file)
            if not output.exists():
                return
            validation_report = validate_workbook(output)
            recalc_result = recalc_workbook(output)
            workbook_summary = inspect_workbook(output)
            workbook_summary["sheet_count"] = len(workbook_summary.get("sheets", []))
            generate_diagnostic_report(
                task_spec,
                task_paths,
                validation_report=validation_report,
                recalc_result=recalc_result,
                workbook_summary=workbook_summary,
                generation_summary=generation_summary,
                api_settings=api_settings,
            )
        except Exception:
            pass  # 后台诊断失败绝不影响主流程

    thread = threading.Thread(target=_worker, name="diagnostic", daemon=True)
    thread.start()
    return thread


def _build_context(
    task_spec: TaskSpec,
    task_paths: TaskPaths,
    validation_report: dict[str, Any],
    recalc_result: dict[str, Any],
    workbook_summary: dict[str, Any],
    generation_summary: dict[str, Any] | None,
) -> dict[str, Any]:
    trace = _load_trace(task_paths)
    warnings = validation_report.get("warnings", [])
    errors = validation_report.get("errors", [])
    return {
        "用户需求": task_spec.user_goal,
        "任务类型": task_spec.task_type,
        "生成方式": (generation_summary or {}).get("mode"),
        "执行轨迹": trace,
        "确定性校验": {
            "status": validation_report.get("status"),
            "warnings": [
                {"check": w.get("check"), "message": w.get("message"), "sheet": w.get("sheet"), "cell": w.get("cell")}
                for w in warnings
                if isinstance(w, dict)
            ][:25],
            "errors": [
                {"check": e.get("check"), "message": e.get("message"), "sheet": e.get("sheet"), "cell": e.get("cell")}
                for e in errors
                if isinstance(e, dict)
            ][:25],
        },
        "真算结果": {
            "available": recalc_result.get("available"),
            "ok": recalc_result.get("ok"),
            "error_cells": recalc_result.get("error_cells", [])[:30],
        },
        "产物结构": workbook_summary,
    }


def _load_trace(task_paths: TaskPaths) -> list[dict[str, Any]]:
    trace_file = Path(task_paths.task_dir) / "agent_trace.json"
    if not trace_file.exists():
        return []
    try:
        return json.loads(trace_file.read_text(encoding="utf-8")).get("steps", [])
    except (ValueError, OSError):
        return []


def _normalize_problems(problems: Any) -> list[dict[str, Any]]:
    if not isinstance(problems, list):
        return []
    result = []
    for item in problems:
        if not isinstance(item, dict):
            continue
        layer = str(item.get("layer", "")).lower()
        if layer not in _VALID_LAYERS:
            layer = "unknown"
        confidence = str(item.get("confidence", "low")).lower()
        if confidence not in _VALID_CONFIDENCE:
            confidence = "low"
        result.append(
            {
                "title": str(item.get("title", "")).strip(),
                "layer": layer,
                "evidence": str(item.get("evidence", "")).strip(),
                "likely_cause": str(item.get("likely_cause", "")).strip(),
                "suggestion": str(item.get("suggestion", "")).strip(),
                "confidence": confidence,
            }
        )
    return result


def _reviewer_settings(api_settings: ApiSettings | None) -> ApiSettings:
    reviewer = get_role_api_settings(
        "reviewer", use_for_intent=False, use_for_review=True, use_for_generation=False
    )
    if reviewer is not None and reviewer.configured and reviewer.use_for_review:
        return reviewer
    if api_settings is not None and api_settings.configured and api_settings.use_for_review:
        return api_settings
    return ApiSettings()


def _save(report: dict[str, Any], task_paths: TaskPaths) -> None:
    try:
        (Path(task_paths.task_dir) / "diagnostic_report.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except OSError:
        pass


def _save_markdown(report: dict[str, Any], task_paths: TaskPaths) -> None:
    lines = ["# 根因诊断报告", "", f"**总评**：{report.get('overall', '')}", ""]
    problems = report.get("problems", [])
    if not problems:
        lines.append("未发现明显问题（或证据不足以定位）。")
    else:
        for idx, p in enumerate(problems, 1):
            layer = _LAYER_CN.get(p.get("layer", "unknown"), p.get("layer"))
            conf = _CONF_CN.get(p.get("confidence", "low"), p.get("confidence"))
            lines += [
                f"## {idx}. {p.get('title', '')}  〔{layer}层 · 把握{conf}〕",
                f"- **证据**：{p.get('evidence', '')}",
                f"- **可能原因**：{p.get('likely_cause', '')}",
                f"- **改进方向**：{p.get('suggestion', '')}",
                "",
            ]
    try:
        (Path(task_paths.task_dir) / "diagnostic_report.md").write_text("\n".join(lines), encoding="utf-8")
    except OSError:
        pass
