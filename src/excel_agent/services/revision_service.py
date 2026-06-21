"""Create a new TaskSpec version from user or review feedback."""

from __future__ import annotations

import re
from copy import deepcopy
from pathlib import Path

from ..api_settings import ApiSettings
from ..content_plan import suggest_output_name
from ..task_spec import TaskSpec
from ..task_spec_builder import build_task_spec_draft
from .api_task_planner import enhance_task_spec_draft


def build_revision_task_spec(
    current: TaskSpec,
    revision_request: str,
    settings: ApiSettings,
) -> TaskSpec:
    request = str(revision_request or "").strip()
    if not request:
        raise ValueError("请先填写需要修改的内容。")

    combined_prompt = (
        f"{current.user_goal}\n\n本次修改要求：\n{request}\n"
        "请保留未被本次要求否定的原有内容。"
    )
    planning_used_api = False
    if settings.configured and settings.use_for_intent:
        draft = build_task_spec_draft(
            combined_prompt,
            [Path(item).name for item in current.input_files],
        )
        planning = enhance_task_spec_draft(
            draft,
            user_prompt=combined_prompt,
            input_file_names=[Path(item).name for item in current.input_files],
            settings=settings,
        )
        revised = planning.draft.task_spec
        planning_used_api = planning.used_api
        if (
            current.options.get("content_plan", {}).get("explicit_structure")
            and not revised.options.get("content_plan", {}).get("explicit_structure")
        ):
            revised.options["content_plan"] = deepcopy(current.options["content_plan"])
            revised.options["generation_policy"] = "custom_content"
    else:
        revised = TaskSpec.from_dict(deepcopy(current.to_dict()))
        revised.user_goal = combined_prompt

    _apply_local_revision_rules(revised, request)
    revised.input_files = list(current.input_files)
    revised.user_answers = deepcopy(current.user_answers)
    revised.user_answers["revision_request"] = request
    revised.options["source_task_id"] = current.options.get("task_id")
    revised.options["revision_index"] = int(current.options.get("revision_index", 1)) + 1
    revised.options["revision_source"] = "custom_api" if planning_used_api else "local_rules"

    old_stem = Path(current.output_name).stem
    if revised.output_name == current.output_name or revised.output_name == "自定义表格.xlsx":
        revised.output_name = f"{old_stem}_修改版{revised.options['revision_index']}.xlsx"
    elif not revised.output_name:
        revised.output_name = suggest_output_name(
            combined_prompt,
            revised.task_type,
            revised.options.get("content_plan", {}).get("title"),
        )
    return revised


def _apply_local_revision_rules(spec: TaskSpec, request: str) -> None:
    lowered = request.lower()
    plan = spec.options.get("content_plan", {})
    if not isinstance(plan, dict):
        plan = {}

    if any(word in lowered for word in ("删除图表", "去掉图表", "不要图表", "不需要图表")):
        spec.include_charts = False
    elif any(word in lowered for word in ("增加图表", "新增图表", "需要图表")):
        spec.include_charts = True

    if any(
        word in lowered
        for word in ("删除汇总页", "去掉汇总页", "不要汇总页", "删除独立汇总页")
    ):
        spec.include_summary = False
    if any(word in lowered for word in ("删除周平均", "去掉周平均", "不要周平均")):
        plan["summary_rules"] = []
    if any(
        word in lowered
        for word in ("一张工作表", "单个工作表", "不要拆分工作表", "不拆分工作表")
    ):
        plan["layout"] = "single_sheet"
        spec.include_instructions_sheet = False
        spec.include_summary = False

    title_match = re.search(
        r"(?:标题|表名)\s*(?:改为|修改为|设为)\s*[：:]?\s*([^\n，,。；;]+)",
        request,
    )
    if title_match:
        title = title_match.group(1).strip(" “”。")
        if title:
            plan["title"] = title
            plan["sheet_name"] = re.sub(r"[\[\]:*?/\\]", "_", title)[:31]

    columns = [dict(item) for item in plan.get("columns", []) if item.get("name")]
    remove_names = [
        item.strip(" “”。")
        for item in re.findall(
            r"(?:删除|去掉|不要)\s*[“\"]?([^”\"\n，,。]+?)[”\"]?\s*列",
            request,
        )
    ]
    if remove_names:
        columns = [
            item
            for item in columns
            if not any(name == item["name"] or name in item["name"] for name in remove_names)
        ]
        plan["formula_rules"] = [
            item
            for item in plan.get("formula_rules", [])
            if item.get("target") in {column["name"] for column in columns}
        ]

    move_matches = re.findall(r"把\s*[“\"]?([^”\"\n，,。]+?)[”\"]?\s*列放到最后", request)
    for name in move_matches:
        matching = [
            item for item in columns if name.strip() == item["name"] or name.strip() in item["name"]
        ]
        if matching:
            columns = [item for item in columns if item not in matching] + matching
    if columns:
        plan["columns"] = columns
        plan["explicit_structure"] = True
        spec.options["generation_policy"] = "custom_content"

    spec.options["content_plan"] = plan
