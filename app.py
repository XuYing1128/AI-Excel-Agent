"""Streamlit V1 for the local AI-Excel-Agent workflow."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import streamlit as st

from excel_agent.intent_classifier import SUPPORTED_TYPES, classify_intent
from excel_agent.manifest import append_manifest_record, build_manifest_record, recent_tasks
from excel_agent.services.generation_service import generate_from_task_spec
from excel_agent.services.subjective_review_service import run_subjective_review
from excel_agent.services.validation_service import validate_generated_workbook
from excel_agent.task_paths import (
    TaskPaths,
    append_run_log_event,
    create_task_paths,
    save_uploaded_bytes,
)
from excel_agent.task_spec import TaskSpec, save_task_spec
from excel_agent.task_spec_builder import (
    TYPE_LABELS,
    build_task_spec_draft,
    merge_user_answers_into_task_spec,
)
from excel_agent.validators import inspect_workbook


st.set_page_config(page_title="AI-Excel-Agent", page_icon="📊", layout="wide")


STATE_DEFAULTS: dict[str, Any] = {
    "user_prompt": "",
    "uploaded_files": [],
    "classification_result": None,
    "clarifying_questions": [],
    "task_spec": None,
    "confirmed": False,
    "task_paths": None,
    "generation_result": None,
    "validation_result": None,
    "subjective_review_result": None,
    "clarification_done": False,
    "analysis_nonce": 0,
}


def initialize_state() -> None:
    for key, default in STATE_DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = default


def reset_after_analysis() -> None:
    st.session_state.confirmed = False
    st.session_state.task_paths = None
    st.session_state.generation_result = None
    st.session_state.validation_result = None
    st.session_state.subjective_review_result = None


def analyze_request(prompt: str, uploads: list[Any]) -> None:
    if not prompt.strip() and not uploads:
        st.error("请先输入需求或上传文件。")
        return
    upload_snapshots = [
        {"name": item.name, "type": item.type, "data": item.getvalue()}
        for item in uploads
    ]
    draft = build_task_spec_draft(prompt, [item["name"] for item in upload_snapshots])
    classification = classify_intent(prompt)
    st.session_state.user_prompt = prompt
    st.session_state.uploaded_files = upload_snapshots
    st.session_state.classification_result = {
        "table_type": classification.table_type,
        "confidence": classification.confidence,
        "matched_keywords": classification.matched_keywords,
        "alternatives": draft.classification_alternatives,
    }
    st.session_state.clarifying_questions = draft.clarifying_questions
    st.session_state.task_spec = draft.task_spec
    st.session_state.clarification_done = not draft.needs_clarification
    st.session_state.analysis_nonce += 1
    reset_after_analysis()


def apply_clarification() -> None:
    spec: TaskSpec = st.session_state.task_spec
    nonce = st.session_state.analysis_nonce
    clarifications = {
        question: st.session_state.get(f"clarification_{nonce}_{index}", "")
        for index, question in enumerate(st.session_state.clarifying_questions)
    }
    answers = {
        "task_type": st.session_state.get(f"clarify_type_{nonce}", spec.task_type),
        "goal_detail": st.session_state.get(f"clarify_detail_{nonce}", ""),
        "data_mode": st.session_state.get(f"clarify_data_mode_{nonce}", "template"),
        "clarifications": clarifications,
    }
    st.session_state.task_spec = merge_user_answers_into_task_spec(spec, answers)
    st.session_state.clarification_done = True


def execute_generation() -> None:
    spec: TaskSpec = st.session_state.task_spec
    task_paths = create_task_paths(spec.task_type)
    try:
        staged_files: list[str] = []
        for upload in st.session_state.uploaded_files:
            saved = save_uploaded_bytes(upload["name"], upload["data"], task_paths)
            staged_files.append(str(saved))
        spec.input_files = staged_files
        save_task_spec(spec, task_paths.task_spec_file)

        generation = generate_from_task_spec(spec, task_paths)
        validation = validate_generated_workbook(task_paths.output_file, spec, task_paths)
        workbook_summary = {}
        if task_paths.output_file.exists():
            try:
                workbook_summary = inspect_workbook(task_paths.output_file)
                workbook_summary["sheet_count"] = len(workbook_summary.get("sheets", []))
            except Exception as exc:  # UI summary must not block downloads.
                workbook_summary = {"summary_error": f"{type(exc).__name__}: {exc}"}
        subjective = run_subjective_review(
            task_spec=spec,
            validation_summary={"status": validation.status, **validation.summary},
            workbook_summary=workbook_summary,
            generation_summary=generation.to_dict(),
            task_paths=task_paths,
        )
        status = validation.status if generation.success else "error"
        append_manifest_record(
            build_manifest_record(
                task_id=task_paths.task_id,
                task_type=spec.task_type,
                user_prompt=st.session_state.user_prompt,
                input_files=spec.input_files,
                output_file=str(task_paths.output_file) if task_paths.output_file.exists() else None,
                validation_report=str(task_paths.validation_report),
                status=status,
                error=generation.error,
            )
        )
        st.session_state.confirmed = True
        st.session_state.task_paths = task_paths
        st.session_state.generation_result = generation.to_dict()
        st.session_state.validation_result = validation.to_dict()
        st.session_state.subjective_review_result = subjective
    except Exception as exc:
        try:
            save_task_spec(spec, task_paths.task_spec_file)
        except Exception:
            pass
        append_run_log_event(
            task_paths,
            event="ui_task_failed",
            status="error",
            details={"error": f"{type(exc).__name__}: {exc}"},
        )
        fallback_report = {
            "status": "error",
            "file": str(task_paths.output_file),
            "summary": {},
            "errors": [{"check": "ui_generation", "message": f"{type(exc).__name__}: {exc}"}],
            "warnings": [],
            "suggestions": ["查看 run_log.json 并重新提交任务。"],
        }
        task_paths.validation_report.write_text(
            json.dumps(fallback_report, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        append_manifest_record(
            build_manifest_record(
                task_id=task_paths.task_id,
                task_type=spec.task_type,
                user_prompt=st.session_state.user_prompt,
                input_files=spec.input_files,
                output_file=None,
                validation_report=str(task_paths.validation_report),
                status="error",
                error=f"{type(exc).__name__}: {exc}",
            )
        )
        st.session_state.confirmed = True
        st.session_state.task_paths = task_paths
        st.session_state.generation_result = {
            "success": False,
            "output_file": None,
            "message": "任务执行失败。",
            "error": f"{type(exc).__name__}: {exc}",
            "used_command": None,
            "mode": "error",
            "notices": [],
        }
        st.session_state.validation_result = {
            "status": "error",
            "issues": fallback_report["errors"],
            "warnings": [],
            "suggestions": fallback_report["suggestions"],
            "report_file": str(task_paths.validation_report),
            "summary": {},
        }
        st.session_state.subjective_review_result = {
            "enabled": False,
            "reviews": [],
            "agreement": "not_run",
            "user_notice": "生成失败，未运行主观审查；不影响查看错误报告。",
        }


def render_task_spec(spec: TaskSpec) -> None:
    left, right = st.columns(2)
    with left:
        st.markdown(f"**表格类型：** {TYPE_LABELS.get(spec.task_type, spec.task_type)}")
        st.markdown(f"**用户目标：** {spec.user_goal}")
        st.markdown(f"**输出文件名：** {spec.output_name}")
        st.markdown(f"**分类置信度：** {spec.confidence:.0%}")
    with right:
        st.markdown(f"**保留模板样式：** {'是' if spec.preserve_template_style else '否'}")
        st.markdown(f"**包含汇总：** {'是' if spec.include_summary else '否'}")
        st.markdown(f"**包含图表：** {'是' if spec.include_charts else '否'}")
        st.markdown(f"**包含说明页：** {'是' if spec.include_instructions_sheet else '否'}")

    st.markdown("**输入文件：**")
    if spec.input_files:
        for item in spec.input_files:
            st.write(f"- {Path(item).name}")
    else:
        st.write("- 未提供，将使用标准模板/demo。")

    st.markdown("**系统假设：**")
    for assumption in spec.assumptions:
        st.write(f"- {assumption}")
    with st.expander("查看 TaskSpec JSON"):
        st.json(spec.to_dict())


def download_file(label: str, path: Path, mime: str, key: str) -> None:
    if path.exists():
        st.download_button(
            label=label,
            data=path.read_bytes(),
            file_name=path.name,
            mime=mime,
            key=key,
        )


initialize_state()

st.title("AI-Excel-Agent 本地表格生成工具")
st.caption(
    "页面仅运行在本机 127.0.0.1。文件默认在本地处理；"
    "大模型不直接操作单元格，生成和客观校验均由确定性 Python 代码完成。"
)

with st.container(border=True):
    st.subheader("1. 输入需求")
    prompt = st.text_area(
        "请描述你需要的表格",
        value=st.session_state.user_prompt,
        placeholder="例如：根据订单数据做销售月报和图表",
        height=120,
    )
    uploads = st.file_uploader(
        "上传本地数据（可选）",
        type=["csv", "xlsx", "xlsm"],
        accept_multiple_files=True,
        help="V1 支持 CSV、XLSX、XLSM；不支持旧版 XLS 上传。",
    )
    if st.button("分析需求", type="primary"):
        analyze_request(prompt, list(uploads or []))

classification = st.session_state.classification_result
spec = st.session_state.task_spec
if classification and spec:
    with st.container(border=True):
        st.subheader("2. 分类与一次性澄清")
        st.write(
            f"识别类型：**{TYPE_LABELS.get(classification['table_type'], classification['table_type'])}** "
            f"（置信度 {classification['confidence']:.0%}）"
        )
        if classification["matched_keywords"]:
            st.caption("命中关键词：" + "、".join(classification["matched_keywords"]))

        if st.session_state.clarifying_questions and not st.session_state.clarification_done:
            st.warning("信息不足，请一次性补充以下内容。提交后不会继续多轮追问。")
            nonce = st.session_state.analysis_nonce
            for index, question in enumerate(st.session_state.clarifying_questions):
                st.text_input(question, key=f"clarification_{nonce}_{index}")
            type_index = SUPPORTED_TYPES.index(spec.task_type)
            st.selectbox(
                "确认表格类型",
                SUPPORTED_TYPES,
                index=type_index,
                format_func=lambda value: TYPE_LABELS.get(value, value),
                key=f"clarify_type_{nonce}",
            )
            st.text_area("其他补充要求", key=f"clarify_detail_{nonce}")
            st.radio(
                "没有输入文件时",
                ["template", "upload"],
                format_func=lambda value: "先生成标准模板/demo" if value == "template" else "稍后补充原始数据",
                horizontal=True,
                key=f"clarify_data_mode_{nonce}",
            )
            if st.button("应用补充信息"):
                apply_clarification()
                st.rerun()
        else:
            st.success("需求信息已足够，可以确认 TaskSpec。")

if spec and st.session_state.clarification_done and not st.session_state.confirmed:
    nonce = st.session_state.analysis_nonce
    with st.container(border=True):
        st.subheader("3. 确认 TaskSpec")
        spec.output_name = st.text_input(
            "输出显示名称",
            value=spec.output_name,
            key=f"output_name_{nonce}",
        )
        option_cols = st.columns(4)
        spec.preserve_template_style = option_cols[0].checkbox(
            "保留模板样式",
            value=spec.preserve_template_style,
            key=f"preserve_{nonce}",
        )
        spec.include_summary = option_cols[1].checkbox(
            "包含汇总",
            value=spec.include_summary,
            key=f"summary_{nonce}",
        )
        spec.include_charts = option_cols[2].checkbox(
            "包含图表",
            value=spec.include_charts,
            key=f"charts_{nonce}",
        )
        spec.include_instructions_sheet = option_cols[3].checkbox(
            "包含说明页",
            value=spec.include_instructions_sheet,
            key=f"instructions_{nonce}",
        )
        st.session_state.task_spec = spec
        render_task_spec(spec)
        confirm_col, modify_col = st.columns(2)
        if confirm_col.button("确认并生成", type="primary", use_container_width=True):
            with st.spinner("正在调用本地确定性内核生成并校验……"):
                execute_generation()
            st.rerun()
        if modify_col.button("返回修改", use_container_width=True):
            st.session_state.task_spec = None
            st.session_state.classification_result = None
            st.session_state.clarifying_questions = []
            st.session_state.clarification_done = False
            reset_after_analysis()
            st.rerun()

if st.session_state.confirmed and st.session_state.task_paths:
    task_paths: TaskPaths = st.session_state.task_paths
    generation = st.session_state.generation_result or {}
    validation = st.session_state.validation_result or {}
    subjective = st.session_state.subjective_review_result or {}

    with st.container(border=True):
        st.subheader("4. 生成结果")
        st.code(f"task_id: {task_paths.task_id}\n任务目录: {task_paths.task_dir}")
        if generation.get("success"):
            st.success(generation.get("message", "生成成功"))
        else:
            st.error(generation.get("error") or generation.get("message", "生成失败"))
        for notice in generation.get("notices", []):
            st.info(notice)

    with st.container(border=True):
        st.subheader("5. 确定性校验")
        status = str(validation.get("status", "error"))
        if status == "pass":
            st.success("pass：未发现确定性错误或警告。")
        elif status == "warn":
            st.warning("warn：文件可下载，但存在需要查看的警告。")
        else:
            st.error(f"{status}：存在错误；文件仍可下载用于人工检查。")
        issues = validation.get("issues", [])
        warnings = validation.get("warnings", [])
        suggestions = validation.get("suggestions", [])
        if issues:
            st.markdown("**错误/问题**")
            st.json(issues)
        if warnings:
            st.markdown("**警告**")
            st.json(warnings)
        if suggestions:
            st.markdown("**建议**")
            for item in suggestions:
                st.write(f"- {item}")

    with st.container(border=True):
        st.subheader("6. 主观审查")
        st.info(subjective.get("user_notice", "主观模型审查未启用，不影响文件下载。"))
        with st.expander("查看主观审查 JSON"):
            st.json(subjective)

    with st.container(border=True):
        st.subheader("7. 下载")
        download_cols = st.columns(4)
        with download_cols[0]:
            download_file(
                "下载 Excel",
                task_paths.output_file,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                f"download_excel_{task_paths.task_id}",
            )
        with download_cols[1]:
            download_file(
                "下载 validation.json",
                task_paths.validation_report,
                "application/json",
                f"download_validation_{task_paths.task_id}",
            )
        with download_cols[2]:
            download_file(
                "下载 task_spec.json",
                task_paths.task_spec_file,
                "application/json",
                f"download_spec_{task_paths.task_id}",
            )
        with download_cols[3]:
            download_file(
                "下载 subjective_review.json",
                task_paths.subjective_review_report,
                "application/json",
                f"download_review_{task_paths.task_id}",
            )

st.divider()
st.subheader("最近 10 个任务")
history = recent_tasks(10)
if history:
    st.dataframe(
        [
            {
                "task_id": item.get("task_id"),
                "task_type": item.get("task_type"),
                "created_at": item.get("created_at"),
                "status": item.get("status"),
            }
            for item in history
        ],
        use_container_width=True,
        hide_index=True,
    )
else:
    st.caption("还没有 V1 任务记录。")
