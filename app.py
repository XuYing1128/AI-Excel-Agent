"""AI-Excel-Agent 本地中文网页。"""

from __future__ import annotations

import html
import json
import time
from copy import deepcopy
from pathlib import Path
from typing import Any

import streamlit as st

from excel_agent.api_settings import (
    ApiSettings,
    delete_api_settings,
    load_api_settings,
    mask_api_key,
    save_api_settings,
)
from excel_agent.chart_spec import CHART_TYPE_LABELS, chart_type_from_text
from excel_agent.intent_classifier import SUPPORTED_TYPES, classify_intent
from excel_agent.manifest import (
    append_manifest_record,
    build_manifest_record,
    recent_tasks,
)
from excel_agent.memory_store import (
    clear_preferences,
    learn_preferences_from_task,
    list_preferences,
    list_task_history,
    record_task_history,
)
from excel_agent.model_registry import (
    ROLE_LABELS,
    ROLE_NAMES,
    ModelSettings,
    ProviderConfig,
    delete_model_settings,
    from_legacy_api_settings,
    get_provider,
    load_model_settings,
    save_model_settings,
    safe_provider_id,
    test_provider,
)
from excel_agent.model_presets import (
    PRESET_BY_KEY,
    PROVIDER_PRESETS,
    detect_provider_key,
)
from excel_agent.preview import (
    combined_revision_prompt,
    sheet_preview_dataframe,
    workbook_chart_preview,
    workbook_preview,
)
from excel_agent.services.api_task_planner import enhance_task_spec_draft
from excel_agent.services.custom_api_service import test_api_connection
from excel_agent.services.revision_service import build_revision_task_spec
from excel_agent.services.diagnostic_report import run_diagnostic_async
from excel_agent.services.runtime_compat import load_generation_service
from excel_agent.services.subjective_review_service import run_subjective_review
from excel_agent.services.validation_service import validate_generated_workbook
from excel_agent.task_paths import (
    TaskPaths,
    append_run_log_event,
    create_task_paths,
    existing_task_paths,
    save_uploaded_bytes,
)
from excel_agent.task_spec import TaskSpec, TaskSpecDraft, load_task_spec, save_task_spec
from excel_agent.task_spec_builder import (
    TYPE_LABELS,
    build_task_spec_draft,
    merge_user_answers_into_task_spec,
)
from excel_agent.validators import inspect_workbook


st.set_page_config(
    page_title="本地表格助手",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
    menu_items={},
)


STATE_DEFAULTS: dict[str, Any] = {
    "active_page": "workbench",
    "user_prompt": "",
    "uploaded_files": [],
    "uploaded_data_files": [],
    "uploaded_template_files": [],
    "pasted_data_text": "",
    "classification_result": None,
    "clarifying_questions": [],
    "task_spec": None,
    "clarification_done": False,
    "analysis_message": "",
    "analysis_nonce": 0,
    "task_paths": None,
    "generation_result": None,
    "validation_result": None,
    "subjective_review_result": None,
    "api_settings": None,
    "model_settings": None,
    "api_message": "",
    "api_message_kind": "info",
    "api_edit_mode": False,
    "provider_editor_open": False,
    "provider_test_results": {},
    "versions": [],
    "revision_request": "",
    "revision_output_name": "",
    "onboarded": False,
}


def initialize_state() -> None:
    for key, default in STATE_DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = deepcopy(default)
    if st.session_state.api_settings is None:
        st.session_state.api_settings = load_api_settings()
    if st.session_state.model_settings is None:
        st.session_state.model_settings = load_model_settings()


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        :root {
            --brand: #4f46e5;
            --brand-dark: #4338ca;
            --brand-soft: #eef2ff;
            --accent: #6366f1;
            --ink: #1e2333;
            --muted: #6b7384;
            --line: #e7eaf3;
            --canvas: #f4f6fb;
            --ok: #16a34a;
            --warn: #d97706;
            --bad: #dc2626;
        }
        html, body, [class*="css"] {
            font-family: "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", -apple-system, sans-serif;
        }
        .stApp {
            background:
                radial-gradient(1100px 480px at 12% -8%, #eef1ff 0%, rgba(238,241,255,0) 55%),
                radial-gradient(900px 420px at 100% 0%, #eafff4 0%, rgba(234,255,244,0) 50%),
                var(--canvas);
            color: var(--ink);
        }
        .block-container { max-width: 1320px; padding-top: 1.1rem; padding-bottom: 4rem; padding-left: 2.4rem; padding-right: 2.4rem; }
        header[data-testid="stHeader"],
        [data-testid="stToolbar"], [data-testid="stDecoration"],
        [data-testid="stStatusWidget"],
        #MainMenu, footer, [data-testid="stSidebar"] { display: none !important; }

        /* Hero header */
        .page-head {
            margin: 6px 0 22px;
            padding: 30px 34px;
            background: linear-gradient(120deg, #4f46e5 0%, #6366f1 45%, #7c83f5 100%);
            border-radius: 22px;
            box-shadow: 0 18px 40px -18px rgba(79,70,229,.55);
            position: relative; overflow: hidden;
        }
        .page-head::after {
            content: ""; position: absolute; right: -40px; top: -60px;
            width: 220px; height: 220px; border-radius: 50%;
            background: rgba(255,255,255,.10);
        }
        .page-title { font-size: 28px; line-height: 1.25; font-weight: 800; color: #fff; margin-bottom: 8px; }
        .page-subtitle { color: rgba(255,255,255,.86); font-size: 14.5px; line-height: 1.7; max-width: 760px; }

        /* Step indicator */
        .steps { display: flex; gap: 10px; flex-wrap: wrap; margin: 2px 0 18px; }
        .step {
            display: flex; align-items: center; gap: 8px;
            padding: 7px 14px; border-radius: 999px; font-size: 13px; font-weight: 600;
            background: #fff; color: var(--muted); border: 1px solid var(--line);
        }
        .step .dot {
            width: 20px; height: 20px; border-radius: 50%; display: grid; place-items: center;
            font-size: 12px; background: #eef1f6; color: var(--muted); font-weight: 700;
        }
        .step.active { background: var(--brand); color: #fff; border-color: var(--brand);
            box-shadow: 0 8px 18px -8px rgba(79,70,229,.6); }
        .step.active .dot { background: rgba(255,255,255,.25); color: #fff; }
        .step.done { color: var(--ok); border-color: #c7ecd3; background: #f2fcf5; }
        .step.done .dot { background: var(--ok); color: #fff; }

        /* Cards */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background: #fff; border: 1px solid var(--line);
            border-radius: 18px; box-shadow: 0 10px 30px -22px rgba(30,40,90,.35);
        }
        h1, h2, h3 { color: var(--ink); letter-spacing: -.01em; }
        h3 { font-size: 19px !important; font-weight: 750 !important; }

        /* Buttons */
        .stButton > button, .stDownloadButton > button {
            border-radius: 11px; min-height: 44px; font-weight: 700; transition: all .15s ease;
        }
        .stButton > button[kind="primary"], .stDownloadButton > button {
            background: linear-gradient(120deg, var(--brand) 0%, var(--accent) 100%);
            border: none; color: #fff; box-shadow: 0 10px 22px -12px rgba(79,70,229,.7);
        }
        .stButton > button[kind="primary"]:hover, .stDownloadButton > button:hover {
            transform: translateY(-1px); box-shadow: 0 14px 26px -12px rgba(79,70,229,.85);
        }
        .stButton > button[kind="secondary"] {
            background: #fff; border: 1px solid var(--line); color: var(--ink);
        }
        .stButton > button[kind="secondary"]:hover { border-color: var(--accent); color: var(--brand); }

        /* Inputs */
        div[data-testid="stTextArea"] textarea,
        div[data-testid="stTextInput"] input {
            border-radius: 11px !important; border: 1px solid var(--line) !important;
        }
        div[data-testid="stTextArea"] textarea:focus,
        div[data-testid="stTextInput"] input:focus {
            border-color: var(--accent) !important; box-shadow: 0 0 0 3px var(--brand-soft) !important;
        }
        [data-testid="stFileUploaderDropzone"] {
            background: #f8faff; border: 1.5px dashed #c3cdf0; border-radius: 14px;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] { display: none; }
        [data-testid="stFileUploaderDropzone"] button > * { display: none !important; }
        [data-testid="stFileUploaderDropzone"] button::after { content: "选择本地文件"; font-size: 14px; }
        [data-testid="stFileUploaderDropzone"] button { font-size: 0; }
        div[data-testid="stAlert"] { border-radius: 13px; }

        /* Tabs as pills */
        button[data-baseweb="tab"] { font-weight: 650; }
        [data-baseweb="tab-list"] { gap: 4px; }

        /* Plan cards */
        .plan-card {
            border: 1px solid var(--line); border-radius: 14px; padding: 15px 17px;
            background: linear-gradient(180deg, #fbfcff 0%, #f7f9ff 100%); min-height: 100px;
        }
        .plan-label { color: var(--muted); font-size: 12px; margin-bottom: 7px; letter-spacing: .02em; }
        .plan-value { color: var(--ink); font-size: 15px; line-height: 1.55; font-weight: 700; }
        .soft-note {
            padding: 13px 15px; background: var(--brand-soft); border-left: 3px solid var(--accent);
            border-radius: 10px; color: #4451b5; line-height: 1.7; font-size: 13.5px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_navigation() -> None:
    workbench, history, settings, spacer = st.columns([1, 1, 1, 5])
    if workbench.button(
        "制作表格",
        type="primary" if st.session_state.active_page == "workbench" else "secondary",
        width="stretch",
    ):
        st.session_state.active_page = "workbench"
        st.rerun()
    if settings.button(
        "接口设置",
        type="primary" if st.session_state.active_page == "settings" else "secondary",
        width="stretch",
    ):
        st.session_state.active_page = "settings"
        st.rerun()
    if history.button(
        "最近文件",
        type="primary" if st.session_state.active_page == "history" else "secondary",
        width="stretch",
    ):
        st.session_state.active_page = "history"
        st.rerun()


def render_header(title: str, subtitle: str) -> None:
    st.markdown(
        f"""
        <div class="page-head">
          <div class="page-title">{html.escape(title)}</div>
          <div class="page-subtitle">{html.escape(subtitle)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_steps(active: int) -> None:
    """Show the 1->4 progress chips so a non-technical user always knows where they are."""

    names = ["描述需求", "完善要求", "确认生成", "查看结果"]
    chips = []
    for index, name in enumerate(names, start=1):
        cls = "step"
        if index < active:
            cls += " done"
            mark = "✓"
        elif index == active:
            cls += " active"
            mark = str(index)
        else:
            mark = str(index)
        chips.append(
            f'<div class="{cls}"><span class="dot">{mark}</span>{html.escape(name)}</div>'
        )
    st.markdown('<div class="steps">' + "".join(chips) + "</div>", unsafe_allow_html=True)


def _onboard_marker() -> Path:
    return Path("config") / ".onboarded"


def show_onboarding_if_first_use() -> None:
    """首次使用时弹一次新手引导；看完写标记，之后不再弹。"""
    if st.session_state.get("onboarded") or _onboard_marker().exists():
        return

    @st.dialog("👋 欢迎使用本地表格助手")
    def _guide() -> None:
        st.markdown(
            "用大白话把你要的表说清楚，AI 帮你做成 Excel。**三步上手：**\n\n"
            "1. **接口设置**（只需一次）：填好模型接口和密钥（如豆包 / DeepSeek）。\n"
            "2. **制作表格**：用中文描述你要的表，可上传数据或模板。\n"
            "3. **生成 → 下载**：点生成，AI 做表并自动校验公式，完成后直接下载。\n\n"
            "贴心：每次生成后台会自动做一次「问题诊断」，报告集中在项目的 `diagnostics/` "
            "文件夹，方便回看和分享。"
        )
        if st.button("开始使用 →", type="primary", width="stretch"):
            try:
                marker = _onboard_marker()
                marker.parent.mkdir(parents=True, exist_ok=True)
                marker.write_text("ok", encoding="utf-8")
            except OSError:
                pass
            st.session_state.onboarded = True
            st.rerun()

    _guide()


def reset_generated_result() -> None:
    st.session_state.task_paths = None
    st.session_state.generation_result = None
    st.session_state.validation_result = None
    st.session_state.subjective_review_result = None
    st.session_state.revision_request = ""
    st.session_state.revision_output_name = ""


def start_new_task() -> None:
    st.session_state.user_prompt = ""
    st.session_state.uploaded_files = []
    st.session_state.uploaded_data_files = []
    st.session_state.uploaded_template_files = []
    st.session_state.pasted_data_text = ""
    for widget_key in (
        "data_file_uploader",
        "template_file_uploader",
        "pasted_data_text_input",
        "template_mode_selector",
        "use_template_data_checkbox",
    ):
        st.session_state.pop(widget_key, None)
    st.session_state.classification_result = None
    st.session_state.clarifying_questions = []
    st.session_state.task_spec = None
    st.session_state.clarification_done = False
    st.session_state.analysis_message = ""
    reset_generated_result()


def analyze_request(
    prompt: str,
    data_uploads: list[Any],
    template_uploads: list[Any],
    pasted_data_text: str,
    template_mode: str,
    use_template_data: bool,
) -> None:
    if not prompt.strip() and not data_uploads and not pasted_data_text.strip():
        st.error("请先描述要制作的表格，或上传数据文件。")
        return
    data_snapshots = [
        {"name": item.name, "type": item.type, "data": item.getvalue()}
        for item in data_uploads
    ]
    template_snapshots = [
        {"name": item.name, "type": item.type, "data": item.getvalue()}
        for item in template_uploads
    ]
    effective_prompt = prompt.strip()
    if pasted_data_text.strip():
        effective_prompt = (
            f"{effective_prompt}\n\n用户粘贴的数据如下：\n{pasted_data_text.strip()}"
        ).strip()
    filenames = [item["name"] for item in data_snapshots]
    draft = build_task_spec_draft(effective_prompt, filenames)
    draft.task_spec.template_files = [
        item["name"] for item in template_snapshots
    ]
    draft.task_spec.preserve_template_style = bool(template_snapshots)
    draft.task_spec.options.update(
        {
            "template_mode": template_mode if template_snapshots else "none",
            "use_template_data": bool(use_template_data and template_snapshots),
            "template_file_names": [
                item["name"] for item in template_snapshots
            ],
            "pasted_data_text": bool(pasted_data_text.strip()),
        }
    )
    api_plan = enhance_task_spec_draft(
        draft,
        user_prompt=effective_prompt,
        input_file_names=filenames,
        settings=st.session_state.api_settings,
    )
    local_result = classify_intent(effective_prompt)
    spec = api_plan.draft.task_spec
    st.session_state.user_prompt = prompt
    st.session_state.uploaded_files = data_snapshots
    st.session_state.uploaded_data_files = data_snapshots
    st.session_state.uploaded_template_files = template_snapshots
    st.session_state.pasted_data_text = pasted_data_text
    st.session_state.classification_result = {
        "table_type": spec.task_type,
        "confidence": spec.confidence,
        "matched_keywords": spec.options.get(
            "classification_keywords", local_result.matched_keywords
        ),
    }
    st.session_state.clarifying_questions = api_plan.draft.clarifying_questions
    st.session_state.task_spec = spec
    st.session_state.clarification_done = not api_plan.draft.needs_clarification
    st.session_state.analysis_message = api_plan.message
    st.session_state.analysis_nonce += 1
    reset_generated_result()


def apply_clarification() -> None:
    spec: TaskSpec = st.session_state.task_spec
    nonce = st.session_state.analysis_nonce
    answers = {
        "task_type": st.session_state.get(f"clarify_type_{nonce}", spec.task_type),
        "goal_detail": st.session_state.get(f"clarify_detail_{nonce}", ""),
        "data_mode": st.session_state.get(f"clarify_data_mode_{nonce}", "template"),
        "clarifications": {
            question: st.session_state.get(f"clarification_{nonce}_{index}", "")
            for index, question in enumerate(st.session_state.clarifying_questions)
        },
    }
    merged = merge_user_answers_into_task_spec(spec, answers)
    refined = enhance_task_spec_draft(
        TaskSpecDraft(
            task_spec=merged,
            clarifying_questions=[],
            classification_alternatives=[],
        ),
        user_prompt=merged.user_goal,
        input_file_names=[
            item["name"] for item in st.session_state.uploaded_data_files
        ],
        settings=st.session_state.api_settings,
    )
    st.session_state.task_spec = refined.draft.task_spec
    st.session_state.analysis_message = refined.message
    st.session_state.clarification_done = True


def execute_generation(
    spec: TaskSpec,
    progress_callback: Any | None = None,
) -> None:
    started_at = time.perf_counter()

    def progress(stage: str, message: str) -> None:
        if progress_callback:
            # 计时统一交给 st.spinner(show_time=True) 的实时秒表；这里不再拼"已运行X秒"——
            # 否则会和实时秒表打架，而且阻塞期间这个数字是卡住的，反而显得坏掉。
            progress_callback(stage, message)

    progress("prepare", "任务已接收，正在准备输入文件和生成目录……")
    spec.output_name = Path(str(spec.output_name or "自定义表格.xlsx")).name
    if not spec.output_name.lower().endswith(".xlsx"):
        spec.output_name = f"{spec.output_name}.xlsx"
    task_paths = create_task_paths(spec.task_type, output_name=spec.output_name)
    spec.options["task_id"] = task_paths.task_id
    try:
        if st.session_state.uploaded_data_files:
            progress("input", "正在安全复制上传文件，原文件不会被覆盖……")
            staged_files = []
            for upload in st.session_state.uploaded_data_files:
                saved = save_uploaded_bytes(
                    upload["name"],
                    upload["data"],
                    task_paths,
                    subdirectory="data",
                )
                staged_files.append(str(saved))
            spec.input_files = staged_files
        if st.session_state.uploaded_template_files:
            progress("input", "正在复制模板文件；模板中的示例数据默认不会作为业务数据……")
            staged_templates = []
            for upload in st.session_state.uploaded_template_files:
                saved = save_uploaded_bytes(
                    upload["name"],
                    upload["data"],
                    task_paths,
                    subdirectory="templates",
                )
                staged_templates.append(str(saved))
            spec.template_files = staged_templates
        save_task_spec(spec, task_paths.task_spec_file)

        generation_service = load_generation_service()
        generation = generation_service.generate_from_task_spec(
            spec,
            task_paths,
            api_settings=st.session_state.api_settings,
            progress=progress,
        )
        progress("validate", "正在执行确定性质量检查和需求一致性检查……")
        validation = validate_generated_workbook(task_paths.output_file, spec, task_paths)
        workbook_summary: dict[str, Any] = {}
        if task_paths.output_file.exists():
            workbook_summary = inspect_workbook(task_paths.output_file)
            workbook_summary["sheet_count"] = len(workbook_summary.get("sheets", []))
        progress("review", "正在根据最终确认需求审查生成结果……")
        subjective = run_subjective_review(
            task_spec=spec,
            validation_summary={"status": validation.status, **validation.summary},
            workbook_summary=workbook_summary,
            generation_summary=generation.to_dict(),
            task_paths=task_paths,
            api_settings=st.session_state.api_settings,
        )
        # 每次生成后在后台独立跑根因诊断，落盘 diagnostic_report.md/.json，不阻塞出表与交付。
        run_diagnostic_async(
            spec,
            task_paths,
            generation_summary=generation.to_dict(),
            api_settings=st.session_state.api_settings,
        )
        status = validation.status if generation.success else "error"
        append_manifest_record(
            build_manifest_record(
                task_id=task_paths.task_id,
                task_type=spec.task_type,
                user_prompt=spec.user_goal,
                input_files=spec.input_files,
                output_file=str(task_paths.output_file)
                if task_paths.output_file.exists()
                else None,
                validation_report=str(task_paths.validation_report),
                status=status,
                error=generation.error,
            )
        )
        record_task_history(
            task_id=task_paths.task_id,
            prompt=spec.user_goal,
            task_type=spec.task_type,
            output_file=str(task_paths.output_file)
            if task_paths.output_file.exists()
            else None,
            status=status,
        )
        learned_preferences = learn_preferences_from_task(spec, workbook_summary)
        st.session_state.task_spec = spec
        st.session_state.task_paths = task_paths
        generation_data = generation.to_dict()
        generation_data["elapsed_seconds"] = round(
            time.perf_counter() - started_at,
            1,
        )
        generation_data["learned_preferences"] = learned_preferences
        st.session_state.generation_result = generation_data
        st.session_state.validation_result = validation.to_dict()
        st.session_state.subjective_review_result = subjective
        st.session_state.versions.append(
            {
                "task_id": task_paths.task_id,
                "output_name": spec.output_name,
                "output_file": str(task_paths.output_file),
                "status": status,
                "revision_index": int(spec.options.get("revision_index", 1)),
            }
        )
        progress("complete", "生成、检查和审查已完成。")
    except Exception as exc:
        append_run_log_event(
            task_paths,
            event="ui_task_failed",
            status="error",
            details={"error": f"{type(exc).__name__}: {exc}"},
        )
        st.session_state.task_paths = task_paths
        st.session_state.generation_result = {
            "success": False,
            "message": "生成失败。",
            "error": f"{type(exc).__name__}: {exc}",
            "notices": [],
            "elapsed_seconds": round(time.perf_counter() - started_at, 1),
        }
        st.session_state.validation_result = {
            "status": "error",
            "issues": [{"message": f"{type(exc).__name__}: {exc}"}],
            "warnings": [],
            "suggestions": ["修改需求或检查输入文件后重新生成。"],
            "summary": {},
        }
        st.session_state.subjective_review_result = {
            "enabled": False,
            "reviews": [],
            "user_notice": "本次未完成审查。",
        }
        progress("error", f"任务失败：{type(exc).__name__}: {exc}")


def _format_elapsed(seconds: int | float) -> str:
    total = max(0, int(seconds))
    minutes, remain = divmod(total, 60)
    if minutes:
        return f"{minutes}分{remain:02d}秒"
    return f"{remain}秒"


def run_generation_with_status(spec: TaskSpec, *, revision: bool = False) -> None:
    """Show deterministic stage progress and a live elapsed-time indicator."""

    label = "修改任务已接收，准备开始……" if revision else "任务已接收，准备开始……"
    status_box = st.status(label, expanded=True)
    progress_bar = st.progress(0, text="准备任务")
    stage_values = {
        "prepare": 5,
        "input": 12,
        "model": 28,
        "build": 58,
        "validate": 76,
        "review": 90,
        "complete": 100,
        "error": 100,
    }

    def show_progress(stage: str, message: str) -> None:
        value = stage_values.get(stage, 20)
        progress_bar.progress(value, text=message)
        status_box.write(message)
        if stage == "complete":
            status_box.update(
                label="修改版已生成并检查完成" if revision else "表格已生成并检查完成",
                state="complete",
            )
        elif stage == "error":
            status_box.update(
                label="修改失败" if revision else "生成失败",
                state="error",
            )

    with st.spinner(
        "正在运行，请勿关闭页面。计时器会持续显示已用时间。",
        show_time=True,
        width="stretch",
    ):
        execute_generation(spec, show_progress)


def _persist_model_settings(updated: ModelSettings) -> None:
    """Save multi-model settings and keep the legacy single-model mirror in sync."""
    save_model_settings(updated)
    st.session_state.model_settings = updated
    builder = get_provider("builder", updated)
    st.session_state.api_settings = builder.to_api_settings() if builder else ApiSettings()
    if builder:
        save_api_settings(st.session_state.api_settings)


def _delete_provider_inline(provider_id: str) -> None:
    current: ModelSettings = st.session_state.model_settings
    providers = [item for item in current.providers if item.id != provider_id]
    roles = {role: pid for role, pid in current.roles.items() if pid != provider_id}
    st.session_state.provider_test_results.pop(provider_id, None)
    if providers:
        for role in ROLE_NAMES:
            roles.setdefault(role, providers[0].id)
    _persist_model_settings(
        ModelSettings(
            providers=providers,
            roles=roles,
            agent_enabled=current.agent_enabled,
            run_python_enabled=current.run_python_enabled,
        )
    )


def _provider_status_badge(provider: ProviderConfig) -> str:
    """Return a UI status that does not confuse config completeness with reachability."""

    if not provider.enabled:
        return "⚪ 已停用"
    if not provider.base_url or not provider.api_key or not provider.model:
        return "🔴 未填完整"
    result = st.session_state.provider_test_results.get(provider.id)
    if not result:
        return "🟡 未测试"
    return "🟢 连接正常" if result.get("success") else "🔴 连接失败"


def _remember_provider_test(provider: ProviderConfig, result) -> None:
    st.session_state.provider_test_results[provider.id] = {
        "success": bool(result.success),
        "message": result.content if result.success else result.error,
        "status_code": result.status_code,
        "latency_ms": result.latency_ms,
        "finish_reason": result.finish_reason,
    }


def render_settings_page() -> None:
    render_header("接口设置", "在这里配置你自己的大模型接口（密钥只存在本机，不会上传）。")
    model_settings: ModelSettings = st.session_state.model_settings
    role_provider = get_provider("builder", model_settings)
    with st.container(border=True):
        if role_provider:
            st.info(
                "当前主用模型："
                f"{role_provider.name} · {role_provider.model} · {mask_api_key(role_provider.api_key)}"
                f" · {_provider_status_badge(role_provider)}"
            )
        else:
            st.info("还没有配置完整且启用的模型。未配置时会使用本地规则生成。")

    with st.container(border=True):
        st.subheader("已配置的模型")
        if not model_settings.providers:
            st.info("还没有添加模型。在下面选好厂商、填入你的密钥即可。")
        else:
            for item in model_settings.providers:
                vendor = PRESET_BY_KEY[detect_provider_key(item.base_url)].name
                roles_here = [
                    ROLE_LABELS[role]
                    for role in ROLE_NAMES
                    if model_settings.roles.get(role) == item.id
                ]
                info_col, test_col, del_col = st.columns([6, 1.3, 1.3])
                info_col.markdown(
                    f"**{item.name}**　`{item.model}`　"
                    + _provider_status_badge(item)
                )
                info_col.caption(
                    f"{vendor} · {item.base_url}"
                    + (f" · 角色：{'、'.join(roles_here)}" if roles_here else " · 暂未指派角色")
                )
                if test_col.button("测试连接", key=f"test_{item.id}", width="stretch"):
                    if not item.configured:
                        st.warning(f"「{item.name}」还没填完整地址、密钥或模型名称。")
                    else:
                        with st.spinner(f"正在测试 {item.name} ……"):
                            result = test_provider(item)
                        _remember_provider_test(item, result)
                        if result.success:
                            st.success(f"{item.name}：{result.content or '连接成功'}")
                        else:
                            st.error(f"{item.name} 连接失败：{result.error}")
                if del_col.button("删除", key=f"del_{item.id}", width="stretch"):
                    st.session_state.pending_delete_id = item.id
                    st.rerun()
                if st.session_state.get("pending_delete_id") == item.id:
                    st.warning(f"确定删除「{item.name}」吗？删除后它的角色会自动改到其它模型。")
                    yes_col, no_col, _ = st.columns([1, 1, 4])
                    if yes_col.button("确定删除", key=f"delyes_{item.id}", type="primary", width="stretch"):
                        _delete_provider_inline(item.id)
                        st.session_state.pending_delete_id = None
                        st.rerun()
                    if no_col.button("取消", key=f"delno_{item.id}", width="stretch"):
                        st.session_state.pending_delete_id = None
                        st.rerun()
                st.divider()

        with st.expander("添加或更新模型", expanded=not model_settings.providers):
            provider_options = ["新建模型", *[item.id for item in model_settings.providers]]
            selected_provider_id = st.selectbox(
                "选择要编辑的模型",
                provider_options,
                key="provider_edit_target",
            )
            selected_provider = next(
                (item for item in model_settings.providers if item.id == selected_provider_id),
                ProviderConfig(),
            )
            is_new_provider = selected_provider_id == "新建模型"
            # 厂商预设（放在表单外，选中后立即把正确的接口地址预填进表单）。
            preset_keys = [item.key for item in PROVIDER_PRESETS]
            default_preset = (
                "custom" if is_new_provider else detect_provider_key(selected_provider.base_url)
            )
            preset_key = st.selectbox(
                "厂商预设（自动填好正确的接口地址，避免填错）",
                preset_keys,
                index=preset_keys.index(default_preset),
                format_func=lambda key: PRESET_BY_KEY[key].name,
                # Key includes the edited model so switching models refreshes the
                # preset to that model's detected vendor (instead of getting stuck).
                key=f"provider_preset_{selected_provider_id}",
            )
            preset = PRESET_BY_KEY[preset_key]
            if preset.note or preset.model_examples:
                hint = "ℹ️ " + preset.note
                if preset.model_examples:
                    hint += "　模型名示例：" + "、".join(preset.model_examples)
                st.caption(hint)
            base_default = preset.base_url if is_new_provider else selected_provider.base_url
            name_default = preset.name if is_new_provider else selected_provider.name
            model_placeholder = (
                preset.model_examples[0] if preset.model_examples else "例如 deepseek-chat"
            )
            with st.form("provider_settings_form"):
                left, right = st.columns(2)
                with left:
                    provider_id = st.text_input(
                        "模型 ID",
                        value="" if is_new_provider else selected_provider.id,
                        placeholder="例如 deepseek-chat",
                    )
                    provider_name = st.text_input(
                        "显示名称",
                        value=name_default,
                    )
                    provider_base = st.text_input(
                        "接口地址（模型）",
                        value=base_default,
                        placeholder="例如：https://api.deepseek.com/v1",
                    )
                with right:
                    provider_model = st.text_input(
                        "模型名称（模型）",
                        value="" if is_new_provider else selected_provider.model,
                        placeholder=model_placeholder,
                    )
                    provider_key = st.text_input(
                        "接口密钥（模型）",
                        value="" if selected_provider_id == "新建模型" else selected_provider.api_key,
                        type="password",
                    )
                    provider_timeout = st.number_input(
                        "等待时间（模型，秒）",
                        min_value=5,
                        max_value=600,
                        value=(
                            120
                            if selected_provider_id == "新建模型"
                            else selected_provider.timeout_seconds
                        ),
                        step=5,
                    )
                    provider_enabled = st.checkbox(
                        "启用这个模型",
                        value=True if selected_provider_id == "新建模型" else selected_provider.enabled,
                    )
                save_provider = st.form_submit_button("保存这个模型", type="primary")
            if save_provider:
                normalized_id = safe_provider_id(provider_id or provider_name or provider_model)
                updated_provider = ProviderConfig(
                    id=normalized_id,
                    name=provider_name or provider_model or "自定义模型",
                    base_url=provider_base,
                    api_key=provider_key,
                    model=provider_model,
                    timeout_seconds=provider_timeout,
                    enabled=provider_enabled,
                )
                providers = [
                    item for item in model_settings.providers if item.id != updated_provider.id
                ]
                providers.append(updated_provider)
                roles = dict(model_settings.roles)
                for role in ROLE_NAMES:
                    roles.setdefault(role, updated_provider.id)
                st.session_state.model_settings = ModelSettings(
                    providers=providers,
                    roles=roles,
                    agent_enabled=model_settings.agent_enabled,
                    run_python_enabled=model_settings.run_python_enabled,
                )
                save_model_settings(st.session_state.model_settings)
                st.session_state.provider_test_results.pop(updated_provider.id, None)
                if updated_provider.configured:
                    st.session_state.api_settings = updated_provider.to_api_settings()
                    save_api_settings(st.session_state.api_settings)
                st.success("模型设置已保存。")
                st.rerun()

        if model_settings.providers:
            with st.form("model_roles_form"):
                st.markdown("**角色指派**")
                choices = [item.id for item in model_settings.providers]
                role_values: dict[str, str] = {}
                for role in ROLE_NAMES:
                    current = model_settings.roles.get(role)
                    default_index = choices.index(current) if current in choices else 0
                    role_values[role] = st.selectbox(
                        ROLE_LABELS.get(role, role),
                        choices,
                        index=default_index,
                        key=f"role_select_{role}",
                    )
                agent_enabled = st.checkbox(
                    "启用自动多步生成",
                    value=model_settings.agent_enabled,
                    help="开启后，复杂需求会自动分步分析并选择本地表格工具生成；明确模板填充任务仍走快路径。",
                )
                run_python_enabled = st.checkbox(
                    "允许安全脚本工具",
                    value=model_settings.run_python_enabled,
                    help="开启后，后续智能体可在任务临时目录中运行受限 Python 脚本。",
                )
                save_roles = st.form_submit_button("保存角色分工", type="primary")
            if save_roles:
                st.session_state.model_settings = ModelSettings(
                    providers=model_settings.providers,
                    roles=role_values,
                    agent_enabled=agent_enabled,
                    run_python_enabled=run_python_enabled,
                )
                save_model_settings(st.session_state.model_settings)
                builder = get_provider("builder", st.session_state.model_settings)
                if builder:
                    st.session_state.api_settings = builder.to_api_settings()
                    save_api_settings(st.session_state.api_settings)
                st.success("角色分工已保存。")
                st.rerun()

    with st.container(border=True):
        st.subheader("我的偏好")
        preferences = list_preferences()
        if preferences:
            st.json(preferences, expanded=False)
            if st.button("清空已学习偏好", width="stretch"):
                clear_preferences()
                st.success("已清空偏好。")
                st.rerun()
        else:
            st.info("还没有学习到偏好。系统只保存低风险偏好，例如常用 sheet 语言、是否偏好图表或模板样式。")

    with st.container(border=True):
        st.subheader("清除设置")
        with st.form("clear_api_settings_form"):
            confirm_clear = st.checkbox("我确认清除本机保存的接口设置")
            clear_clicked = st.form_submit_button("清除设置")
        if clear_clicked:
            if not confirm_clear:
                st.warning("请先勾选确认。")
            else:
                delete_api_settings()
                delete_model_settings()
                st.session_state.api_settings = ApiSettings()
                st.session_state.model_settings = ModelSettings()
                st.session_state.api_edit_mode = False
                st.success("接口设置已清除。")


def open_history_task(item: dict[str, Any]) -> None:
    paths = existing_task_paths(
        str(item.get("task_id", "")),
        item.get("output_file"),
    )
    spec = load_task_spec(paths.task_spec_file)
    spec.options["task_id"] = paths.task_id
    validation_payload = (
        json.loads(paths.validation_report.read_text(encoding="utf-8"))
        if paths.validation_report.exists()
        else {
            "status": "warn",
            "errors": [],
            "warnings": [],
            "suggestions": ["历史任务缺少检查报告，可重新生成一个修改版。"],
            "summary": {},
        }
    )
    subjective = (
        json.loads(paths.subjective_review_report.read_text(encoding="utf-8"))
        if paths.subjective_review_report.exists()
        else {
            "enabled": False,
            "reviews": [],
            "user_notice": "历史任务没有审查建议。",
        }
    )
    st.session_state.user_prompt = spec.user_goal
    st.session_state.uploaded_files = []
    st.session_state.task_spec = spec
    st.session_state.task_paths = paths
    st.session_state.generation_result = {
        "success": paths.output_file.exists(),
        "message": "已打开历史表格。",
        "error": None,
        "notices": [],
        "mode": "history",
    }
    st.session_state.validation_result = {
        "status": validation_payload.get("status", "warn"),
        "issues": validation_payload.get("errors", []),
        "warnings": validation_payload.get("warnings", []),
        "suggestions": validation_payload.get("suggestions", []),
        "summary": validation_payload.get("summary", {}),
        "report_file": str(paths.validation_report),
    }
    st.session_state.subjective_review_result = subjective
    st.session_state.clarification_done = True
    st.session_state.versions = [
        {
            "task_id": paths.task_id,
            "output_name": spec.output_name,
            "output_file": str(paths.output_file),
            "status": validation_payload.get("status", "warn"),
            "revision_index": int(spec.options.get("revision_index", 1)),
        }
    ]
    st.session_state.active_page = "workbench"


def render_history_page() -> None:
    render_header(
        "最近文件",
        "重新打开以前生成的表格，继续预览、查看报告或生成修改版。",
    )
    memory_items = list_task_history(30)
    manifest_items = recent_tasks(30)
    merged: dict[str, dict[str, Any]] = {}
    for item in [*manifest_items, *memory_items]:
        task_id = str(item.get("task_id") or "")
        if task_id and task_id not in merged:
            merged[task_id] = item
    items = list(merged.values())[:30]
    if not items:
        with st.container(border=True):
            st.info("还没有生成记录。")
        return
    for item in items:
        output_name = (
            Path(str(item.get("output_file"))).name
            if item.get("output_file")
            else TYPE_LABELS.get(item.get("task_type"), "表格")
        )
        with st.container(border=True):
            text_col, status_col, action_col = st.columns([5, 1.2, 1.2])
            text_col.markdown(f"**{output_name}**")
            created = str(item.get("created_at", "")).replace("T", " ")[:19]
            text_col.caption(created)
            status_col.write(status_label(str(item.get("status", ""))))
            if action_col.button(
                "打开",
                key=f"open_history_{item.get('task_id')}",
                width="stretch",
            ):
                try:
                    open_history_task(item)
                    st.rerun()
                except Exception as exc:
                    st.error(f"无法打开这个历史任务：{exc}")


def plan_card(label: str, value: str) -> None:
    st.markdown(
        f"""
        <div class="plan-card">
          <div class="plan-label">{html.escape(label)}</div>
          <div class="plan-value">{html.escape(value)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_plan(spec: TaskSpec) -> None:
    plan = spec.options.get("content_plan", {})
    title = str(plan.get("title") or TYPE_LABELS.get(spec.task_type, "自定义表格"))
    columns = [str(item.get("name")) for item in plan.get("columns", []) if item.get("name")]
    formulas = [
        str(item.get("target"))
        for item in plan.get("formula_rules", [])
        if item.get("target")
    ]
    row_count = int(plan.get("expected_data_rows") or 0)
    cards = st.columns(3)
    with cards[0]:
        plan_card("表格内容", title)
    with cards[1]:
        plan_card("文件名称", spec.output_name)
    with cards[2]:
        source = (
            f"{len(spec.input_files)} 个上传文件"
            if spec.input_files
            else f"需求中识别到 {row_count} 条数据"
            if row_count
            else "创建可填写表格"
        )
        plan_card("数据来源", source)

    if columns:
        st.markdown("**将生成的列**")
        st.write("、".join(columns))
    else:
        st.markdown("**表格类型**")
        st.write(TYPE_LABELS.get(spec.task_type, "通用表格"))
    if formulas:
        st.markdown("**自动计算**")
        st.write("、".join(formulas))
    summary_rules = plan.get("summary_rules", [])
    if summary_rules:
        st.markdown("**汇总内容**")
        st.write("、".join(str(item.get("label")) for item in summary_rules))
    if spec.assumptions:
        st.markdown(
            '<div class="soft-note">'
            + "<br>".join(html.escape(item) for item in spec.assumptions)
            + "</div>",
            unsafe_allow_html=True,
        )
    if spec.template_files:
        mode_labels = {
            "reference": "仅参考样式与布局",
            "flexible": "灵活套用，需求优先",
            "strict": "严格遵守模板结构",
        }
        mode = str(spec.options.get("template_mode") or "reference")
        data_policy = (
            "保留模板中的数据"
            if spec.options.get("use_template_data")
            else "忽略模板中的示例数据"
        )
        st.markdown("**模板使用方式**")
        st.write(
            f"{mode_labels.get(mode, mode)}；{data_policy}；"
            f"模板：{'、'.join(Path(item).name for item in spec.template_files)}"
        )


def render_requirement_check(spec: TaskSpec) -> None:
    plan = spec.options.get("content_plan", {})
    columns = [item for item in plan.get("columns", []) if item.get("name")]
    formulas = [item for item in plan.get("formula_rules", []) if item.get("target")]
    inline_tables = [
        item for item in plan.get("inline_tables", []) if isinstance(item, dict)
    ]
    rows = int(plan.get("expected_data_rows") or 0)
    checks = [
        ("用途", TYPE_LABELS.get(spec.task_type, "通用表格")),
        (
            "数据",
            f"已识别 {len(inline_tables)} 个内嵌数据表，主表 {rows} 行"
            if inline_tables
            else f"已识别 {rows} 条文字数据"
            if rows
            else f"将使用 {len(spec.input_files)} 个输入文件"
            if spec.input_files
            else "未提供数据，将创建可填写内容",
        ),
        (
            "结构",
            f"已明确 {len(columns)} 列"
            if columns
            else "将采用该类型的标准结构，可在下方继续补充",
        ),
        (
            "计算",
            f"已识别 {len(formulas)} 项自动计算"
            if formulas
            else "将依据最终需求和表格类型设置必要公式",
        ),
    ]
    st.markdown("**需求检查结果**")
    check_columns = st.columns(4)
    for container, (label, value) in zip(check_columns, checks):
        with container:
            plan_card(label, value)
    model_summary = str(spec.options.get("model_goal_summary") or "").strip()
    if model_summary:
        st.caption(f"整理后的目标：{model_summary}")


def render_clarification(spec: TaskSpec) -> None:
    nonce = st.session_state.analysis_nonce
    gaps = {
        str(item.get("question")): item
        for item in spec.options.get("requirement_gaps", [])
        if isinstance(item, dict)
    }
    with st.container(border=True):
        st.subheader("完善生成要求")
        st.caption("下面这些内容会直接影响表格结构和计算结果。补充后，系统会重新整理生成方案。")
        render_requirement_check(spec)
        for index, question in enumerate(st.session_state.clarifying_questions):
            gap = gaps.get(question, {})
            if gap.get("title"):
                st.markdown(f"**{gap['title']}**")
            st.text_area(
                question,
                placeholder=str(gap.get("example") or "请尽量写清楚具体字段、计算口径或样式要求。"),
                key=f"clarification_{nonce}_{index}",
                height=88,
            )
        type_index = SUPPORTED_TYPES.index(spec.task_type)
        st.selectbox(
            "确认表格类型",
            SUPPORTED_TYPES,
            index=type_index,
            format_func=lambda value: TYPE_LABELS.get(value, "通用表格"),
            key=f"clarify_type_{nonce}",
        )
        st.text_area(
            "其他必须满足的要求",
            placeholder="例如：工作表名称、是否保留原模板、打印方向、冻结位置、颜色规范、不能出现的内容。",
            key=f"clarify_detail_{nonce}",
        )
        st.radio(
            "没有原始数据时",
            ["template", "upload"],
            format_func=lambda value: "先创建可填写表格"
            if value == "template"
            else "稍后补充数据",
            horizontal=True,
            key=f"clarify_data_mode_{nonce}",
        )
        use_defaults = st.checkbox(
            "未填写的可选项按系统建议处理",
            value=False,
            key=f"use_defaults_{nonce}",
        )
        if st.button("补充完成，重新整理方案", type="primary", width="stretch"):
            required_questions = {
                item.get("question")
                for item in gaps.values()
                if item.get("required", True)
            }
            unanswered = [
                question
                for index, question in enumerate(st.session_state.clarifying_questions)
                if question in required_questions
                and not str(
                    st.session_state.get(f"clarification_{nonce}_{index}", "")
                ).strip()
            ]
            if unanswered and not use_defaults:
                st.warning("仍有会影响结果的必填信息未补充；请填写，或勾选按系统建议处理。")
                return
            apply_clarification()
            st.rerun()


def render_confirmation(spec: TaskSpec) -> None:
    nonce = st.session_state.analysis_nonce
    with st.container(border=True):
        st.subheader("确认生成内容")
        st.success("需求检查已完成。请核对下面的结构和选项，确认后再生成。")
        render_requirement_check(spec)
        spec.output_name = st.text_input(
            "文件名称",
            value=spec.output_name,
            key=f"output_name_{nonce}",
        )
        render_plan(spec)
        with st.expander("图表与其他选项", expanded=spec.include_charts):
            col_a, col_b = st.columns(2)
            spec.include_charts = col_a.checkbox(
                "生成图表",
                value=spec.include_charts,
                key=f"charts_{nonce}",
                help="勾选后可在下面挑选想要的图表样式。",
            )
            spec.include_summary = col_b.checkbox(
                "生成独立汇总页",
                value=spec.include_summary,
                key=f"summary_{nonce}",
            )
            if spec.include_charts:
                default_types = spec.options.get("chart_types") or [
                    chart_type_from_text(spec.user_goal, default="column")
                ]
                chosen = st.multiselect(
                    "想要的图表样式（可多选，不确定就保留默认）",
                    options=list(CHART_TYPE_LABELS.keys()),
                    default=[item for item in default_types if item in CHART_TYPE_LABELS]
                    or ["column"],
                    format_func=lambda value: CHART_TYPE_LABELS.get(value, value),
                    key=f"chart_types_{nonce}",
                )
                spec.options["chart_types"] = chosen or ["column"]
            else:
                spec.options["chart_types"] = []
        st.session_state.task_spec = spec
        confirm, modify = st.columns(2)
        if confirm.button("确认并生成", type="primary", width="stretch"):
            run_generation_with_status(spec)
            st.rerun()
        if modify.button("重新描述需求", width="stretch"):
            st.session_state.task_spec = None
            st.session_state.classification_result = None
            st.session_state.clarifying_questions = []
            st.session_state.clarification_done = False
            reset_generated_result()
            st.rerun()


def status_label(status: str) -> str:
    return {
        "pass": "通过",
        "warn": "有提醒",
        "fail": "未通过",
        "error": "执行异常",
        "success": "成功",
    }.get(str(status), "未知")


def render_workbook_preview(task_paths: TaskPaths) -> None:
    if not task_paths.output_file.exists():
        st.error("没有可预览的表格文件。")
        return
    try:
        preview = workbook_preview(task_paths.output_file)
    except Exception as exc:
        st.error(f"暂时无法预览：{exc}")
        return
    sheet_names = preview["sheet_names"]
    selected = st.selectbox("工作表", sheet_names, key=f"preview_{task_paths.task_id}")
    sheet = next(item for item in preview["sheets"] if item["name"] == selected)
    if sheet.get("title") and str(sheet["title"]).strip() != selected:
        st.markdown(f"**{sheet['title']}**")
    st.caption(
        f"共 {sheet['max_row']} 行、{sheet['max_column']} 列；"
        f"包含 {sheet['formula_count']} 个自动计算单元格。"
    )
    frame = sheet_preview_dataframe(sheet)
    st.dataframe(frame, width="stretch", hide_index=True, height=520)
    chart_preview = workbook_chart_preview(preview)
    if chart_preview is not None:
        chart_frame, chart_kind = chart_preview
        st.markdown("**图表预览**")
        if chart_kind == "line":
            st.line_chart(chart_frame)
        else:
            st.bar_chart(chart_frame)


def render_plan_preview(spec: TaskSpec) -> None:
    render_plan(spec)
    plan = spec.options.get("content_plan", {})
    st.markdown("**工作表安排**")
    if plan.get("layout") == "single_sheet":
        st.write("所有数据、计算和汇总放在同一个工作表中。")
    elif spec.input_files:
        st.write("说明、原始数据、清洗数据、汇总和清洗报告。")
    else:
        st.write("说明、数据和汇总。")


def render_validation(validation: dict[str, Any]) -> None:
    status = str(validation.get("status", "error"))
    if status == "pass":
        st.success("自动检查通过。")
    elif status == "warn":
        st.warning("表格可以使用，但有需要留意的内容。")
    else:
        st.error("自动检查未通过，建议修改后重新生成。")
    sections = [
        ("发现的问题", validation.get("issues", [])),
        ("需要留意", validation.get("warnings", [])),
        ("改进建议", validation.get("suggestions", [])),
    ]
    for title, items in sections:
        if not items:
            continue
        st.markdown(f"**{title}**")
        for item in items:
            message = item.get("message", item) if isinstance(item, dict) else item
            st.write(f"• {message}")


def review_items(subjective: dict[str, Any]) -> tuple[list[str], list[str]]:
    concerns: list[str] = []
    suggestions: list[str] = []
    for review in subjective.get("reviews", []):
        concerns.extend(str(item) for item in review.get("concerns", []))
        suggestions.extend(str(item) for item in review.get("suggestions", []))
    return concerns, suggestions


def render_review(
    subjective: dict[str, Any],
    validation: dict[str, Any],
) -> None:
    reviews = subjective.get("reviews", [])
    if not reviews:
        st.info("本次没有启用建议审查。你仍可以在“继续修改”中直接提出修改要求。")
        return
    labels = {
        "requirement_review": "需求一致性",
        "excel_usability_review": "Excel 可用性",
    }
    st.markdown("**审查项**")
    for review in reviews:
        name = labels.get(str(review.get("reviewer")), str(review.get("reviewer", "审查")))
        st.write(f"• {name}：{status_label(str(review.get('status', 'warn')))}")
    concerns, suggestions = review_items(subjective)
    if concerns:
        st.markdown("**需要调整**")
        for item in concerns:
            st.write(f"• {item}")
    if suggestions:
        st.markdown("**建议做法**")
        for item in suggestions:
            st.write(f"• {item}")
    revision_prompt = (
        subjective.get("revision_prompt")
        or combined_revision_prompt(subjective, validation)
    )
    if revision_prompt and st.button(
        "把这些建议带入下一次修改",
        type="primary",
        width="stretch",
    ):
        st.session_state.revision_request = revision_prompt
        st.rerun()


def render_revision(spec: TaskSpec) -> None:
    st.markdown("**继续修改当前表格**")
    st.caption("原文件会保留，每次修改都会生成一个新版本。")
    request = st.text_area(
        "修改要求",
        key="revision_request",
        placeholder="例如：删除独立汇总页，把周平均放在主表底部；标题改为……",
        height=150,
    )
    default_name = (
        st.session_state.revision_output_name
        or f"{Path(spec.output_name).stem}_修改版.xlsx"
    )
    output_name = st.text_input(
        "修改版文件名称",
        value=default_name,
        key=f"revision_output_name_{spec.options.get('task_id', 'current')}",
    )
    if st.button("生成修改版", type="primary", width="stretch"):
        if not request.strip():
            st.warning("请先在上面填写要修改的内容，再点“生成修改版”。")
            return
        try:
            revised = build_revision_task_spec(
                spec,
                request,
                st.session_state.api_settings,
            )
            revised.output_name = output_name
            run_generation_with_status(revised, revision=True)
            # Note: do NOT assign to st.session_state["revision_request"] here —
            # that key belongs to the text_area widget above and Streamlit forbids
            # writing it after the widget is created (it raised silently before,
            # so the page never refreshed and the new version looked "stuck").
            st.rerun()
        except Exception as exc:
            st.error(f"生成修改版时出错：{exc}")
    if len(st.session_state.versions) > 1:
        st.markdown("**本次会话中的版本**")
        for index, item in enumerate(st.session_state.versions, start=1):
            st.write(
                f"{index}. {item['output_name']} · {status_label(item['status'])}"
            )


def download_file(
    label: str,
    path: Path,
    mime: str,
    key: str,
    filename: str | None = None,
) -> None:
    if not path.exists():
        return
    st.download_button(
        label=label,
        data=path.read_bytes(),
        file_name=filename or path.name,
        mime=mime,
        key=key,
        width="stretch",
    )


def render_downloads(task_paths: TaskPaths, spec: TaskSpec) -> None:
    st.caption("可直接预览，也可以保存到电脑继续编辑。")
    col_a, col_b = st.columns(2)
    with col_a:
        download_file(
            "下载表格",
            task_paths.output_file,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            f"download_excel_{task_paths.task_id}",
            filename=spec.output_name,
        )
        download_file(
            "下载检查报告",
            task_paths.validation_report,
            "application/json",
            f"download_validation_{task_paths.task_id}",
            filename=f"{Path(spec.output_name).stem}_检查报告.json",
        )
    with col_b:
        download_file(
            "下载生成方案",
            task_paths.task_spec_file,
            "application/json",
            f"download_spec_{task_paths.task_id}",
            filename=f"{Path(spec.output_name).stem}_生成方案.json",
        )
        download_file(
            "下载审查建议",
            task_paths.subjective_review_report,
            "application/json",
            f"download_review_{task_paths.task_id}",
            filename=f"{Path(spec.output_name).stem}_审查建议.json",
        )


def render_result() -> None:
    task_paths: TaskPaths = st.session_state.task_paths
    spec: TaskSpec = st.session_state.task_spec
    generation = st.session_state.generation_result or {}
    validation = st.session_state.validation_result or {}
    subjective = st.session_state.subjective_review_result or {}
    with st.container(border=True):
        if generation.get("success"):
            st.success(f"已生成：{spec.output_name}")
        else:
            st.error(generation.get("error") or "生成失败。")
        if generation.get("elapsed_seconds") is not None:
            st.caption(
                f"本次运行用时：{_format_elapsed(generation['elapsed_seconds'])}"
            )
        for notice in generation.get("notices", []):
            st.info(str(notice))

    tabs = st.tabs(
        ["表格预览", "生成方案", "检查报告", "审查建议", "继续修改", "下载"]
    )
    with tabs[0]:
        render_workbook_preview(task_paths)
    with tabs[1]:
        render_plan_preview(spec)
    with tabs[2]:
        render_validation(validation)
    with tabs[3]:
        render_review(subjective, validation)
    with tabs[4]:
        render_revision(spec)
    with tabs[5]:
        render_downloads(task_paths, spec)


def render_workbench() -> None:
    render_header(
        "本地表格助手",
        "用大白话描述你要的表格，确认后一键生成；生成后可在线预览、检查、调整图表并继续修改。",
    )
    spec_state = st.session_state.task_spec
    if st.session_state.task_paths is not None:
        active_step = 4
    elif spec_state and not st.session_state.clarification_done:
        active_step = 2
    elif spec_state:
        active_step = 3
    else:
        active_step = 1
    render_steps(active_step)

    if st.session_state.task_paths is not None:
        action_col, spacer = st.columns([1.4, 6])
        if action_col.button("＋ 制作另一张表格", width="stretch"):
            start_new_task()
            st.rerun()
        render_result()
        return

    with st.container(border=True):
        st.subheader("制作新表格")
        prompt = st.text_area(
            "需求描述",
            value=st.session_state.user_prompt,
            placeholder=(
                "例如：根据下面的数据制作天气表，列为日期、城市、最高气温、"
                "最低气温和出行建议，日均气温用公式计算……"
            ),
            height=160,
            label_visibility="collapsed",
        )
        st.markdown("**数据来源**")
        st.caption("数据用于生成内容；支持旧版 XLS、Excel、CSV、TSV 和文本文件。")
        data_uploads = st.file_uploader(
            "上传数据文件",
            type=["csv", "tsv", "txt", "xls", "xlsx", "xlsm"],
            accept_multiple_files=True,
            help="这里上传真实业务数据，不要把只用于参考格式的模板放在这里。",
            key="data_file_uploader",
        )
        pasted_data_text = st.text_area(
            "或粘贴文本数据",
            value=st.session_state.pasted_data_text,
            placeholder=(
                "可直接粘贴名单、制表符数据、CSV 内容或逐行文本。"
                "请尽量保留表头，并在需求中说明各列含义。"
            ),
            height=110,
            key="pasted_data_text_input",
        )

        st.markdown("**模板文件（可选）**")
        st.caption("模板只决定格式和结构；默认不把模板里的示例数据带入结果。")
        template_uploads = st.file_uploader(
            "上传模板文件",
            type=["xls", "xlsx", "xlsm"],
            accept_multiple_files=False,
            help="这里上传参考格式、固定导入格式或必须遵守的工作簿模板。",
            key="template_file_uploader",
        )
        template_mode_labels = {
            "reference": "仅作参考：参考配色、字体和大致布局，需求可以重新设计结构",
            "flexible": "灵活套用：尽量保留模板形式，冲突时以本次需求为准",
            "strict": "严格遵守：保持字段顺序和工作表结构，冲突时停止并提示",
        }
        template_mode = st.radio(
            "模板约束方式",
            list(template_mode_labels),
            format_func=lambda value: template_mode_labels[value],
            index=1,
            disabled=template_uploads is None,
            key="template_mode_selector",
        )
        use_template_data = st.checkbox(
            "同时使用模板中已有的数据",
            value=False,
            disabled=template_uploads is None,
            help="通常不要勾选。模板中的内容一般只是填写示例，默认会被忽略。",
            key="use_template_data_checkbox",
        )
        if st.button("检查并完善需求", type="primary", width="stretch"):
            with st.spinner("正在逐项检查数据、字段、计算、汇总和图表要求……"):
                analyze_request(
                    prompt,
                    list(data_uploads or []),
                    [template_uploads] if template_uploads else [],
                    pasted_data_text,
                    template_mode,
                    use_template_data,
                )
            st.rerun()

    spec = st.session_state.task_spec
    if spec and not st.session_state.clarification_done:
        render_clarification(spec)
    elif spec and st.session_state.task_paths is None:
        render_confirmation(spec)



initialize_state()
inject_styles()
show_onboarding_if_first_use()
render_navigation()
if st.session_state.active_page == "settings":
    render_settings_page()
elif st.session_state.active_page == "history":
    render_history_page()
else:
    render_workbench()
