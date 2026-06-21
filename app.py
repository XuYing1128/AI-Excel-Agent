"""AI-Excel-Agent 本地中文网页。"""

from __future__ import annotations

import html
import json
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
from excel_agent.intent_classifier import SUPPORTED_TYPES, classify_intent
from excel_agent.manifest import (
    append_manifest_record,
    build_manifest_record,
    recent_tasks,
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
    "api_message": "",
    "api_message_kind": "info",
    "api_edit_mode": False,
    "versions": [],
    "revision_request": "",
    "revision_output_name": "",
}


def initialize_state() -> None:
    for key, default in STATE_DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = deepcopy(default)
    if st.session_state.api_settings is None:
        st.session_state.api_settings = load_api_settings()


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        :root {
            --brand: #2563eb;
            --brand-dark: #1d4ed8;
            --ink: #172033;
            --muted: #64748b;
            --line: #e3e8f1;
            --canvas: #f5f7fb;
        }
        html, body, [class*="css"] {
            font-family: "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", sans-serif;
        }
        .stApp { background: var(--canvas); color: var(--ink); }
        .block-container {
            max-width: 1180px;
            padding-top: 1.4rem;
            padding-bottom: 4rem;
        }
        header[data-testid="stHeader"],
        [data-testid="stToolbar"],
        [data-testid="stDecoration"],
        [data-testid="stStatusWidget"],
        #MainMenu, footer, [data-testid="stSidebar"] {
            display: none !important;
        }
        .page-head {
            margin: 12px 0 24px;
            padding: 26px 30px;
            background: #fff;
            border: 1px solid var(--line);
            border-radius: 18px;
        }
        .page-title {
            font-size: 30px;
            line-height: 1.3;
            font-weight: 800;
            color: #10203a;
            margin-bottom: 7px;
        }
        .page-subtitle {
            color: var(--muted);
            font-size: 15px;
            line-height: 1.75;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background: #fff;
            border: 1px solid var(--line);
            border-radius: 16px;
            box-shadow: 0 5px 18px rgba(30, 50, 90, .04);
        }
        h1, h2, h3 { color: var(--ink); letter-spacing: -.02em; }
        h3 { font-size: 20px !important; }
        .stButton > button, .stDownloadButton > button {
            border-radius: 9px;
            min-height: 42px;
            font-weight: 700;
        }
        .stButton > button[kind="primary"] {
            background: var(--brand);
            border-color: var(--brand);
        }
        .stButton > button[kind="primary"]:hover {
            background: var(--brand-dark);
            border-color: var(--brand-dark);
        }
        div[data-testid="stTextArea"] textarea,
        div[data-testid="stTextInput"] input { border-radius: 9px; }
        [data-testid="stFileUploaderDropzone"] {
            background: #f8faff;
            border: 1px dashed #b8c8e5;
            border-radius: 11px;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] { display: none; }
        [data-testid="stFileUploaderDropzone"] button > * { display: none !important; }
        [data-testid="stFileUploaderDropzone"] button::after {
            content: "选择本地文件";
            font-size: 14px;
        }
        [data-testid="stFileUploaderDropzone"] button { font-size: 0; }
        div[data-testid="stAlert"] { border-radius: 11px; }
        .plan-card {
            border: 1px solid var(--line);
            border-radius: 12px;
            padding: 14px 16px;
            background: #fbfcff;
            min-height: 102px;
        }
        .plan-label { color: #718096; font-size: 12px; margin-bottom: 7px; }
        .plan-value {
            color: var(--ink);
            font-size: 15px;
            line-height: 1.55;
            font-weight: 650;
        }
        .soft-note {
            padding: 12px 14px;
            background: #f8fafc;
            border-left: 3px solid #94a3b8;
            border-radius: 8px;
            color: #475569;
            line-height: 1.7;
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
    def progress(stage: str, message: str) -> None:
        if progress_callback:
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
        st.session_state.task_spec = spec
        st.session_state.task_paths = task_paths
        st.session_state.generation_result = generation.to_dict()
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


def render_settings_page() -> None:
    render_header("接口设置", "接口配置单独保存在这里，只有点击保存后才会生效。")
    settings: ApiSettings = st.session_state.api_settings
    with st.container(border=True):
        st.subheader("当前状态")
        if settings.configured:
            st.success(
                f"已启用：{settings.provider_name} / {settings.model} / {mask_api_key(settings.api_key)}"
            )
        else:
            st.info("当前使用本地规则。需要时可以配置兼容对话补全格式的接口。")
        if st.button("编辑接口设置", width="content"):
            st.session_state.api_edit_mode = not st.session_state.api_edit_mode
            st.rerun()

    if st.session_state.api_edit_mode:
        with st.container(border=True):
            st.subheader("编辑设置")
            with st.form("api_settings_form"):
                col_a, col_b = st.columns(2)
                with col_a:
                    provider = st.text_input("接口名称", value=settings.provider_name)
                    base_url = st.text_input(
                        "接口地址",
                        value=settings.base_url,
                        placeholder="例如：https://你的接口地址/v1",
                    )
                    model = st.text_input("模型名称", value=settings.model)
                with col_b:
                    api_key = st.text_input(
                        "接口密钥",
                        value=settings.api_key,
                        type="password",
                    )
                    timeout = st.number_input(
                        "等待时间（秒）",
                        min_value=5,
                        max_value=180,
                        value=settings.timeout_seconds,
                        step=5,
                    )
                    enabled = st.checkbox("启用这个接口", value=settings.enabled)
                use_intent = st.checkbox(
                    "用于理解需求",
                    value=settings.use_for_intent,
                )
                use_review = st.checkbox(
                    "用于审查生成结果",
                    value=settings.use_for_review,
                )
                use_generation = st.checkbox(
                    "由大模型调用本地工具生成表格",
                    value=settings.use_for_generation,
                    help="启用后，大模型会制定完整工作簿方案并调用本地 Excel 工具；失败时不会静默退回无关模板。",
                )
                save_clicked = st.form_submit_button(
                    "保存设置",
                    type="primary",
                    width="stretch",
                )
            if save_clicked:
                updated = ApiSettings(
                    enabled=enabled,
                    base_url=base_url,
                    api_key=api_key,
                    model=model,
                    provider_name=provider,
                    timeout_seconds=timeout,
                    use_for_intent=use_intent,
                    use_for_review=use_review,
                    use_for_generation=use_generation,
                )
                save_api_settings(updated)
                st.session_state.api_settings = updated
                st.session_state.api_message = "设置已保存。"
                st.session_state.api_message_kind = "success"
                st.rerun()

            test_col, cancel_col = st.columns(2)
            if test_col.button("测试连接", width="stretch"):
                candidate = ApiSettings(
                    enabled=enabled,
                    base_url=base_url,
                    api_key=api_key,
                    model=model,
                    provider_name=provider,
                    timeout_seconds=timeout,
                    use_for_intent=use_intent,
                    use_for_review=use_review,
                    use_for_generation=use_generation,
                )
                if not candidate.configured:
                    st.warning("请先填写地址、密钥和模型名称，并启用接口。")
                else:
                    with st.spinner("正在测试连接……"):
                        result = test_api_connection(candidate)
                    if result.success:
                        st.success("连接成功。")
                    else:
                        st.error(result.error or "连接失败。")
            if cancel_col.button("取消编辑", width="stretch"):
                st.session_state.api_edit_mode = False
                st.rerun()

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
                st.session_state.api_settings = ApiSettings()
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
    items = recent_tasks(30)
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
    rows = int(plan.get("expected_data_rows") or 0)
    checks = [
        ("用途", TYPE_LABELS.get(spec.task_type, "通用表格")),
        (
            "数据",
            f"已识别 {rows} 条文字数据"
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
        with st.expander("可选内容"):
            col_a, col_b = st.columns(2)
            spec.include_charts = col_a.checkbox(
                "生成图表",
                value=spec.include_charts,
                key=f"charts_{nonce}",
            )
            spec.include_summary = col_b.checkbox(
                "生成独立汇总页",
                value=spec.include_summary,
                key=f"summary_{nonce}",
            )
        st.session_state.task_spec = spec
        confirm, modify = st.columns(2)
        if confirm.button("确认并生成", type="primary", width="stretch"):
            status_box = st.status("任务已接收，准备开始……", expanded=True)

            def show_progress(stage: str, message: str) -> None:
                status_box.write(message)
                if stage == "complete":
                    status_box.update(label="表格已生成并检查完成", state="complete")
                elif stage == "error":
                    status_box.update(label="生成失败", state="error")

            execute_generation(spec, show_progress)
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
        try:
            revised = build_revision_task_spec(
                spec,
                request,
                st.session_state.api_settings,
            )
            revised.output_name = output_name
            status_box = st.status("修改任务已接收，准备开始……", expanded=True)

            def show_revision_progress(stage: str, message: str) -> None:
                status_box.write(message)
                if stage == "complete":
                    status_box.update(label="修改版已生成并检查完成", state="complete")
                elif stage == "error":
                    status_box.update(label="修改失败", state="error")

            execute_generation(revised, show_revision_progress)
            st.session_state.revision_request = ""
            st.session_state.revision_output_name = ""
            st.rerun()
        except Exception as exc:
            st.error(str(exc))
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
        "描述你需要的内容，确认后生成表格；生成后可直接预览、检查并继续修改。",
    )
    if st.session_state.task_paths is not None:
        action_col, spacer = st.columns([1.4, 6])
        if action_col.button("制作另一张表格", width="stretch"):
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
render_navigation()
if st.session_state.active_page == "settings":
    render_settings_page()
elif st.session_state.active_page == "history":
    render_history_page()
else:
    render_workbench()
