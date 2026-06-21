"""AI-Excel-Agent 本地中文网页。"""

from __future__ import annotations

import html
import json
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
from excel_agent.manifest import append_manifest_record, build_manifest_record, recent_tasks
from excel_agent.services.api_task_planner import enhance_task_spec_draft
from excel_agent.services.custom_api_service import test_api_connection
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


st.set_page_config(
    page_title="本地表格助手",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
    menu_items={},
)


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
    "analysis_message": "",
    "analysis_used_api": False,
    "api_settings": None,
    "api_message": "",
    "api_message_kind": "info",
    "api_nonce": 0,
}


def initialize_state() -> None:
    for key, default in STATE_DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = default
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
            --line: #e4e9f2;
            --panel: #ffffff;
            --canvas: #f5f7fb;
        }
        html, body, [class*="css"] {
            font-family: "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", sans-serif;
        }
        .stApp {
            background: var(--canvas);
            color: var(--ink);
        }
        .block-container {
            max-width: 1120px;
            padding-top: 2rem;
            padding-bottom: 4rem;
        }
        header[data-testid="stHeader"],
        [data-testid="stToolbar"],
        [data-testid="stDecoration"],
        [data-testid="stStatusWidget"],
        #MainMenu,
        footer {
            display: none !important;
        }
        [data-testid="stSidebar"] {
            display: none !important;
        }
        .app-hero {
            background: linear-gradient(135deg, #ffffff 0%, #edf4ff 100%);
            border: 1px solid #dbe7fb;
            border-radius: 22px;
            padding: 30px 34px;
            margin-bottom: 22px;
            box-shadow: 0 12px 34px rgba(36, 83, 160, 0.08);
        }
        .app-kicker {
            display: inline-block;
            color: var(--brand);
            background: #eaf2ff;
            border-radius: 999px;
            padding: 6px 12px;
            font-size: 13px;
            font-weight: 700;
            letter-spacing: .04em;
        }
        .app-title {
            margin: 14px 0 8px;
            font-size: 34px;
            line-height: 1.25;
            font-weight: 800;
            color: #10203a;
        }
        .app-subtitle {
            color: var(--muted);
            font-size: 15px;
            line-height: 1.8;
        }
        .status-row {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-top: 18px;
        }
        .status-pill {
            display: inline-flex;
            align-items: center;
            gap: 7px;
            border-radius: 999px;
            padding: 7px 12px;
            background: #ffffff;
            border: 1px solid var(--line);
            color: #475569;
            font-size: 13px;
        }
        .status-dot {
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: #22c55e;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background: var(--panel);
            border: 1px solid var(--line);
            border-radius: 18px;
            box-shadow: 0 8px 24px rgba(30, 50, 90, 0.05);
        }
        div[data-testid="stVerticalBlockBorderWrapper"] > div {
            padding: 4px 5px;
        }
        h1, h2, h3 {
            color: var(--ink);
            letter-spacing: -.02em;
        }
        h3 {
            font-size: 20px !important;
        }
        .stButton > button,
        .stDownloadButton > button {
            border-radius: 10px;
            min-height: 42px;
            font-weight: 700;
            transition: all .15s ease;
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
        div[data-testid="stTextInput"] input {
            border-radius: 10px;
        }
        [data-testid="stFileUploaderDropzone"] {
            background: #f8faff;
            border: 1px dashed #b8c8e5;
            border-radius: 12px;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] {
            display: none;
        }
        [data-testid="stFileUploaderDropzone"] button > * {
            display: none !important;
        }
        [data-testid="stFileUploaderDropzone"] button::after {
            content: "选择本地文件";
            font-size: 14px;
        }
        [data-testid="stFileUploaderDropzone"] button {
            font-size: 0;
        }
        div[data-testid="stAlert"] {
            border-radius: 12px;
        }
        .spec-card {
            border: 1px solid var(--line);
            border-radius: 14px;
            padding: 16px 18px;
            background: #fbfcff;
            min-height: 116px;
        }
        .spec-label {
            color: #718096;
            font-size: 12px;
            margin-bottom: 7px;
        }
        .spec-value {
            color: var(--ink);
            font-size: 16px;
            line-height: 1.55;
            font-weight: 650;
        }
        .result-path {
            border-radius: 10px;
            background: #f1f5f9;
            color: #475569;
            padding: 11px 13px;
            font-size: 13px;
            word-break: break-all;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def reset_after_analysis() -> None:
    st.session_state.confirmed = False
    st.session_state.task_paths = None
    st.session_state.generation_result = None
    st.session_state.validation_result = None
    st.session_state.subjective_review_result = None


def render_header() -> None:
    settings: ApiSettings = st.session_state.api_settings
    if settings.configured:
        interface_status = (
            f"接口已连接：{html.escape(settings.provider_name)} · "
            f"{html.escape(settings.model)}"
        )
    else:
        interface_status = "接口未配置，当前使用本地规则"
    st.markdown(
        f"""
        <div class="app-hero">
          <span class="app-kicker">本机运行 · 无需登录</span>
          <div class="app-title">本地表格助手</div>
          <div class="app-subtitle">
            用一句话生成、整理和检查电子表格。文件默认只在本机处理，
            公式、样式和客观校验由确定性程序完成。
          </div>
          <div class="status-row">
            <span class="status-pill"><span class="status-dot"></span>本地地址 127.0.0.1</span>
            <span class="status-pill">无需账号</span>
            <span class="status-pill">{interface_status}</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def current_api_form_settings() -> ApiSettings:
    nonce = st.session_state.api_nonce
    return ApiSettings(
        enabled=st.session_state.get(f"api_enabled_{nonce}", False),
        base_url=st.session_state.get(f"api_base_url_{nonce}", ""),
        api_key=st.session_state.get(f"api_key_{nonce}", ""),
        model=st.session_state.get(f"api_model_{nonce}", ""),
        provider_name=st.session_state.get(f"api_provider_{nonce}", "自定义模型"),
        timeout_seconds=st.session_state.get(f"api_timeout_{nonce}", 45),
        use_for_intent=st.session_state.get(f"api_intent_{nonce}", True),
        use_for_review=st.session_state.get(f"api_review_{nonce}", True),
    )


def render_api_settings() -> None:
    settings: ApiSettings = st.session_state.api_settings
    nonce = st.session_state.api_nonce
    status_text = (
        f"已配置：{settings.provider_name} · {settings.model} · {mask_api_key(settings.api_key)}"
        if settings.configured
        else "未配置。无需配置也可以直接使用本地分类和表格内核。"
    )
    with st.expander("接口设置", expanded=False):
        st.caption(status_text)
        st.info(
            "接口只用于理解文字需求和生成建议，不会接收完整工作簿数据，也不能修改具体单元格。"
        )
        col_a, col_b = st.columns(2)
        with col_a:
            st.text_input(
                "接口名称",
                value=settings.provider_name,
                placeholder="例如：公司模型、DeepSeek、豆包",
                key=f"api_provider_{nonce}",
            )
            st.text_input(
                "接口地址",
                value=settings.base_url,
                placeholder="例如：https://你的接口地址/v1",
                key=f"api_base_url_{nonce}",
                help="支持常见的对话补全兼容格式，也可以直接填写完整补全地址。",
            )
            st.text_input(
                "模型名称",
                value=settings.model,
                placeholder="填写接口要求的模型标识",
                key=f"api_model_{nonce}",
            )
        with col_b:
            st.text_input(
                "接口密钥",
                value=settings.api_key,
                type="password",
                placeholder="密钥只保存在本机",
                key=f"api_key_{nonce}",
            )
            st.number_input(
                "超时时间（秒）",
                min_value=5,
                max_value=180,
                value=settings.timeout_seconds,
                step=5,
                key=f"api_timeout_{nonce}",
            )
            st.checkbox(
                "启用自定义接口",
                value=settings.enabled,
                key=f"api_enabled_{nonce}",
            )

        option_a, option_b = st.columns(2)
        option_a.checkbox(
            "用于辅助理解需求",
            value=settings.use_for_intent,
            key=f"api_intent_{nonce}",
        )
        option_b.checkbox(
            "用于生成后建议审查",
            value=settings.use_for_review,
            key=f"api_review_{nonce}",
        )

        action_a, action_b, action_c = st.columns(3)
        if action_a.button("保存设置", use_container_width=True):
            updated = current_api_form_settings()
            save_api_settings(updated)
            st.session_state.api_settings = updated
            st.session_state.api_message = (
                "接口设置已保存在本机。"
                if updated.configured
                else "设置已保存，但接口尚未完整启用。"
            )
            st.session_state.api_message_kind = "success"
            st.rerun()
        if action_b.button("测试连接", use_container_width=True):
            candidate = current_api_form_settings()
            if not candidate.configured:
                st.session_state.api_message = "请先填写接口地址、密钥、模型名称并启用接口。"
                st.session_state.api_message_kind = "warning"
            else:
                with st.spinner("正在测试接口连接……"):
                    result = test_api_connection(candidate)
                st.session_state.api_message = (
                    f"连接成功，耗时约 {result.latency_ms or 0} 毫秒。"
                    if result.success
                    else result.error or "连接失败。"
                )
                st.session_state.api_message_kind = "success" if result.success else "error"
            st.rerun()
        if action_c.button("清除设置", use_container_width=True):
            delete_api_settings()
            st.session_state.api_settings = ApiSettings()
            st.session_state.api_message = "本机接口设置已清除。"
            st.session_state.api_message_kind = "success"
            st.session_state.api_nonce += 1
            st.rerun()

        message = st.session_state.api_message
        if message:
            getattr(st, st.session_state.api_message_kind, st.info)(message)


def analyze_request(prompt: str, uploads: list[Any]) -> None:
    if not prompt.strip() and not uploads:
        st.error("请先输入需求或上传文件。")
        return
    upload_snapshots = [
        {"name": item.name, "type": item.type, "data": item.getvalue()}
        for item in uploads
    ]
    filenames = [item["name"] for item in upload_snapshots]
    draft = build_task_spec_draft(prompt, filenames)
    api_plan = enhance_task_spec_draft(
        draft,
        user_prompt=prompt,
        input_file_names=filenames,
        settings=st.session_state.api_settings,
    )
    local_result = classify_intent(prompt)
    final_spec = api_plan.draft.task_spec
    matched_keywords = final_spec.options.get(
        "classification_keywords", local_result.matched_keywords
    )
    st.session_state.user_prompt = prompt
    st.session_state.uploaded_files = upload_snapshots
    st.session_state.classification_result = {
        "table_type": final_spec.task_type,
        "confidence": final_spec.confidence,
        "matched_keywords": matched_keywords,
        "alternatives": api_plan.draft.classification_alternatives,
    }
    st.session_state.clarifying_questions = api_plan.draft.clarifying_questions
    st.session_state.task_spec = final_spec
    st.session_state.clarification_done = not api_plan.draft.needs_clarification
    st.session_state.analysis_message = api_plan.message
    st.session_state.analysis_used_api = api_plan.used_api
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
    st.session_state.classification_result["table_type"] = st.session_state.task_spec.task_type
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
            except Exception as exc:
                workbook_summary = {"summary_error": f"{type(exc).__name__}: {exc}"}
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
            "suggestions": ["查看运行记录后重新提交任务。"],
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
            "user_notice": "生成失败，未运行建议审查；不影响查看错误报告。",
        }


def card(label: str, value: str) -> None:
    st.markdown(
        f"""
        <div class="spec-card">
          <div class="spec-label">{html.escape(label)}</div>
          <div class="spec-value">{html.escape(value)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_task_spec(spec: TaskSpec) -> None:
    row_a = st.columns(3)
    with row_a[0]:
        card("表格类型", TYPE_LABELS.get(spec.task_type, "通用表格"))
    with row_a[1]:
        card("输出文件", spec.output_name)
    with row_a[2]:
        card("识别把握", f"{spec.confidence:.0%}")
    st.markdown("")
    row_b = st.columns(2)
    with row_b[0]:
        card("用户目标", spec.user_goal)
    with row_b[1]:
        features = "、".join(
            [
                name
                for name, enabled in (
                    ("保留模板样式", spec.preserve_template_style),
                    ("汇总页", spec.include_summary),
                    ("图表", spec.include_charts),
                    ("说明页", spec.include_instructions_sheet),
                )
                if enabled
            ]
        ) or "基础表格"
        card("生成内容", features)

    st.markdown("**输入文件**")
    if spec.input_files:
        for item in spec.input_files:
            st.write(f"• {Path(item).name}")
    else:
        st.write("• 未提供，将使用标准模板中的示例内容。")

    st.markdown("**系统说明**")
    for assumption in spec.assumptions:
        st.write(f"• {assumption.replace('demo', '示例')}")

    with st.expander("查看技术明细"):
        technical = {
            "表格类型": TYPE_LABELS.get(spec.task_type, "通用表格"),
            "用户目标": spec.user_goal,
            "输入文件": [Path(item).name for item in spec.input_files],
            "输出文件": spec.output_name,
            "保留模板样式": spec.preserve_template_style,
            "包含图表": spec.include_charts,
            "包含汇总": spec.include_summary,
            "包含说明页": spec.include_instructions_sheet,
            "识别把握": spec.confidence,
            "系统说明": spec.assumptions,
        }
        st.json(technical)


def download_file(label: str, path: Path, mime: str, key: str, filename: str | None = None) -> None:
    if path.exists():
        st.download_button(
            label=label,
            data=path.read_bytes(),
            file_name=filename or path.name,
            mime=mime,
            key=key,
            use_container_width=True,
        )


def status_label(status: str) -> str:
    return {
        "pass": "通过",
        "warn": "有提醒",
        "fail": "未通过",
        "error": "执行异常",
        "success": "成功",
    }.get(str(status), "未知")


def render_subjective_review(subjective: dict[str, Any]) -> None:
    st.info(subjective.get("user_notice", "建议审查未启用，不影响文件下载。"))
    reviews = subjective.get("reviews", [])
    for review in reviews:
        concerns = review.get("concerns", [])
        suggestions = review.get("suggestions", [])
        st.markdown(
            f"**{review.get('model', '自定义模型')}：{status_label(review.get('status', 'warn'))}**"
        )
        if concerns:
            st.markdown("需要留意：")
            for item in concerns:
                st.write(f"• {item}")
        if suggestions:
            st.markdown("改进建议：")
            for item in suggestions:
                st.write(f"• {item}")


def render_history() -> None:
    st.subheader("最近任务")
    history = recent_tasks(10)
    if not history:
        st.caption("还没有任务记录。")
        return
    for item in history:
        cols = st.columns([2.4, 1.4, 1.4])
        cols[0].write(str(item.get("created_at", "")).replace("T", " ")[:19])
        cols[1].write(TYPE_LABELS.get(item.get("task_type"), "通用表格"))
        cols[2].write(status_label(item.get("status", "")))


initialize_state()
inject_styles()
render_header()
render_api_settings()

with st.container(border=True):
    st.subheader("开始制作")
    st.caption("描述你想要的表格，也可以同时上传现有数据。")
    prompt = st.text_area(
        "需求描述",
        value=st.session_state.user_prompt,
        placeholder="例如：根据订单数据生成销售月报，包含月度趋势和品类汇总",
        height=130,
        label_visibility="collapsed",
    )
    uploads = st.file_uploader(
        "上传本地数据",
        type=["csv", "xlsx", "xlsm"],
        accept_multiple_files=True,
        help="支持 CSV、XLSX、XLSM。旧版 XLS 请先另存为 XLSX。",
    )
    if st.button("分析需求", type="primary", use_container_width=True):
        with st.spinner("正在整理需求……"):
            analyze_request(prompt, list(uploads or []))
        st.rerun()

classification = st.session_state.classification_result
spec = st.session_state.task_spec
if classification and spec:
    with st.container(border=True):
        st.subheader("需求分析")
        if st.session_state.analysis_message:
            if st.session_state.analysis_used_api:
                st.success(st.session_state.analysis_message)
            else:
                st.info(st.session_state.analysis_message)
        st.write(
            f"识别为 **{TYPE_LABELS.get(classification['table_type'], '通用表格')}**，"
            f"识别把握约为 **{classification['confidence']:.0%}**。"
        )

        if st.session_state.clarifying_questions and not st.session_state.clarification_done:
            st.warning("还需要你一次性补充少量信息。")
            nonce = st.session_state.analysis_nonce
            for index, question in enumerate(st.session_state.clarifying_questions):
                st.text_input(question, key=f"clarification_{nonce}_{index}")
            type_index = SUPPORTED_TYPES.index(spec.task_type)
            st.selectbox(
                "确认表格类型",
                SUPPORTED_TYPES,
                index=type_index,
                format_func=lambda value: TYPE_LABELS.get(value, "通用表格"),
                key=f"clarify_type_{nonce}",
            )
            st.text_area("其他补充要求", key=f"clarify_detail_{nonce}")
            st.radio(
                "没有原始数据时",
                ["template", "upload"],
                format_func=lambda value: (
                    "先生成标准模板"
                    if value == "template"
                    else "稍后补充原始数据"
                ),
                horizontal=True,
                key=f"clarify_data_mode_{nonce}",
            )
            if st.button("确认补充信息", use_container_width=True):
                apply_clarification()
                st.rerun()
        else:
            st.success("需求信息已足够，请确认下方生成方案。")

if spec and st.session_state.clarification_done and not st.session_state.confirmed:
    nonce = st.session_state.analysis_nonce
    with st.container(border=True):
        st.subheader("确认生成方案")
        spec.output_name = st.text_input(
            "输出文件名",
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
            with st.spinner("正在生成并检查表格……"):
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
        st.subheader("生成结果")
        if generation.get("success"):
            st.success(generation.get("message", "表格生成成功。"))
        else:
            st.error(generation.get("error") or generation.get("message", "表格生成失败。"))
        for notice in generation.get("notices", []):
            st.info(str(notice).replace("demo", "示例"))
        st.markdown(
            f'<div class="result-path">任务编号：{html.escape(task_paths.task_id)}</div>',
            unsafe_allow_html=True,
        )

    with st.container(border=True):
        st.subheader("自动检查")
        status = str(validation.get("status", "error"))
        if status == "pass":
            st.success("检查通过：未发现确定性错误或警告。")
        elif status == "warn":
            st.warning("检查完成：文件可以下载，但有需要留意的提醒。")
        else:
            st.error("检查未通过：请查看问题列表；文件仍可下载进行人工检查。")
        issues = validation.get("issues", [])
        warnings = validation.get("warnings", [])
        suggestions = validation.get("suggestions", [])
        if issues:
            st.markdown("**发现的问题**")
            for item in issues:
                st.write(f"• {item.get('message', item)}")
        if warnings:
            st.markdown("**需要留意**")
            for item in warnings:
                st.write(f"• {item.get('message', item)}")
        if suggestions:
            st.markdown("**处理建议**")
            for item in suggestions:
                st.write(f"• {item}")

    with st.container(border=True):
        st.subheader("建议审查")
        render_subjective_review(subjective)

    with st.container(border=True):
        st.subheader("下载文件")
        download_cols = st.columns(4)
        with download_cols[0]:
            download_file(
                "下载表格",
                task_paths.output_file,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                f"download_excel_{task_paths.task_id}",
                filename=st.session_state.task_spec.output_name,
            )
        with download_cols[1]:
            download_file(
                "下载检查报告",
                task_paths.validation_report,
                "application/json",
                f"download_validation_{task_paths.task_id}",
                filename="检查报告.json",
            )
        with download_cols[2]:
            download_file(
                "下载生成方案",
                task_paths.task_spec_file,
                "application/json",
                f"download_spec_{task_paths.task_id}",
                filename="生成方案.json",
            )
        with download_cols[3]:
            download_file(
                "下载审查建议",
                task_paths.subjective_review_report,
                "application/json",
                f"download_review_{task_paths.task_id}",
                filename="审查建议.json",
            )

st.divider()
render_history()
