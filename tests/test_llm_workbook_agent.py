from excel_agent.api_settings import ApiSettings
from excel_agent.services import llm_workbook_agent as agent
from excel_agent.services.custom_api_service import ToolCall, ToolChatResult
from excel_agent.task_paths import create_task_paths
from excel_agent.task_spec import TaskSpec
from tests.test_rich_workbook_builder import sales_blueprint


def test_llm_agent_calls_local_build_tool_and_respects_chart_requirement(
    tmp_path, monkeypatch
):
    calls = iter(
        [
            ToolChatResult(
                success=True,
                content="",
                tool_calls=[
                    ToolCall(
                        id="call_1",
                        name="build_workbook",
                        arguments={"blueprint": sales_blueprint()},
                    )
                ],
                error=None,
                status_code=200,
                latency_ms=10,
                message={
                    "role": "assistant",
                    "content": None,
                    "tool_calls": [
                        {
                            "id": "call_1",
                            "type": "function",
                            "function": {
                                "name": "build_workbook",
                                "arguments": "{}",
                            },
                        }
                    ],
                },
            ),
            ToolChatResult(
                success=True,
                content="",
                tool_calls=[
                    ToolCall(
                        id="call_2",
                        name="complete_task",
                        arguments={"summary": "全部要求已完成"},
                    )
                ],
                error=None,
                status_code=200,
                latency_ms=10,
                message={
                    "role": "assistant",
                    "content": None,
                    "tool_calls": [
                        {
                            "id": "call_2",
                            "type": "function",
                            "function": {
                                "name": "complete_task",
                                "arguments": "{}",
                            },
                        }
                    ],
                },
            ),
        ]
    )

    monkeypatch.setattr(
        agent,
        "chat_completion_with_tools",
        lambda *args, **kwargs: next(calls),
    )
    paths = create_task_paths("sales_report", tmp_path / "tasks")
    spec = TaskSpec(
        task_type="sales_report",
        user_goal="生成复杂销售业绩统计表",
        include_charts=True,
    )
    settings = ApiSettings(
        enabled=True,
        base_url="https://example.com/v1",
        api_key="secret",
        model="model",
        use_for_generation=True,
    )
    result = agent.generate_with_llm_agent(spec, paths, settings)

    assert result.success is True
    assert result.tool_calls == 2
    assert paths.output_file.exists()
    assert spec.options["agent_blueprint"]["title"] == "2026年第一季度销售业绩统计表"


def test_llm_agent_corrects_plain_text_reply_and_retries_tool_call(
    tmp_path, monkeypatch
):
    calls = iter(
        [
            ToolChatResult(
                success=True,
                content="我先分析一下需求。",
                tool_calls=[],
                error=None,
                status_code=200,
                latency_ms=10,
                message={"role": "assistant", "content": "我先分析一下需求。"},
            ),
            ToolChatResult(
                success=True,
                content="",
                tool_calls=[
                    ToolCall(
                        id="call_2",
                        name="build_workbook",
                        arguments={"blueprint": sales_blueprint()},
                    )
                ],
                error=None,
                status_code=200,
                latency_ms=10,
                message={
                    "role": "assistant",
                    "content": None,
                    "tool_calls": [
                        {
                            "id": "call_2",
                            "type": "function",
                            "function": {
                                "name": "build_workbook",
                                "arguments": "{}",
                            },
                        }
                    ],
                },
            ),
            ToolChatResult(
                success=True,
                content="",
                tool_calls=[
                    ToolCall(
                        id="call_3",
                        name="complete_task",
                        arguments={"summary": "全部要求已完成"},
                    )
                ],
                error=None,
                status_code=200,
                latency_ms=10,
                message={
                    "role": "assistant",
                    "content": None,
                    "tool_calls": [
                        {
                            "id": "call_3",
                            "type": "function",
                            "function": {
                                "name": "complete_task",
                                "arguments": "{}",
                            },
                        }
                    ],
                },
            ),
        ]
    )
    monkeypatch.setattr(
        agent,
        "chat_completion_with_tools",
        lambda *args, **kwargs: next(calls),
    )
    paths = create_task_paths("sales_report", tmp_path / "tasks")
    spec = TaskSpec(
        task_type="sales_report",
        user_goal="生成复杂销售业绩统计表",
        include_charts=True,
    )
    settings = ApiSettings(
        enabled=True,
        base_url="https://example.com/v1",
        api_key="secret",
        model="model",
        use_for_generation=True,
    )

    result = agent.generate_with_llm_agent(spec, paths, settings)

    assert result.success is True
    assert result.tool_calls == 2


def test_blueprint_revision_replaces_duplicate_business_columns():
    previous = {
        "title": "预算表",
        "sheet_name": "预算",
        "columns": [
            {"key": "cat", "label": "类别", "type": "text"},
            {"key": "bud", "label": "预算金额", "type": "money"},
        ],
        "records": [{"cat": "餐饮", "bud": 1000}],
    }
    revision = {
        "title": "预算表",
        "sheet_name": "预算",
        "columns": [
            {"key": "category", "label": "类别", "type": "text"},
            {"key": "budget", "label": "预算金额", "type": "money"},
        ],
        "records": [],
    }

    merged = agent._merge_blueprint_revision(previous, revision)

    assert [item["label"] for item in merged["columns"]] == ["类别", "预算金额"]
    assert merged["records"] == [{"category": "餐饮", "budget": 1000}]


def test_uploaded_data_replaces_model_demo_records(tmp_path):
    source = tmp_path / "data.csv"
    source.write_text("姓名,成绩\n张三,95\n李四,88\n", encoding="utf-8-sig")
    spec = TaskSpec(
        task_type="generic_table",
        user_goal="根据上传数据生成成绩表",
        input_files=[str(source)],
    )
    blueprint = {
        "title": "成绩表",
        "sheet_name": "成绩",
        "columns": [
            {"key": "name", "label": "姓名", "type": "text"},
            {"key": "score", "label": "成绩", "type": "number"},
        ],
        "records": [{"name": "模型示例", "score": 100}],
    }

    attached = agent._attach_input_records(blueprint, spec)

    assert attached["records"] == [
        {"name": "张三", "score": 95},
        {"name": "李四", "score": 88},
    ]
