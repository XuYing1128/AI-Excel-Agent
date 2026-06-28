"""Agent tool registry."""

from __future__ import annotations

from .base import AgentTool, ToolContext, ToolResult
from .excel_tools import excel_tools
from .run_python import run_python_tool


def registered_tools(ctx: ToolContext) -> list[AgentTool]:
    tools = excel_tools()
    if ctx.run_python_enabled:
        tools.append(run_python_tool())
    return tools


def tool_schemas(ctx: ToolContext) -> list[dict]:
    return [tool.openai_schema() for tool in registered_tools(ctx)]


def tool_map(ctx: ToolContext) -> dict[str, AgentTool]:
    return {tool.name: tool for tool in registered_tools(ctx)}


__all__ = [
    "AgentTool",
    "ToolContext",
    "ToolResult",
    "registered_tools",
    "tool_schemas",
    "tool_map",
]

