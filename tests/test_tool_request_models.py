from __future__ import annotations

import pytest
from fastmcp import FastMCP

from mcp_microsoft.tools import contacts, drafts, mail, onedrive, profiles, services, sharepoint


def _inner_params_schema(tool) -> dict:
    schema = tool.parameters
    params_schema = schema["properties"]["params"]
    ref = params_schema["$ref"].split("/")[-1]
    return schema["$defs"][ref]


@pytest.mark.asyncio
async def test_create_contact_tool_uses_standardized_params_object() -> None:
    mcp = FastMCP("test-server")
    contacts.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}
    tool = tools["create_contact"]

    assert set(tool.parameters["properties"]) == {"params"}
    assert tool.parameters["required"] == ["params"]

    params_schema = _inner_params_schema(tool)
    assert "display_name" in params_schema["properties"]
    assert "display_name" in params_schema["required"]
    assert "display_name" not in tool.parameters["properties"]


@pytest.mark.asyncio
async def test_send_email_tool_hides_ctx_and_uses_params_object() -> None:
    mcp = FastMCP("test-server")
    mail.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}
    tool = tools["send_email"]

    assert set(tool.parameters["properties"]) == {"params"}
    assert "ctx" not in tool.parameters["properties"]

    params_schema = _inner_params_schema(tool)
    assert set(params_schema["required"]) == {"to", "subject", "body"}
    assert "confirm" in params_schema["properties"]


@pytest.mark.asyncio
async def test_create_reply_draft_schema_and_annotations() -> None:
    mcp = FastMCP("test-server")
    drafts.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}
    tool = tools["create_reply_draft"]
    params_schema = _inner_params_schema(tool)

    assert params_schema["required"] == ["message_id"]
    assert set(params_schema["properties"]) == {
        "message_id",
        "body",
        "reply_all",
        "body_type",
        "profile",
    }
    assert params_schema["properties"]["body"]["default"] is None
    assert params_schema["properties"]["reply_all"]["default"] is False
    assert set(tool.output_schema["properties"]) == {
        "success",
        "action",
        "error",
        "draft_id",
        "original_message_id",
        "reply_all",
        "body_type",
    }
    assert set(tool.output_schema["required"]) == {"success", "action"}
    assert "Mail.ReadWrite" in tool.description
    assert "not sent" in tool.description
    assert tool.annotations.destructiveHint is False
    assert tool.annotations.readOnlyHint is not True


@pytest.mark.asyncio
async def test_move_or_copy_item_tool_preserves_copy_field_name() -> None:
    mcp = FastMCP("test-server")
    onedrive.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}
    tool = tools["move_or_copy_item"]
    params_schema = _inner_params_schema(tool)

    assert "copy" in params_schema["properties"]
    assert "copy_value" not in params_schema["properties"]


@pytest.mark.asyncio
async def test_zero_arg_tools_do_not_require_empty_params_wrapper() -> None:
    mcp = FastMCP("test-server")
    profiles.register(mcp)
    services.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}

    assert tools["list_ms_profiles"].parameters == {
        "additionalProperties": False,
        "properties": {},
        "type": "object",
    }
    assert tools["list_enabled_services"].parameters == {
        "additionalProperties": False,
        "properties": {},
        "type": "object",
    }


@pytest.mark.asyncio
async def test_registered_tools_expose_human_friendly_titles() -> None:
    mcp = FastMCP("test-server")
    mail.register(mcp)
    profiles.register(mcp)
    sharepoint.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}

    assert tools["read_email"].title == "Read Email"
    assert tools["list_ms_profiles"].title == "List MS Profiles"
    assert tools["search_sharepoint_sites"].title == "Search SharePoint Sites"
