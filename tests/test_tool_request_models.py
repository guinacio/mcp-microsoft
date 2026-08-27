from __future__ import annotations

import pytest
from fastmcp import FastMCP

from mcp_microsoft.tools import contacts, mail, onedrive, profiles, services, sharepoint


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
async def test_filter_emails_schema_exposes_safe_recipient_search() -> None:
    mcp = FastMCP("test-server")
    mail.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}
    params_schema = _inner_params_schema(tools["filter_emails"])

    recipient_schema = params_schema["properties"]["to_address"]
    string_schema = next(
        variant
        for variant in recipient_schema["anyOf"]
        if variant.get("type") == "string"
    )
    assert string_schema["maxLength"] == 320
    assert "documented to: search" in recipient_schema["description"]
    assert params_schema["additionalProperties"] is False


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
