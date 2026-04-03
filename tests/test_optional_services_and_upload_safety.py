from __future__ import annotations

import base64
import importlib
from pathlib import Path

import pytest
from fastmcp import FastMCP

import mcp_microsoft.feature_flags as feature_flags
from mcp_microsoft.profiles import ProfileConfig, build_default_scopes, is_corporate_account
from mcp_microsoft.tools import onedrive, sharepoint


def test_build_default_scopes_excludes_optional_services_without_flags_or_profile(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("MCP_ENABLE_TEAMS", raising=False)
    monkeypatch.delenv("MCP_ENABLE_SHAREPOINT", raising=False)
    monkeypatch.setattr(
        feature_flags,
        "_resolve_profile_for_detection",
        lambda _profile_name: (_ for _ in ()).throw(ValueError("no profile")),
    )

    scopes = build_default_scopes()

    assert "Team.ReadBasic.All" not in scopes
    assert "OnlineMeetings.ReadWrite" not in scopes
    assert "Sites.ReadWrite.All" not in scopes


def test_build_default_scopes_includes_optional_services_when_enabled(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("MCP_ENABLE_TEAMS", "true")
    monkeypatch.setenv("MCP_ENABLE_SHAREPOINT", "true")

    scopes = build_default_scopes()

    assert "Team.ReadBasic.All" in scopes
    assert "OnlineMeetings.ReadWrite" in scopes
    assert "Sites.ReadWrite.All" in scopes


def test_build_default_scopes_auto_enable_optional_services_for_common_profile(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("MCP_ENABLE_TEAMS", raising=False)
    monkeypatch.delenv("MCP_ENABLE_SHAREPOINT", raising=False)
    monkeypatch.setattr(
        feature_flags,
        "_resolve_profile_for_detection",
        lambda _profile_name: ProfileConfig(name="default", client_id="client", tenant_id="common"),
    )

    scopes = build_default_scopes("default")

    assert "Team.ReadBasic.All" in scopes
    assert "OnlineMeetings.ReadWrite" in scopes
    assert "Sites.ReadWrite.All" in scopes


def test_build_default_scopes_explicit_false_overrides_corporate_fallback(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("MCP_ENABLE_TEAMS", "false")
    monkeypatch.setenv("MCP_ENABLE_SHAREPOINT", "false")
    monkeypatch.setattr(
        feature_flags,
        "_resolve_profile_for_detection",
        lambda _profile_name: ProfileConfig(name="default", client_id="client", tenant_id="common"),
    )

    scopes = build_default_scopes("default")

    assert "Team.ReadBasic.All" not in scopes
    assert "OnlineMeetings.ReadWrite" not in scopes
    assert "Sites.ReadWrite.All" not in scopes


def test_is_corporate_account_treats_common_as_corporate() -> None:
    profile = ProfileConfig(name="default", client_id="client", tenant_id="common")
    assert is_corporate_account(profile) is True


@pytest.mark.asyncio
async def test_optional_services_auto_register_for_common_profile_without_flags(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    import mcp_microsoft.server as server_mod
    from mcp_microsoft.profiles import ProfileManager

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    monkeypatch.setenv("MS365_CLIENT_ID", "client-id")
    monkeypatch.setenv("MS365_TENANT_ID", "common")
    monkeypatch.delenv("MCP_ENABLE_TEAMS", raising=False)
    monkeypatch.delenv("MCP_ENABLE_SHAREPOINT", raising=False)
    ProfileManager.reset()
    importlib.reload(server_mod)
    tool_names = {tool.name for tool in await server_mod.mcp.list_tools(run_middleware=False)}

    assert "search_sharepoint_sites" in tool_names
    assert "teams_list_joined" in tool_names

    monkeypatch.setenv("MCP_ENABLE_TEAMS", "false")
    monkeypatch.setenv("MCP_ENABLE_SHAREPOINT", "false")
    ProfileManager.reset()
    importlib.reload(server_mod)
    tool_names = {tool.name for tool in await server_mod.mcp.list_tools(run_middleware=False)}

    assert "search_sharepoint_sites" not in tool_names
    assert "teams_list_joined" not in tool_names

    monkeypatch.delenv("MS365_CLIENT_ID", raising=False)
    monkeypatch.delenv("MS365_TENANT_ID", raising=False)
    monkeypatch.delenv("MS365_CREDENTIALS_DIR", raising=False)
    monkeypatch.delenv("MCP_ENABLE_TEAMS", raising=False)
    monkeypatch.delenv("MCP_ENABLE_SHAREPOINT", raising=False)
    ProfileManager.reset()
    importlib.reload(server_mod)


@pytest.mark.asyncio
async def test_onedrive_base64_upload_uses_generated_temp_file(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    dangerous_target = tmp_path / "escape.txt"
    payload = b"hello world"
    captured: dict[str, object] = {}

    class DummyGraph:
        async def put(self, path: str, content: bytes):
            captured["path"] = path
            captured["content"] = content
            return {"id": "drive-item", "webUrl": "https://example.invalid/file"}

    monkeypatch.setattr(onedrive, "get_graph", lambda _profile: DummyGraph())

    result = await onedrive.upload_file(
        local_path=None,
        filename=str(dangerous_target),
        content_base64=base64.b64encode(payload).decode("ascii"),
    )

    assert result.success is True
    assert captured["content"] == payload
    assert result.filename == str(dangerous_target)
    assert not dangerous_target.exists()


@pytest.mark.asyncio
async def test_sharepoint_base64_upload_uses_generated_temp_file(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    dangerous_target = tmp_path / "escape.txt"
    payload = b"hello world"
    captured: dict[str, object] = {}

    class DummyGraph:
        async def put(self, path: str, content: bytes):
            captured["path"] = path
            captured["content"] = content
            return {"id": "drive-item", "webUrl": "https://example.invalid/file"}

    monkeypatch.setattr(sharepoint, "_get_sharepoint_graph", lambda _profile: DummyGraph())

    result = await sharepoint.upload_to_site(
        "site-id",
        "drive-id",
        local_path=None,
        filename=str(dangerous_target),
        content_base64=base64.b64encode(payload).decode("ascii"),
    )

    assert result.success is True
    assert captured["content"] == payload
    assert result.filename == str(dangerous_target)
    assert not dangerous_target.exists()


@pytest.mark.asyncio
async def test_sharepoint_upload_tool_exposes_local_path_as_optional() -> None:
    mcp = FastMCP("test-server")
    sharepoint.register(mcp)

    tools = await mcp.list_tools(run_middleware=False)
    upload_tool = next(tool for tool in tools if tool.name == "upload_to_site")
    required = upload_tool.parameters.get("required", [])

    assert "local_path" not in required
