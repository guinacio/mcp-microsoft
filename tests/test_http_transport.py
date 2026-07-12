"""Tests for dual-transport config, scope building, and http-mode wiring."""

from __future__ import annotations

from pathlib import Path

import pytest

from mcp_microsoft.config import AppConfig, reset_app_config, validate_http_config
from mcp_microsoft.runtime import reset_runtime_state
from mcp_microsoft.server import build_graph_authorize_scopes

PROFILE_TOOLS = frozenset(
    {
        "list_ms_profiles",
        "add_ms_profile",
        "remove_ms_profile",
        "authenticate_ms_profile",
        "set_default_ms_profile",
    }
)

_HTTP_ENV_VARS = (
    "MCP_TRANSPORT",
    "MCP_HTTP_HOST",
    "MCP_HTTP_PORT",
    "MCP_HTTP_STATELESS",
    "MCP_BASE_URL",
    "MCP_AUTH_CLIENT_ID",
    "MCP_AUTH_CLIENT_SECRET",
    "MCP_AUTH_TENANT_ID",
    "MCP_AUTH_REQUIRED_SCOPE",
)


@pytest.fixture(autouse=True)
def _reset_cached_config() -> None:
    reset_app_config()
    yield
    reset_app_config()


def _clear_http_env(monkeypatch: pytest.MonkeyPatch) -> None:
    for name in _HTTP_ENV_VARS:
        monkeypatch.delenv(name, raising=False)


def _http_config(**overrides: object) -> AppConfig:
    base: dict[str, object] = dict(
        transport="http",
        base_url="https://mcp.example.com",
        auth_client_id="cid",
        auth_client_secret="secret",
        auth_tenant_id="organizations",
    )
    base.update(overrides)
    return AppConfig(**base)  # type: ignore[arg-type]


# --------------------------------------------------------------------------
# config.py — env parsing & normalization
# --------------------------------------------------------------------------


def test_from_env_defaults_to_stdio(monkeypatch: pytest.MonkeyPatch) -> None:
    _clear_http_env(monkeypatch)
    config = AppConfig.from_env()

    assert config.transport == "stdio"
    assert config.http_host == "127.0.0.1"
    assert config.http_port == 8000
    assert config.http_stateless is False
    assert config.base_url == ""
    assert config.auth_client_id == ""
    assert config.auth_required_scope == "mcp-access"


def test_from_env_parses_http_fields_and_normalizes_transport(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _clear_http_env(monkeypatch)
    monkeypatch.setenv("MCP_TRANSPORT", "  HTTP  ")  # mixed case + whitespace
    monkeypatch.setenv("MCP_HTTP_HOST", "0.0.0.0")
    monkeypatch.setenv("MCP_HTTP_PORT", "9001")
    monkeypatch.setenv("MCP_HTTP_STATELESS", "yes")
    monkeypatch.setenv("MCP_BASE_URL", "https://mcp.example.com")
    monkeypatch.setenv("MCP_AUTH_CLIENT_ID", "cid")
    monkeypatch.setenv("MCP_AUTH_CLIENT_SECRET", "secret")
    monkeypatch.setenv("MCP_AUTH_TENANT_ID", "tenant-guid")
    monkeypatch.setenv("MCP_AUTH_REQUIRED_SCOPE", "custom-scope")

    config = AppConfig.from_env()

    assert config.transport == "http"
    assert config.http_host == "0.0.0.0"
    assert config.http_port == 9001
    assert config.http_stateless is True
    assert config.base_url == "https://mcp.example.com"
    assert config.auth_client_id == "cid"
    assert config.auth_client_secret == "secret"
    assert config.auth_tenant_id == "tenant-guid"
    assert config.auth_required_scope == "custom-scope"


def test_from_env_blank_required_scope_falls_back_to_default(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _clear_http_env(monkeypatch)
    monkeypatch.setenv("MCP_AUTH_REQUIRED_SCOPE", "   ")
    assert AppConfig.from_env().auth_required_scope == "mcp-access"


def test_from_env_invalid_port_raises_clear_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _clear_http_env(monkeypatch)
    monkeypatch.setenv("MCP_HTTP_PORT", "not-a-number")
    with pytest.raises(ValueError, match="MCP_HTTP_PORT must be an integer"):
        AppConfig.from_env()


# --------------------------------------------------------------------------
# config.py — validate_http_config
# --------------------------------------------------------------------------


def test_validate_http_config_accepts_complete_config() -> None:
    assert validate_http_config(_http_config()) == []


def test_validate_http_config_reports_every_missing_field() -> None:
    problems = validate_http_config(AppConfig(transport="http"))
    joined = " ".join(problems)

    assert "MCP_BASE_URL" in joined
    assert "MCP_AUTH_CLIENT_ID" in joined
    assert "MCP_AUTH_CLIENT_SECRET" in joined
    assert "MCP_AUTH_TENANT_ID" in joined


def test_validate_http_config_rejects_non_http_base_url() -> None:
    problems = validate_http_config(_http_config(base_url="mcp.example.com"))
    assert any("http:// or https://" in p for p in problems)


def test_validate_http_config_rejects_invalid_transport() -> None:
    problems = validate_http_config(_http_config(transport="grpc"))
    assert any("MCP_TRANSPORT" in p for p in problems)


# --------------------------------------------------------------------------
# server.py — build_graph_authorize_scopes helper
# --------------------------------------------------------------------------


def test_graph_scopes_default_only(monkeypatch: pytest.MonkeyPatch) -> None:
    _clear_http_env(monkeypatch)
    # http mode + unset flags -> optional services resolve False (no fallback).
    scopes = build_graph_authorize_scopes(_http_config())

    assert "https://graph.microsoft.com/Mail.ReadWrite" in scopes
    assert "https://graph.microsoft.com/Files.ReadWrite" in scopes
    # No Teams/SharePoint scopes when their flags are unset.
    assert "https://graph.microsoft.com/Team.ReadBasic.All" not in scopes
    assert "https://graph.microsoft.com/Sites.ReadWrite.All" not in scopes
    # offline_access is present, unprefixed, and last.
    assert scopes[-1] == "offline_access"
    assert not any(s.endswith("/offline_access") for s in scopes)


def test_graph_scopes_include_teams_and_sharepoint_when_enabled(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _clear_http_env(monkeypatch)
    config = _http_config(
        enable_teams=True,
        enable_sharepoint=True,
        enable_teams_meeting_artifacts=True,
        enable_teams_ai_insights=True,
    )
    scopes = build_graph_authorize_scopes(config)

    assert "https://graph.microsoft.com/Team.ReadBasic.All" in scopes
    assert "https://graph.microsoft.com/OnlineMeetings.ReadWrite" in scopes
    assert "https://graph.microsoft.com/OnlineMeetingTranscript.Read.All" in scopes
    assert "https://graph.microsoft.com/OnlineMeetingAiInsight.Read.All" in scopes
    assert "https://graph.microsoft.com/Sites.ReadWrite.All" in scopes
    assert "offline_access" in scopes


def test_graph_scopes_meeting_artifacts_require_teams(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _clear_http_env(monkeypatch)
    # Artifact/insight flags on, but Teams off -> no Teams-derived scopes at all.
    config = _http_config(
        enable_teams=False,
        enable_teams_meeting_artifacts=True,
        enable_teams_ai_insights=True,
    )
    scopes = build_graph_authorize_scopes(config)

    assert "https://graph.microsoft.com/Team.ReadBasic.All" not in scopes
    assert "https://graph.microsoft.com/OnlineMeetingTranscript.Read.All" not in scopes


# --------------------------------------------------------------------------
# server.py — tool-registration matrix per transport
# --------------------------------------------------------------------------


async def _tool_names(monkeypatch: pytest.MonkeyPatch, tmp_path: Path, **env: str) -> set[str]:
    import mcp_microsoft.server as server_mod

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    monkeypatch.setenv("MS365_CLIENT_ID", "boot-client")
    monkeypatch.setenv("MS365_TENANT_ID", "common")
    for name, value in env.items():
        monkeypatch.setenv(name, value)
    reset_runtime_state()

    return {
        tool.name
        for tool in await server_mod.get_mcp_server(reset=True).list_tools(
            run_middleware=False
        )
    }


@pytest.mark.asyncio
async def test_http_mode_omits_profile_tools_but_keeps_core(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    _clear_http_env(monkeypatch)
    names = await _tool_names(
        monkeypatch,
        tmp_path,
        MCP_TRANSPORT="http",
        MCP_BASE_URL="https://mcp.example.com",
        MCP_AUTH_CLIENT_ID="cid",
        MCP_AUTH_CLIENT_SECRET="secret",
        MCP_AUTH_TENANT_ID="organizations",
    )

    assert PROFILE_TOOLS.isdisjoint(names), (
        f"profile tools must not be registered in http mode: {PROFILE_TOOLS & names}"
    )
    # Core mail/calendar tools still register.
    assert "send_email" in names
    assert "create_event" in names


@pytest.mark.asyncio
async def test_stdio_mode_registers_profile_tools(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    _clear_http_env(monkeypatch)
    names = await _tool_names(monkeypatch, tmp_path)

    assert PROFILE_TOOLS.issubset(names), (
        f"profile tools must be registered in stdio mode; missing: "
        f"{PROFILE_TOOLS - names}"
    )
    assert "send_email" in names


@pytest.mark.asyncio
async def test_http_mode_disables_corporate_profile_teams_fallback(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """A corporate ('common') default profile auto-enables Teams in stdio mode,
    but http mode must resolve optional-service flags from env only.
    """
    _clear_http_env(monkeypatch)
    # MCP_ENABLE_TEAMS/SHAREPOINT deliberately unset -> would fall back to
    # corporate-account detection in stdio mode.
    monkeypatch.delenv("MCP_ENABLE_TEAMS", raising=False)
    monkeypatch.delenv("MCP_ENABLE_SHAREPOINT", raising=False)

    names = await _tool_names(
        monkeypatch,
        tmp_path,
        MCP_TRANSPORT="http",
        MCP_BASE_URL="https://mcp.example.com",
        MCP_AUTH_CLIENT_ID="cid",
        MCP_AUTH_CLIENT_SECRET="secret",
        MCP_AUTH_TENANT_ID="organizations",
    )

    assert "teams_list_joined" not in names
    assert "search_sharepoint_sites" not in names
