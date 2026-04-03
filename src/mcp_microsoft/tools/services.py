"""
Service discovery tool for mcp-microsoft.

Exposes a single tool that reports which Microsoft 365 service groups are
currently active, based on feature flags and account type.
"""

from __future__ import annotations

import os


def _env_flag(name: str) -> bool | None:
    """Return True/False if the env var is explicitly set, or None if absent."""
    val = os.getenv(name, "")
    if not val:
        return None
    return val.lower() in ("1", "true", "yes", "on")


def _is_corporate_service_active(env_name: str) -> bool:
    """Same logic as server.py _should_register_corporate_service."""
    flag = _env_flag(env_name)
    if flag is not None:
        return flag
    try:
        from mcp_microsoft.profiles import ProfileManager, is_corporate_account
        profile = ProfileManager.get().resolve_profile(None)
        return is_corporate_account(profile)
    except Exception:
        return False


async def list_enabled_services() -> dict:
    """List which Microsoft 365 service groups are currently active."""
    return {
        "mail": True,
        "calendar": True,
        "contacts": True,
        "onedrive": True,
        "sharepoint": _is_corporate_service_active("MCP_ENABLE_SHAREPOINT"),
        "teams": _is_corporate_service_active("MCP_ENABLE_TEAMS"),
        "drafts": True,
        "folders": True,
        "attachments": True,
    }


def register(server) -> None:
    server.tool()(list_enabled_services)
