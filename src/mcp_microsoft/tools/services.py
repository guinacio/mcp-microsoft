"""
Service discovery tool for mcp-microsoft.

Exposes a single tool that reports which Microsoft 365 service groups are
currently active, based on feature flags and account type.
"""

from __future__ import annotations

from mcp_microsoft.feature_flags import is_teams_enabled


def _is_teams_active() -> bool:
    if is_teams_enabled():
        return True
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
        "sharepoint": True,  # always registered; Graph returns 403 for personal accounts
        "teams": _is_teams_active(),
        "drafts": True,
        "folders": True,
        "attachments": True,
    }


def register(server) -> None:
    server.tool()(list_enabled_services)
