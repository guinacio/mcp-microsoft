"""Service discovery tool for mcp-microsoft.

Exposes a single tool that reports which Microsoft 365 service groups are
currently active, based on explicit feature flags.
"""

from __future__ import annotations

from mcp_microsoft.feature_flags import is_sharepoint_enabled, is_teams_enabled

async def list_enabled_services() -> dict:
    """List which Microsoft 365 service groups are currently active."""
    return {
        "mail": True,
        "calendar": True,
        "contacts": True,
        "onedrive": True,
        "sharepoint": is_sharepoint_enabled(),
        "teams": is_teams_enabled(),
        "drafts": True,
        "folders": True,
        "attachments": True,
    }


def register(server) -> None:
    server.tool()(list_enabled_services)
