"""Service discovery tool for mcp-microsoft.

Exposes a single tool that reports which Microsoft 365 service groups are
currently active, based on explicit feature flags.
"""

from __future__ import annotations

from mcp_microsoft.common.tooling import LOCAL_READ_TOOL, register_tool
from mcp_microsoft.feature_flags import is_sharepoint_enabled, is_teams_enabled
from mcp_microsoft.models import ServiceFlagsResponse

async def list_enabled_services() -> ServiceFlagsResponse:
    """List which Microsoft 365 service groups are currently active."""
    return ServiceFlagsResponse(
        sharepoint=is_sharepoint_enabled(),
        teams=is_teams_enabled(),
    )


def register(server) -> None:
    register_tool(server, list_enabled_services, annotations=LOCAL_READ_TOOL)
