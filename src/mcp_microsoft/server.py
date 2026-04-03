"""
mcp-microsoft FastMCP server.

Registers all Microsoft 365 tools (Mail, Calendar, OneDrive) from the tools/ submodules.
Entry point: mcp-microsoft (see pyproject.toml [project.scripts]).

Run:
    python -m mcp_microsoft.server
    # or via the installed script:
    mcp-microsoft
"""

from __future__ import annotations

import logging
from contextlib import asynccontextmanager

from fastmcp import FastMCP

import os

from mcp_microsoft.graph import close_http_clients, initialize_http_clients

_log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# FastMCP app
# ---------------------------------------------------------------------------

@asynccontextmanager
async def app_lifespan(_server: FastMCP):
    """Initialize shared runtime resources once for the server process."""
    await initialize_http_clients()
    try:
        yield {}
    finally:
        await close_http_clients()


mcp = FastMCP("mcp-microsoft", lifespan=app_lifespan)

# ---------------------------------------------------------------------------
# Tool registration — import submodules and call their register(mcp) functions
# ---------------------------------------------------------------------------

from mcp_microsoft.tools import mail
from mcp_microsoft.tools import drafts
from mcp_microsoft.tools import folders
from mcp_microsoft.tools import attachments
from mcp_microsoft.tools import calendar
from mcp_microsoft.tools import onedrive
from mcp_microsoft.tools import sharepoint
from mcp_microsoft.tools import profiles
from mcp_microsoft.tools import contacts

mail.register(mcp)
drafts.register(mcp)
folders.register(mcp)
attachments.register(mcp)
calendar.register(mcp)
onedrive.register(mcp)
profiles.register(mcp)
contacts.register(mcp)


def _env_flag(name: str) -> bool | None:
    """Return True/False if the env var is explicitly set, or None if absent."""
    val = os.getenv(name, "")
    if not val:
        return None
    return val.lower() in ("1", "true", "yes", "on")


def _should_register_corporate_service(env_name: str) -> bool:
    """Decide whether to register a corporate-only service (Teams / SharePoint).

    If the env var is explicitly set (MCPB toggle), respect it.
    If unset (non-MCPB / manual setup), auto-detect from account type.
    """
    flag = _env_flag(env_name)
    if flag is not None:
        return flag
    try:
        from mcp_microsoft.profiles import ProfileManager, is_corporate_account
        profile = ProfileManager.get().resolve_profile(None)
        return is_corporate_account(profile)
    except Exception:
        return False


# SharePoint — corporate accounts only.
if _should_register_corporate_service("MCP_ENABLE_SHAREPOINT"):
    sharepoint.register(mcp)
else:
    _log.info("SharePoint tools not registered (disabled or personal account)")

# Teams — corporate accounts only.
if _should_register_corporate_service("MCP_ENABLE_TEAMS"):
    from mcp_microsoft.tools import teams
    teams.register(mcp)
else:
    _log.info("Teams tools not registered (disabled or personal account)")

# Service discovery tool
from mcp_microsoft.tools import services  # noqa: E402
services.register(mcp)

# ---------------------------------------------------------------------------
# Tool annotations — infer readOnlyHint / destructiveHint / idempotentHint
# ---------------------------------------------------------------------------

from mcp_microsoft.common.tool_annotations import apply_tool_annotations  # noqa: E402

apply_tool_annotations(mcp)

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    """CLI entry point — starts the MCP server over stdio."""
    mcp.run()


if __name__ == "__main__":
    main()
