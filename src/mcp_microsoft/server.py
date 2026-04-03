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
sharepoint.register(mcp)
profiles.register(mcp)
contacts.register(mcp)

# Teams tools — only available for corporate/work accounts.
# Personal accounts (Outlook.com, Hotmail, Live) always receive 401 from the
# Teams Graph endpoints, so there is no value in exposing 18 tools that can
# never succeed.  We gate registration at startup based on the default profile's
# tenant_id.  If no profile is configured yet, we default to NOT registering
# (fail-safe).
def _register_teams_if_corporate() -> None:
    try:
        from mcp_microsoft.profiles import ProfileManager, is_corporate_account

        mgr = ProfileManager.get()
        # resolve_profile(None) returns the default profile; raises if none configured.
        profile = mgr.resolve_profile(None)
        if is_corporate_account(profile):
            from mcp_microsoft.tools import teams
            teams.register(mcp)
        else:
            _log.info(
                "Teams tools not registered: profile %r uses a personal Microsoft "
                "account (tenant_id=%r). Teams requires a work/school account.",
                profile.name,
                profile.tenant_id,
            )
    except Exception:
        # No profiles configured yet — skip Teams registration (fail-safe).
        _log.debug(
            "Teams tools not registered: no active profile found at startup.",
            exc_info=True,
        )


_register_teams_if_corporate()

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
