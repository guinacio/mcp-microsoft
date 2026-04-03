"""
mcp-microsoft FastMCP server.

Registers all Microsoft 365 tools (Mail, Calendar, OneDrive, etc.) from the
tools/ submodules. Entry point: mcp-microsoft (see pyproject.toml).

Run:
    python -m mcp_microsoft.server
    # or via the installed script:
    mcp-microsoft
"""

from __future__ import annotations

import logging
from contextlib import asynccontextmanager
from typing import Any

from fastmcp import FastMCP

from mcp_microsoft.config import AppConfig, get_app_config
from mcp_microsoft.feature_flags import is_sharepoint_enabled, is_teams_enabled
from mcp_microsoft.graph import close_http_clients, initialize_http_clients
from mcp_microsoft.profiles import get_profile_manager

_log = logging.getLogger(__name__)
_mcp_server: FastMCP | None = None


@asynccontextmanager
async def app_lifespan(_server: FastMCP):
    """Initialize shared runtime resources once for the server process."""
    await initialize_http_clients()
    try:
        yield {}
    finally:
        await close_http_clients()


def create_mcp_server(config: AppConfig | None = None) -> FastMCP:
    """Build a FastMCP server using the current runtime configuration."""
    runtime_config = config or get_app_config()
    get_profile_manager(config=runtime_config)
    mcp = FastMCP("mcp-microsoft", lifespan=app_lifespan)

    from mcp_microsoft.tools import attachments
    from mcp_microsoft.tools import calendar
    from mcp_microsoft.tools import contacts
    from mcp_microsoft.tools import drafts
    from mcp_microsoft.tools import folders
    from mcp_microsoft.tools import mail
    from mcp_microsoft.tools import onedrive
    from mcp_microsoft.tools import profiles
    from mcp_microsoft.tools import services
    from mcp_microsoft.tools import sharepoint

    mail.register(mcp)
    drafts.register(mcp)
    folders.register(mcp)
    attachments.register(mcp)
    calendar.register(mcp)
    onedrive.register(mcp)
    profiles.register(mcp)
    contacts.register(mcp)

    if is_sharepoint_enabled(config=runtime_config):
        sharepoint.register(mcp)
    else:
        _log.info("SharePoint tools not registered (service disabled)")

    if is_teams_enabled(config=runtime_config):
        from mcp_microsoft.tools import teams

        teams.register(mcp)
    else:
        _log.info("Teams tools not registered (service disabled)")

    services.register(mcp)
    return mcp


def get_mcp_server(reset: bool = False, config: AppConfig | None = None) -> FastMCP:
    """Return the lazily constructed FastMCP server."""
    global _mcp_server
    if reset or _mcp_server is None or config is not None:
        _mcp_server = create_mcp_server(config=config)
    return _mcp_server


def reset_mcp_server() -> None:
    """Drop the cached FastMCP server so it is rebuilt on next access."""
    global _mcp_server
    _mcp_server = None


class _MCPProxy:
    """Compatibility proxy for code that imports ``server.mcp``."""

    def __getattr__(self, name: str) -> Any:
        return getattr(get_mcp_server(), name)


mcp = _MCPProxy()


def main() -> None:
    """CLI entry point — starts the MCP server over stdio."""
    get_mcp_server().run()


if __name__ == "__main__":
    main()
