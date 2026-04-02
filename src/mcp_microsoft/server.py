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

from contextlib import asynccontextmanager

from fastmcp import FastMCP

from mcp_microsoft.graph import close_http_clients, initialize_http_clients

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
# Tool registration — import submodules so their @mcp.tool() decorators run
# ---------------------------------------------------------------------------

# Mail tools
from mcp_microsoft.tools import mail  # noqa: E402, F401

# Draft tools
from mcp_microsoft.tools import drafts  # noqa: E402, F401

# Folder tools
from mcp_microsoft.tools import folders  # noqa: E402, F401

# Attachment tools
from mcp_microsoft.tools import attachments  # noqa: E402, F401

# Calendar tools
from mcp_microsoft.tools import calendar  # noqa: E402, F401

# OneDrive tools
from mcp_microsoft.tools import onedrive  # noqa: E402, F401

# SharePoint tools
from mcp_microsoft.tools import sharepoint  # noqa: E402, F401

# Profile management tools
from mcp_microsoft.tools import profiles  # noqa: E402, F401

# Contacts tools
from mcp_microsoft.tools import contacts  # noqa: E402, F401

# Teams tools
from mcp_microsoft.tools import teams  # noqa: E402, F401

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    """CLI entry point — starts the MCP server over stdio."""
    mcp.run()


if __name__ == "__main__":
    main()
