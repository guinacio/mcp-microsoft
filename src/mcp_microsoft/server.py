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

from fastmcp import FastMCP

# ---------------------------------------------------------------------------
# FastMCP app
# ---------------------------------------------------------------------------

mcp = FastMCP("mcp-microsoft")

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

# SharePoint tools — enabled by default (first run / work accounts);
# disabled only when ALL configured profiles use the "consumers" tenant.
try:
    from mcp_microsoft.profiles import ProfileManager as _PM
    _sp_profiles = _PM.get()._profiles
    if not _sp_profiles or any(
        cfg.tenant_id != "consumers"
        for cfg in _sp_profiles.values()
    ):
        from mcp_microsoft.tools import sharepoint  # noqa: E402, F401
except Exception:
    pass  # Startup issue — SharePoint tools skipped

# Profile management tools
from mcp_microsoft.tools import profiles  # noqa: E402, F401

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    """CLI entry point — starts the MCP server over stdio."""
    mcp.run()


if __name__ == "__main__":
    main()
