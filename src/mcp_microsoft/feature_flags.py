"""
Feature flags for optional mcp-microsoft services.

Some Microsoft 365 services (Teams, SharePoint) are only available to
work/organisational accounts.  These flags let users opt-in explicitly so
the server doesn't silently omit tools or attempt to register services that
will always fail.

Supported environment variables:
  MCP_ENABLE_TEAMS      — set to "1", "true", "yes", or "on" to enable Teams tools
  MCP_ENABLE_SHAREPOINT — set to "1", "true", "yes", or "on" to enable SharePoint tools
"""

from __future__ import annotations

import logging
import os

_log = logging.getLogger(__name__)


def _env_truthy(name: str) -> bool:
    return os.getenv(name, "").lower() in ("1", "true", "yes", "on")


def is_teams_enabled() -> bool:
    """Teams tools require a corporate/work account with a tenant_id."""
    return _env_truthy("MCP_ENABLE_TEAMS")


def is_sharepoint_enabled() -> bool:
    """SharePoint tools require a corporate/work account."""
    return _env_truthy("MCP_ENABLE_SHAREPOINT")
