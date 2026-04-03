"""Feature flags for optional mcp-microsoft services.

Environment flags are authoritative when present. When a flag is absent, the
server falls back to account-type detection so manual installs with corporate
profiles can auto-enable Teams and SharePoint.

Supported environment variables:
  MCP_ENABLE_TEAMS      — set to "1", "true", "yes", or "on" to enable Teams tools
  MCP_ENABLE_SHAREPOINT — set to "1", "true", "yes", or "on" to enable SharePoint tools
"""

from __future__ import annotations

import os

_TRUTHY_VALUES = ("1", "true", "yes", "on")


def env_flag(name: str) -> bool | None:
    """Return True/False if *name* is set, or None when it is absent."""
    value = os.getenv(name, "").strip()
    if not value:
        return None
    return value.lower() in _TRUTHY_VALUES


def is_flag_enabled(name: str) -> bool:
    """Return True when *name* is explicitly enabled."""
    return env_flag(name) is True


def _resolve_profile_for_detection(profile_name: str | None):
    from mcp_microsoft.profiles import ProfileManager

    return ProfileManager.get().resolve_profile(profile_name)


def resolve_optional_service_enabled(env_name: str, profile_name: str | None = None) -> bool:
    """Resolve an optional service flag with account-type fallback.

    If the env var is set, respect it even when the value is falsy.
    If it is absent, fall back to corporate-account detection for the resolved
    profile. If no profile is available yet, fail safe to False.
    """
    explicit = env_flag(env_name)
    if explicit is not None:
        return explicit

    try:
        from mcp_microsoft.profiles import is_corporate_account

        profile = _resolve_profile_for_detection(profile_name)
        return is_corporate_account(profile)
    except Exception:
        return False


def is_teams_enabled() -> bool:
    """Return True when Teams should be enabled for the active/default profile."""
    return resolve_optional_service_enabled("MCP_ENABLE_TEAMS")


def is_sharepoint_enabled() -> bool:
    """Return True when SharePoint should be enabled for the active/default profile."""
    return resolve_optional_service_enabled("MCP_ENABLE_SHAREPOINT")
