"""
MSAL authentication helpers for Microsoft Graph API.

Thin facade over ProfileManager (profiles.py).  For multi-account usage,
call get_profile_manager().get_token(profile) directly.
"""

from __future__ import annotations

from mcp_microsoft.profiles import DEFAULT_SCOPES, get_profile_manager

SCOPES: list[str] = DEFAULT_SCOPES


def get_token(profile: str | None = None) -> str:
    """Acquire access token for the given profile."""
    return get_profile_manager().get_token(profile)


def get_headers(profile: str | None = None) -> dict[str, str]:
    """Return authenticated HTTP headers for the given profile."""
    return get_profile_manager().get_headers(profile)
