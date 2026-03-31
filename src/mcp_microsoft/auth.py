"""
MSAL authentication for Microsoft Graph API.

This module is a backward-compatible facade.  All authentication logic
lives in ProfileManager (profiles.py).  The functions here delegate to
the default profile so that existing code continues to work unchanged.

For multi-account usage, call ProfileManager.get().get_token(profile)
or ProfileManager.get().get_headers(profile) directly.
"""

from __future__ import annotations

from mcp_microsoft.profiles import DEFAULT_SCOPES, ProfileManager

# Re-export for backward compatibility
SCOPES: list[str] = DEFAULT_SCOPES


def get_token(profile: str | None = None) -> str:
    """Acquire access token.  Defaults to the default profile."""
    return ProfileManager.get().get_token(profile)


def get_headers(profile: str | None = None) -> dict[str, str]:
    """Return authenticated HTTP headers.  Defaults to the default profile."""
    return ProfileManager.get().get_headers(profile)
