"""
Profile management tools for mcp-microsoft.

These tools allow LLM clients to list, add, remove, and authenticate
Microsoft 365 profiles at runtime without editing config files.

Implemented:
  - list_ms_profiles
  - add_ms_profile
  - remove_ms_profile
  - authenticate_ms_profile
  - set_default_ms_profile
"""

from __future__ import annotations

from typing import Optional

from mcp_microsoft.server import mcp


# ---------------------------------------------------------------------------
# list_ms_profiles
# ---------------------------------------------------------------------------


@mcp.tool()
async def list_ms_profiles() -> str:
    """
    List all configured Microsoft 365 profiles.

    Shows each profile's name, client ID (partially masked), tenant,
    authentication status, and whether it is the default.

    Returns:
        Markdown-formatted table of profiles.
    """
    from mcp_microsoft.profiles import ProfileManager

    pm = ProfileManager.get()
    profiles = pm.profiles
    default_name = pm.default_profile_name

    if not profiles:
        return "No profiles configured."

    lines = ["## Microsoft 365 Profiles\n"]
    lines.append("| Profile | Client ID | Tenant | Default | Authenticated |")
    lines.append("|---|---|---|---|---|")

    for name, cfg in sorted(profiles.items()):
        # Mask client ID: show first 8 chars + ...
        cid = cfg.client_id
        masked = f"{cid[:8]}..." if len(cid) > 8 else cid
        is_default = "Yes" if name == default_name else ""
        try:
            authed = "Yes" if pm.is_authenticated(name) else "No"
        except Exception:
            authed = "Error"
        lines.append(f"| **{name}** | `{masked}` | {cfg.tenant_id} | {is_default} | {authed} |")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# add_ms_profile
# ---------------------------------------------------------------------------


@mcp.tool()
async def add_ms_profile(
    name: str,
    client_id: str,
    tenant_id: str = "common",
    set_as_default: bool = False,
) -> str:
    """
    Add a new Microsoft 365 profile.

    Each profile represents a separate Azure App Registration / account.
    After adding, call authenticate_ms_profile to sign in interactively.

    Args:
        name: Profile name (letters, digits, hyphens, underscores only).
        client_id: Azure App Registration Application (client) ID.
        tenant_id: Azure AD tenant. Use 'common' for personal + work accounts,
                   'consumers' for personal only, 'organizations' for work only,
                   or a specific tenant domain/GUID. Defaults to 'common'.
        set_as_default: When True, make this the default profile. Defaults to False.

    Returns:
        Confirmation string.
    """
    from mcp_microsoft.profiles import ProfileManager

    pm = ProfileManager.get()
    try:
        cfg = pm.add_profile(
            name=name,
            client_id=client_id,
            tenant_id=tenant_id,
            set_as_default=set_as_default,
        )
    except ValueError as exc:
        return f"Error: {exc}"

    lines = [
        f"Profile **{cfg.name}** added successfully.",
        f"**Client ID:** `{cfg.client_id}`",
        f"**Tenant:** {cfg.tenant_id}",
        f"**Cache:** `{cfg.cache_path}`",
    ]
    if set_as_default:
        lines.append("**Set as default profile.**")
    lines.append(
        "\nCall `authenticate_ms_profile` with this profile name to sign in."
    )
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# remove_ms_profile
# ---------------------------------------------------------------------------


@mcp.tool()
async def remove_ms_profile(name: str) -> str:
    """
    Remove a Microsoft 365 profile and its cached tokens.

    Cannot remove the last remaining profile.

    Args:
        name: Profile name to remove.

    Returns:
        Confirmation string.
    """
    from mcp_microsoft.profiles import ProfileManager

    pm = ProfileManager.get()
    try:
        pm.remove_profile(name)
    except ValueError as exc:
        return f"Error: {exc}"

    return f"Profile **{name}** removed. Token cache deleted."


# ---------------------------------------------------------------------------
# authenticate_ms_profile
# ---------------------------------------------------------------------------


@mcp.tool()
async def authenticate_ms_profile(profile: Optional[str] = None) -> str:
    """
    Trigger interactive authentication for a profile.

    Opens a browser window for the Microsoft OAuth consent flow.
    Useful after adding a new profile or when tokens have expired.

    Args:
        profile: Profile name to authenticate. Omit to use the default profile.

    Returns:
        Confirmation string with the authenticated account.
    """
    from mcp_microsoft.profiles import ProfileManager

    pm = ProfileManager.get()
    try:
        cfg = pm.resolve_profile(profile)
        # Force token acquisition (will trigger interactive if needed)
        pm.get_token(cfg.name)
    except (ValueError, RuntimeError) as exc:
        return f"Authentication failed: {exc}"

    return (
        f"Profile **{cfg.name}** authenticated successfully.\n"
        f"**Tenant:** {cfg.tenant_id}\n"
        f"Tokens cached at `{cfg.cache_path}`."
    )


# ---------------------------------------------------------------------------
# set_default_ms_profile
# ---------------------------------------------------------------------------


@mcp.tool()
async def set_default_ms_profile(name: str) -> str:
    """
    Change which profile is used by default when no profile is specified.

    Args:
        name: Profile name to set as the new default.

    Returns:
        Confirmation string.
    """
    from mcp_microsoft.profiles import ProfileManager

    pm = ProfileManager.get()
    try:
        pm.set_default(name)
    except ValueError as exc:
        return f"Error: {exc}"

    return f"Default profile changed to **{name}**."
