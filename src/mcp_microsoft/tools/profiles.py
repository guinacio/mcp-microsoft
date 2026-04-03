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

from mcp.types import ToolAnnotations

from mcp_microsoft.models import (
    AddProfileResponse,
    AddedProfileInfo,
    AuthenticateProfileResponse,
    ListProfilesResponse,
    ProfileInfo,
    RemoveProfileResponse,
    SetDefaultProfileResponse,
)


# ---------------------------------------------------------------------------
# list_ms_profiles
# ---------------------------------------------------------------------------

_LOCAL_READ = ToolAnnotations(readOnlyHint=True, openWorldHint=False)
_LOCAL_WRITE = ToolAnnotations(destructiveHint=False, openWorldHint=False)
_LOCAL_IDEMPOTENT = ToolAnnotations(
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=False,
)
_LOCAL_DESTRUCTIVE = ToolAnnotations(destructiveHint=True, openWorldHint=False)
_AUTH_WRITE = ToolAnnotations(destructiveHint=False, openWorldHint=True)

async def list_ms_profiles() -> ListProfilesResponse:
    """
    List all configured Microsoft 365 profiles.

    Shows each profile's name, client ID (partially masked), tenant,
    authentication status, and whether it is the default.

    Returns:
        Structured profile configuration data.
    """
    from mcp_microsoft.profiles import ProfileManager

    pm = ProfileManager.get()
    profiles = pm.profiles
    default_name = pm.default_profile_name

    items: list[ProfileInfo] = []
    for name, cfg in sorted(profiles.items()):
        cid = cfg.client_id
        masked = f"{cid[:8]}..." if len(cid) > 8 else cid
        try:
            authed = pm.is_authenticated(name)
        except Exception:
            authed = None
        items.append(
            ProfileInfo(
                name=name,
                client_id_masked=masked,
                tenant_id=cfg.tenant_id,
                is_default=name == default_name,
                is_authenticated=authed,
                cache_path=str(cfg.cache_path),
            )
        )

    return ListProfilesResponse(default_profile=default_name or None, count=len(items), profiles=items)


# ---------------------------------------------------------------------------
# add_ms_profile
# ---------------------------------------------------------------------------


async def add_ms_profile(
    name: str,
    client_id: str,
    tenant_id: str = "common",
    set_as_default: bool = False,
) -> AddProfileResponse:
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
        Structured profile creation confirmation.
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
        return AddProfileResponse(success=False, action="add_profile", error=str(exc))

    return AddProfileResponse(
        success=True,
        action="add_profile",
        profile=AddedProfileInfo(
            name=cfg.name,
            client_id=cfg.client_id,
            tenant_id=cfg.tenant_id,
            cache_path=str(cfg.cache_path),
            is_default=set_as_default or pm.default_profile_name == cfg.name,
        ),
    )


# ---------------------------------------------------------------------------
# remove_ms_profile
# ---------------------------------------------------------------------------


async def remove_ms_profile(name: str) -> RemoveProfileResponse:
    """
    Remove a Microsoft 365 profile and its cached tokens.

    Cannot remove the last remaining profile.

    Args:
        name: Profile name to remove.

    Returns:
        Structured profile removal confirmation.
    """
    from mcp_microsoft.profiles import ProfileManager

    pm = ProfileManager.get()
    try:
        pm.remove_profile(name)
    except ValueError as exc:
        return RemoveProfileResponse(success=False, action="remove_profile", profile=name, error=str(exc))

    return RemoveProfileResponse(success=True, action="remove_profile", profile=name)


# ---------------------------------------------------------------------------
# authenticate_ms_profile
# ---------------------------------------------------------------------------


async def authenticate_ms_profile(profile: Optional[str] = None) -> AuthenticateProfileResponse:
    """
    Trigger authentication for a profile.

    Tries interactive browser auth first. If no browser is available
    (e.g. MCPB, SSH, containers), falls back to device code flow and
    returns instructions to sign in at https://microsoft.com/devicelogin.

    Args:
        profile: Profile name to authenticate. Omit to use the default profile.

    Returns:
        Structured authentication result. Check device_code_message for
        sign-in instructions when running headless.
    """
    from mcp_microsoft.profiles import ProfileManager

    pm = ProfileManager.get()
    try:
        cfg = pm.resolve_profile(profile)
        # Force token acquisition (will trigger interactive or device code)
        pm.get_token(cfg.name)
    except (ValueError, RuntimeError) as exc:
        return AuthenticateProfileResponse(
            success=False,
            action="authenticate_profile",
            profile=profile,
            error=str(exc),
            device_code_message=getattr(pm, "_last_device_code_message", None),
        )

    return AuthenticateProfileResponse(
        success=True,
        action="authenticate_profile",
        profile=cfg.name,
        tenant_id=cfg.tenant_id,
        cache_path=str(cfg.cache_path),
        device_code_message=getattr(pm, "_last_device_code_message", None),
    )


# ---------------------------------------------------------------------------
# set_default_ms_profile
# ---------------------------------------------------------------------------


async def set_default_ms_profile(name: str) -> SetDefaultProfileResponse:
    """
    Change which profile is used by default when no profile is specified.

    Args:
        name: Profile name to set as the new default.

    Returns:
        Structured default-profile update confirmation.
    """
    from mcp_microsoft.profiles import ProfileManager

    pm = ProfileManager.get()
    try:
        pm.set_default(name)
    except ValueError as exc:
        return SetDefaultProfileResponse(success=False, action="set_default_profile", profile=name, error=str(exc))

    return SetDefaultProfileResponse(success=True, action="set_default_profile", profile=name)


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------


def register(server) -> None:
    """Register all profile management tools with the given FastMCP server instance."""
    server.tool(annotations=_LOCAL_READ)(list_ms_profiles)
    server.tool(annotations=_LOCAL_WRITE)(add_ms_profile)
    server.tool(annotations=_LOCAL_DESTRUCTIVE)(remove_ms_profile)
    server.tool(annotations=_AUTH_WRITE)(authenticate_ms_profile)
    server.tool(annotations=_LOCAL_IDEMPOTENT)(set_default_ms_profile)
