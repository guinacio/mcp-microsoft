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

import asyncio

from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.tooling import (
    LOCAL_DESTRUCTIVE_TOOL,
    LOCAL_IDEMPOTENT_TOOL,
    LOCAL_READ_TOOL,
    LOCAL_WRITE_TOOL,
    WRITE_TOOL,
    register_tool,
)
from mcp_microsoft.feature_flags import is_deletion_disabled
from mcp_microsoft.models import (
    AddProfileResponse,
    AddedProfileInfo,
    AuthenticateProfileResponse,
    ListProfilesResponse,
    ProfileInfo,
    RemoveProfileResponse,
    SetDefaultProfileResponse,
)
from mcp_microsoft.profiles import get_profile_manager


# ---------------------------------------------------------------------------
# list_ms_profiles
# ---------------------------------------------------------------------------

class AddProfileInput(ToolRequestModel):
    name: str
    client_id: str
    tenant_id: str = "common"
    set_as_default: bool = False


class ProfileNameInput(ToolRequestModel):
    name: str


class AuthenticateProfileInput(ToolRequestModel):
    profile: str | None = None


async def list_ms_profiles() -> ListProfilesResponse:
    """
    List all configured Microsoft 365 profiles.

    Shows each profile's name, client ID (partially masked), tenant,
    authentication status, and whether it is the default.

    Returns:
        Structured profile configuration data.
    """
    pm = get_profile_manager()
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
    params: AddProfileInput,
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
    pm = get_profile_manager()
    try:
        cfg = pm.add_profile(
            name=params.name,
            client_id=params.client_id,
            tenant_id=params.tenant_id,
            set_as_default=params.set_as_default,
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
            is_default=params.set_as_default or pm.default_profile_name == cfg.name,
        ),
    )


# ---------------------------------------------------------------------------
# remove_ms_profile
# ---------------------------------------------------------------------------


async def remove_ms_profile(params: ProfileNameInput) -> RemoveProfileResponse:
    """
    Remove a Microsoft 365 profile and its cached tokens.

    Cannot remove the last remaining profile.

    Args:
        name: Profile name to remove.

    Returns:
        Structured profile removal confirmation.
    """
    pm = get_profile_manager()
    try:
        pm.remove_profile(params.name)
    except ValueError as exc:
        return RemoveProfileResponse(
            success=False,
            action="remove_profile",
            profile=params.name,
            error=str(exc),
        )

    return RemoveProfileResponse(success=True, action="remove_profile", profile=params.name)


# ---------------------------------------------------------------------------
# authenticate_ms_profile
# ---------------------------------------------------------------------------


async def authenticate_ms_profile(params: AuthenticateProfileInput) -> AuthenticateProfileResponse:
    """
    Start or check a sign-in for a profile (non-blocking device-code flow).

    First call: starts the sign-in and returns immediately with status
    "awaiting_user", a verification_uri and a user_code. IMPORTANT: show the
    URL and code to the user in the chat — they must open the URL in a browser
    and enter the code. This tool never waits for them.

    Follow-up calls: report progress. Status stays "awaiting_user" (same code)
    until the user finishes; it becomes "authenticated" once sign-in completes.
    Tokens are then cached on disk and refreshed silently, so all other
    Microsoft 365 tools work without further sign-ins.

    Args:
        profile: Profile name to authenticate. Omit to use the default profile.

    Returns:
        Structured authentication state. When status is "awaiting_user", relay
        verification_uri and user_code to the user, then call this tool again
        after they confirm they signed in. When status is "error", calling
        again starts a fresh sign-in.
    """
    pm = get_profile_manager()
    try:
        cfg = pm.resolve_profile(params.profile)
        # begin_device_login does network I/O (silent probe / flow initiation)
        # but never waits for the user; to_thread keeps the event loop free.
        info = await asyncio.to_thread(pm.begin_device_login, cfg.name)
    except (ValueError, RuntimeError) as exc:
        return AuthenticateProfileResponse(
            success=False,
            action="authenticate_profile",
            profile=params.profile,
            status="error",
            error=str(exc),
        )

    if info["status"] == "authenticated":
        return AuthenticateProfileResponse(
            success=True,
            action="authenticate_profile",
            profile=cfg.name,
            tenant_id=cfg.tenant_id,
            cache_path=str(cfg.cache_path),
            status="authenticated",
            instructions="Signed in. All Microsoft 365 tools are ready to use.",
        )

    if info["status"] == "awaiting_user":
        return AuthenticateProfileResponse(
            success=True,
            action="authenticate_profile",
            profile=cfg.name,
            tenant_id=cfg.tenant_id,
            cache_path=str(cfg.cache_path),
            status="awaiting_user",
            user_code=info["user_code"],
            verification_uri=info["verification_uri"],
            expires_in_seconds=info["expires_in_seconds"],
            device_code_message=info["message"],
            instructions=(
                "Show verification_uri and user_code to the user now and ask "
                "them to complete the sign-in in their browser. After they "
                "confirm, call authenticate_ms_profile again to verify."
            ),
        )

    return AuthenticateProfileResponse(
        success=False,
        action="authenticate_profile",
        profile=cfg.name,
        tenant_id=cfg.tenant_id,
        status="error",
        error=info.get("error", "Authentication failed."),
        instructions="Call authenticate_ms_profile again to start a fresh sign-in.",
    )


# ---------------------------------------------------------------------------
# set_default_ms_profile
# ---------------------------------------------------------------------------


async def set_default_ms_profile(params: ProfileNameInput) -> SetDefaultProfileResponse:
    """
    Change which profile is used by default when no profile is specified.

    Args:
        name: Profile name to set as the new default.

    Returns:
        Structured default-profile update confirmation.
    """
    pm = get_profile_manager()
    try:
        pm.set_default(params.name)
    except ValueError as exc:
        return SetDefaultProfileResponse(
            success=False,
            action="set_default_profile",
            profile=params.name,
            error=str(exc),
        )

    return SetDefaultProfileResponse(
        success=True,
        action="set_default_profile",
        profile=params.name,
    )


def register(server) -> None:
    """Register all profile management tools with the given FastMCP server instance."""
    register_tool(server, list_ms_profiles, annotations=LOCAL_READ_TOOL)
    register_tool(server, add_ms_profile, annotations=LOCAL_WRITE_TOOL)
    if not is_deletion_disabled():
        register_tool(server, remove_ms_profile, annotations=LOCAL_DESTRUCTIVE_TOOL)
    register_tool(server, authenticate_ms_profile, annotations=WRITE_TOOL)
    register_tool(server, set_default_ms_profile, annotations=LOCAL_IDEMPOTENT_TOOL)
