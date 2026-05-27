"""Feature flags for optional mcp-microsoft services.

Environment flags are authoritative when present. When a flag is absent, the
server falls back to account-type detection so manual installs with corporate
profiles can auto-enable Teams and SharePoint.

"""

from __future__ import annotations

from mcp_microsoft.config import AppConfig, env_flag, get_app_config

_EXPLICIT_ONLY_FLAGS = frozenset(
    {
        "MCP_ENABLE_TEAMS_MEETING_ARTIFACTS",
        "MCP_ENABLE_TEAMS_AI_INSIGHTS",
    }
)


def _resolve_profile_for_detection(
    profile_name: str | None,
    config: AppConfig | None = None,
):
    from mcp_microsoft.profiles import get_profile_manager

    return get_profile_manager(config=config).resolve_profile(profile_name)


def _get_configured_flag(
    env_name: str,
    config: AppConfig | None = None,
) -> bool | None:
    runtime_config = config or get_app_config()
    if env_name == "MCP_ENABLE_TEAMS":
        return runtime_config.enable_teams
    if env_name == "MCP_ENABLE_SHAREPOINT":
        return runtime_config.enable_sharepoint
    if env_name == "MCP_ENABLE_TEAMS_MEETING_ARTIFACTS":
        return runtime_config.enable_teams_meeting_artifacts
    if env_name == "MCP_ENABLE_TEAMS_AI_INSIGHTS":
        return runtime_config.enable_teams_ai_insights
    return env_flag(env_name)


def resolve_optional_service_enabled(
    env_name: str,
    profile_name: str | None = None,
    config: AppConfig | None = None,
) -> bool:
    """Resolve an optional service flag with account-type fallback.

    If the env var is set, respect it even when the value is falsy.
    If it is absent, fall back to corporate-account detection for the resolved
    profile. If no profile is available yet, fail safe to False.
    """
    explicit = _get_configured_flag(env_name, config=config)
    if env_name in _EXPLICIT_ONLY_FLAGS:
        return bool(explicit)
    if explicit is not None:
        return explicit

    try:
        from mcp_microsoft.profiles import is_corporate_account

        profile = _resolve_profile_for_detection(profile_name, config=config)
        return is_corporate_account(profile)
    except Exception:
        return False


def is_teams_enabled(config: AppConfig | None = None) -> bool:
    """Return True when Teams should be enabled for the active/default profile."""
    return resolve_optional_service_enabled("MCP_ENABLE_TEAMS", config=config)


def is_sharepoint_enabled(config: AppConfig | None = None) -> bool:
    """Return True when SharePoint should be enabled for the active/default profile."""
    return resolve_optional_service_enabled("MCP_ENABLE_SHAREPOINT", config=config)


def is_teams_meeting_artifacts_enabled(config: AppConfig | None = None) -> bool:
    """Return True when Teams meeting transcript/recording tools should be enabled."""
    runtime_config = config or get_app_config()
    if not is_teams_enabled(config=runtime_config):
        return False
    return bool(runtime_config.enable_teams_meeting_artifacts)


def is_teams_ai_insights_enabled(config: AppConfig | None = None) -> bool:
    """Return True when Teams Copilot AI insight tools should be enabled."""
    runtime_config = config or get_app_config()
    if not is_teams_enabled(config=runtime_config):
        return False
    return bool(runtime_config.enable_teams_ai_insights)


def is_deletion_disabled(config: AppConfig | None = None) -> bool:
    """Return True when destructive hard-delete tools should NOT be registered.

    Controlled by ``MCP_DISABLE_DELETION_TOOLS``. When truthy, the server
    omits all permanent-deletion tools (e.g. delete_email, delete_event,
    delete_drive_item, remove_ms_profile). Recoverable variants such as
    trash_email remain available.
    """
    runtime_config = config or get_app_config()
    return bool(runtime_config.disable_deletion_tools)
