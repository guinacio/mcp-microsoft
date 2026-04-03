from __future__ import annotations

from mcp_microsoft.config import reset_app_config
from mcp_microsoft.profiles import reset_profile_manager


def reset_runtime_state() -> None:
    """Reset cached runtime singletons so env/config changes are re-read."""
    from mcp_microsoft.server import reset_mcp_server

    reset_mcp_server()
    reset_profile_manager()
    reset_app_config()
