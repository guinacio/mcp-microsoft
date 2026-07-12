from __future__ import annotations

from mcp_microsoft.config import reset_app_config
from mcp_microsoft.profiles import reset_profile_manager


def reset_runtime_state() -> None:
    """Reset cached runtime singletons so env/config changes are re-read."""
    import mcp_microsoft.graph as graph
    from mcp_microsoft.server import reset_mcp_server

    reset_mcp_server()
    reset_profile_manager()
    reset_app_config()
    # The shared http-mode GraphClient is bound to whichever transport/config
    # was live when first built; drop it so a config change is honored and a
    # client tied to a closed mock transport can't leak between tests.
    graph._obo_graph_client = None
