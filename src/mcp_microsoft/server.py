"""
mcp-microsoft FastMCP server.

Registers all Microsoft 365 tools (Mail, Calendar, OneDrive, etc.) from the
tools/ submodules. Entry point: mcp-microsoft (see pyproject.toml).

Run:
    python -m mcp_microsoft.server
    # or via the installed script:
    mcp-microsoft
"""

from __future__ import annotations

import logging
from contextlib import asynccontextmanager
from typing import Any

from fastmcp import FastMCP

from mcp_microsoft.config import AppConfig, get_app_config, validate_http_config
from mcp_microsoft.feature_flags import (
    is_sharepoint_enabled,
    is_teams_ai_insights_enabled,
    is_teams_enabled,
    is_teams_meeting_artifacts_enabled,
)
from mcp_microsoft.graph import close_http_clients, initialize_http_clients
from mcp_microsoft.profiles import get_profile_manager

_log = logging.getLogger(__name__)
_mcp_server: FastMCP | None = None

# Microsoft Graph resource prefix for delegated permission URIs advertised to
# Azure via AzureProvider.additional_authorize_scopes.
_GRAPH_RESOURCE = "https://graph.microsoft.com"

# Shared rate-limit bucket key for any request whose caller identity cannot be
# resolved. Kept separate from every authenticated user's per-oid bucket so an
# unauthenticated caller can only starve other unauthenticated callers.
_UNAUTHENTICATED_CLIENT_ID = "unauthenticated"


def _rate_limit_client_id(_context: Any) -> str:
    """Return the per-client rate-limit bucket key for the current request.

    Passed to fastmcp's ``RateLimitingMiddleware(get_client_id=...)``. In
    fastmcp 3.4.4 the middleware calls this with the ``MiddlewareContext`` and
    accepts a sync ``str`` return; with no ``get_client_id`` it keys EVERY
    request under the single literal ``"global"`` bucket, letting one user
    throttle everyone. We instead key on the caller's validated Entra identity
    (``oid``, falling back to ``sub``) so each user gets an independent bucket.

    Must never raise — any failure reading the ambient token collapses to the
    shared ``"unauthenticated"`` bucket. The bearer token itself is never read
    into the key.
    """
    try:
        from fastmcp.server.dependencies import get_access_token

        token = get_access_token()
        if token is None:
            return _UNAUTHENTICATED_CLIENT_ID
        claims = getattr(token, "claims", None) or {}
        identity = claims.get("oid") or claims.get("sub")
        return str(identity) if identity else _UNAUTHENTICATED_CLIENT_ID
    except Exception:
        return _UNAUTHENTICATED_CLIENT_ID


def build_graph_authorize_scopes(config: AppConfig) -> list[str]:
    """Build the delegated Graph scopes for ``additional_authorize_scopes``.

    Mirrors ``profiles.build_default_scopes`` semantics but emits the full
    ``https://graph.microsoft.com/<Scope>`` URI form that AzureProvider forwards
    upstream. ``DEFAULT_SCOPES`` are always present; Teams/SharePoint scopes are
    added only when their feature flag resolves True (in http mode the flags are
    env-only — no corporate-profile fallback). ``offline_access`` is appended
    unprefixed because it is an OIDC scope, not a Graph resource scope.
    """
    from mcp_microsoft.profiles import (
        DEFAULT_SCOPES,
        SHAREPOINT_SCOPES,
        TEAMS_AI_INSIGHT_SCOPES,
        TEAMS_MEETING_ARTIFACT_SCOPES,
        TEAMS_SCOPES,
    )

    graph_scopes: list[str] = list(DEFAULT_SCOPES)
    if is_teams_enabled(config=config):
        graph_scopes.extend(TEAMS_SCOPES)
        if is_teams_meeting_artifacts_enabled(config=config):
            graph_scopes.extend(TEAMS_MEETING_ARTIFACT_SCOPES)
        if is_teams_ai_insights_enabled(config=config):
            graph_scopes.extend(TEAMS_AI_INSIGHT_SCOPES)
    if is_sharepoint_enabled(config=config):
        graph_scopes.extend(SHAREPOINT_SCOPES)

    deduped = list(dict.fromkeys(graph_scopes))
    scopes = [f"{_GRAPH_RESOURCE}/{scope}" for scope in deduped]
    scopes.append("offline_access")
    return scopes


def _build_http_middleware(config: AppConfig) -> list[Any]:
    """Build the middleware stack for http (multi-user) transport.

    Order matters: rate limiting runs first so throttled requests are
    rejected before audit logging (and tool execution) ever sees them. Audit
    logging always runs so every authenticated tool call that gets through
    is recorded.
    """
    from fastmcp.server.middleware.rate_limiting import RateLimitingMiddleware

    from mcp_microsoft.middleware import AuditLoggingMiddleware

    stack: list[Any] = []
    if config.rate_limit_rps > 0:
        stack.append(
            RateLimitingMiddleware(
                max_requests_per_second=config.rate_limit_rps,
                get_client_id=_rate_limit_client_id,
            )
        )
    stack.append(AuditLoggingMiddleware())
    return stack


def _register_health_route(mcp: FastMCP) -> None:
    """Register an unauthenticated ``GET /health`` for load balancers.

    ``@custom_route`` handlers are mounted outside FastMCP's
    ``RequireAuthMiddleware`` (which wraps only the MCP endpoint route
    itself), so this stays reachable without a bearer token — exactly what a
    health check needs.
    """
    from starlette.requests import Request
    from starlette.responses import JSONResponse, Response

    @mcp.custom_route("/health", methods=["GET"])
    async def health(_request: Request) -> Response:
        return JSONResponse({"status": "ok", "transport": "http"})


def _build_azure_provider(config: AppConfig):
    """Construct the FastMCP AzureProvider for http (multi-user) transport.

    Imported lazily so stdio mode never touches ``fastmcp[azure]``. The
    constructor performs no network I/O (JWKS/OBO are fetched lazily at request
    time), so it is safe to build during server construction and in tests.
    """
    from fastmcp.server.auth.providers.azure import AzureProvider

    return AzureProvider(
        client_id=config.auth_client_id,
        client_secret=config.auth_client_secret,
        tenant_id=config.auth_tenant_id,
        base_url=config.base_url,
        required_scopes=[config.auth_required_scope],
        additional_authorize_scopes=build_graph_authorize_scopes(config),
    )


@asynccontextmanager
async def app_lifespan(_server: FastMCP):
    """Initialize shared runtime resources once for the server process."""
    await initialize_http_clients()
    try:
        yield {}
    finally:
        await close_http_clients()


def create_mcp_server(config: AppConfig | None = None) -> FastMCP:
    """Build a FastMCP server using the current runtime configuration."""
    if config is not None:
        # Programmatic-embedding path: make the global cache the SAME object we
        # thread below, so the ~11 sites that read get_app_config() directly
        # (disk-tool gates, runtime path rejections, graph.get_graph dispatch)
        # can't disagree with this explicit config. See config.set_app_config.
        from mcp_microsoft.config import set_app_config

        set_app_config(config)
    runtime_config = config or get_app_config()
    http_mode = runtime_config.transport == "http"

    if not http_mode:
        # http mode never touches ProfileManager — identity always comes from
        # the caller's bearer token via OboTokenProvider (see graph.get_graph
        # and feature_flags.resolve_optional_service_enabled). Warming it up
        # here would only create ~/.microsoft-mcp on a server that will never
        # read or write it.
        get_profile_manager(config=runtime_config)

    auth = _build_azure_provider(runtime_config) if http_mode else None
    middleware = _build_http_middleware(runtime_config) if http_mode else None
    mcp = FastMCP(
        "mcp-microsoft",
        lifespan=app_lifespan,
        auth=auth,
        middleware=middleware,
        mask_error_details=True if http_mode else None,
    )

    if http_mode:
        _register_health_route(mcp)

    from mcp_microsoft.tools import attachments
    from mcp_microsoft.tools import calendar
    from mcp_microsoft.tools import contacts
    from mcp_microsoft.tools import drafts
    from mcp_microsoft.tools import folders
    from mcp_microsoft.tools import mail
    from mcp_microsoft.tools import onedrive
    from mcp_microsoft.tools import profiles
    from mcp_microsoft.tools import services
    from mcp_microsoft.tools import sharepoint

    mail.register(mcp)
    drafts.register(mcp)
    folders.register(mcp)
    attachments.register(mcp)
    calendar.register(mcp)
    onedrive.register(mcp)
    if http_mode:
        _log.info(
            "Profile-management tools not registered "
            "(http transport; identity is per-request from the bearer token)"
        )
    else:
        profiles.register(mcp)
    contacts.register(mcp)

    if is_sharepoint_enabled(config=runtime_config):
        sharepoint.register(mcp)
    else:
        _log.info("SharePoint tools not registered (service disabled)")

    if is_teams_enabled(config=runtime_config):
        from mcp_microsoft.tools import teams

        teams.register(mcp)
    else:
        _log.info("Teams tools not registered (service disabled)")

    services.register(mcp)
    return mcp


def get_mcp_server(reset: bool = False, config: AppConfig | None = None) -> FastMCP:
    """Return the lazily constructed FastMCP server."""
    global _mcp_server
    if reset or _mcp_server is None or config is not None:
        _mcp_server = create_mcp_server(config=config)
    return _mcp_server


def reset_mcp_server() -> None:
    """Drop the cached FastMCP server so it is rebuilt on next access."""
    global _mcp_server
    _mcp_server = None


class _MCPProxy:
    """Compatibility proxy for code that imports ``server.mcp``."""

    def __getattr__(self, name: str) -> Any:
        return getattr(get_mcp_server(), name)


mcp = _MCPProxy()


def main() -> None:
    """CLI entry point — starts the MCP server over stdio or Streamable HTTP."""
    config = get_app_config()
    if config.transport not in ("stdio", "http"):
        raise SystemExit(
            "Cannot start mcp-microsoft — MCP_TRANSPORT must be 'stdio' or "
            f"'http' (got {config.transport!r})"
        )
    if config.transport == "http":
        problems = validate_http_config(config)
        if problems:
            raise SystemExit(
                "Cannot start http transport — fix the following configuration "
                "problems:\n  - " + "\n  - ".join(problems)
            )
        get_mcp_server().run(
            transport="http",
            host=config.http_host,
            port=config.http_port,
            stateless_http=config.http_stateless,
        )
    else:
        get_mcp_server().run()


if __name__ == "__main__":
    main()
