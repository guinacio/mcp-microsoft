from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

TRUTHY_ENV_VALUES = ("1", "true", "yes", "on")


@dataclass(frozen=True)
class AppConfig:
    bootstrap_client_id: str = ""
    bootstrap_tenant_id: str = "common"
    credentials_dir: Path = Path()
    enable_teams: bool | None = None
    enable_sharepoint: bool | None = None
    enable_teams_meeting_artifacts: bool | None = None
    enable_teams_ai_insights: bool | None = None
    disable_deletion_tools: bool = False
    # Transport / remote (http) mode. stdio mode ignores every MCP_HTTP_* /
    # MCP_AUTH_* value below.
    transport: str = "stdio"
    http_host: str = "127.0.0.1"
    http_port: int = 8000
    http_stateless: bool = False
    base_url: str = ""
    auth_client_id: str = ""
    auth_client_secret: str = ""
    auth_tenant_id: str = ""
    auth_required_scope: str = "mcp-access"
    # Per-client requests/second ceiling enforced by fastmcp's token-bucket
    # rate limiter in http mode. stdio mode ignores this entirely. 0 disables
    # rate limiting.
    rate_limit_rps: float = 10.0

    @classmethod
    def from_env(cls) -> "AppConfig":
        load_dotenv()
        credentials_dir = os.environ.get("MS365_CREDENTIALS_DIR", "").strip()
        if credentials_dir:
            base_dir = Path(credentials_dir)
        else:
            base_dir = Path.home() / ".microsoft-mcp"

        return cls(
            bootstrap_client_id=os.environ.get("MS365_CLIENT_ID", "").strip(),
            bootstrap_tenant_id=os.environ.get("MS365_TENANT_ID", "common").strip() or "common",
            credentials_dir=base_dir,
            enable_teams=env_flag("MCP_ENABLE_TEAMS"),
            enable_sharepoint=env_flag("MCP_ENABLE_SHAREPOINT"),
            enable_teams_meeting_artifacts=env_flag("MCP_ENABLE_TEAMS_MEETING_ARTIFACTS"),
            enable_teams_ai_insights=env_flag("MCP_ENABLE_TEAMS_AI_INSIGHTS"),
            disable_deletion_tools=bool(env_flag("MCP_DISABLE_DELETION_TOOLS")),
            transport=os.environ.get("MCP_TRANSPORT", "stdio").strip().lower() or "stdio",
            http_host=os.environ.get("MCP_HTTP_HOST", "127.0.0.1").strip() or "127.0.0.1",
            http_port=_int_env("MCP_HTTP_PORT", 8000),
            http_stateless=bool(env_flag("MCP_HTTP_STATELESS")),
            base_url=os.environ.get("MCP_BASE_URL", "").strip(),
            auth_client_id=os.environ.get("MCP_AUTH_CLIENT_ID", "").strip(),
            auth_client_secret=os.environ.get("MCP_AUTH_CLIENT_SECRET", "").strip(),
            auth_tenant_id=os.environ.get("MCP_AUTH_TENANT_ID", "").strip(),
            auth_required_scope=os.environ.get("MCP_AUTH_REQUIRED_SCOPE", "mcp-access").strip()
            or "mcp-access",
            rate_limit_rps=_float_env("MCP_RATE_LIMIT_RPS", 10.0),
        )


def env_flag(name: str) -> bool | None:
    value = os.getenv(name, "").strip()
    if not value:
        return None
    return value.lower() in TRUTHY_ENV_VALUES


def _int_env(name: str, default: int) -> int:
    raw = os.environ.get(name, "").strip()
    if not raw:
        return default
    try:
        return int(raw)
    except ValueError as exc:
        raise ValueError(f"{name} must be an integer, got {raw!r}") from exc


def _float_env(name: str, default: float) -> float:
    raw = os.environ.get(name, "").strip()
    if not raw:
        return default
    try:
        return float(raw)
    except ValueError as exc:
        raise ValueError(f"{name} must be a number, got {raw!r}") from exc


def validate_http_config(config: AppConfig) -> list[str]:
    """Return human-readable configuration problems for http transport mode.

    An empty list means the configuration is complete enough to start the
    Streamable HTTP server. Dependency-free: performs no imports, disk, or
    network I/O so it can run before any provider is constructed.
    """
    problems: list[str] = []

    if config.transport not in ("stdio", "http"):
        problems.append(
            f"MCP_TRANSPORT must be 'stdio' or 'http' (got {config.transport!r})"
        )

    if not config.base_url:
        problems.append("MCP_BASE_URL is required in http mode")
    elif not config.base_url.startswith(("http://", "https://")):
        problems.append(
            f"MCP_BASE_URL must start with http:// or https:// (got {config.base_url!r})"
        )

    if not config.auth_client_id:
        problems.append("MCP_AUTH_CLIENT_ID is required in http mode")
    if not config.auth_client_secret:
        problems.append("MCP_AUTH_CLIENT_SECRET is required in http mode")
    if not config.auth_tenant_id:
        problems.append("MCP_AUTH_TENANT_ID is required in http mode")

    if not 1 <= config.http_port <= 65535:
        problems.append(
            f"MCP_HTTP_PORT must be between 1 and 65535 (got {config.http_port})"
        )

    return problems


_config_cache: AppConfig | None = None


def get_app_config() -> AppConfig:
    global _config_cache
    if _config_cache is None:
        _config_cache = AppConfig.from_env()
    return _config_cache


def set_app_config(config: AppConfig) -> None:
    """Install *config* as the process-wide cached AppConfig.

    This is the programmatic-embedding hook: when a host builds the server via
    ``create_mcp_server(config=...)`` instead of the env-driven bootstrap, the
    global cache must point at that same object. Otherwise the many call sites
    that read ``get_app_config()`` directly (disk-tool registration gates and
    runtime path rejections in ``tools/*`` and ``graph.get_graph`` dispatch)
    would silently fall back to env-derived config and disagree with the
    explicitly threaded one — e.g. an embedder passing an http-mode config
    would get a server that skips http-mode disk gating.
    """
    global _config_cache
    _config_cache = config


def reset_app_config() -> None:
    global _config_cache
    _config_cache = None
