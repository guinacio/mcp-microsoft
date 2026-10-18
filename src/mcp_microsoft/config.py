from __future__ import annotations

import base64
import binascii
import os
import re
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlsplit

from dotenv import load_dotenv

TRUTHY_ENV_VALUES = ("1", "true", "yes", "on")

# A directory (tenant) ID GUID: 8-4-4-4-12 hexadecimal, case-insensitive.
# MCP_AUTH_TENANT_ID must be a concrete GUID (see validate_http_config for why
# pseudo-tenants and verified domains cannot work with fastmcp's issuer pinning).
_TENANT_GUID_RE = re.compile(
    r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",
    re.IGNORECASE,
)

# The fixed tenant every personal Microsoft account (Outlook.com / Live) signs
# in through. Tokens for those accounts carry it as tid and in iss, so it is
# the concrete GUID http mode needs for personal accounts.
CONSUMER_TENANT_ID = "9188040d-6c67-4c5b-b112-36a304b66dad"


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
    # Timezone applied to naive calendar datetimes when a tool call names none
    # (MCP_DEFAULT_TIMEZONE, an IANA name such as "America/Sao_Paulo"). Empty
    # means calendar tools reject naive datetimes instead of assuming UTC.
    default_timezone: str = ""
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
    # How this server proves its own identity to Entra (authorization-code
    # redemption, refresh, and the On-Behalf-Of exchange). "secret" keeps the
    # historical client-secret flow; the others are secretless client
    # assertions: "certificate" (PEM key + cert), "managed_identity" (Azure
    # managed identity as a federated identity credential), and
    # "workload_identity" (a projected federated token file, e.g. AKS).
    auth_credential_type: str = "secret"
    auth_client_certificate_path: str = ""
    auth_client_certificate_password: str = ""
    # Client ID of a user-assigned managed identity; empty = system-assigned.
    auth_managed_identity_client_id: str = ""
    auth_federated_token_file: str = ""
    # Key that signs the FastMCP-issued tokens handed to MCP clients and keys
    # the OAuth proxy's encrypted state. Without it fastmcp derives the key from
    # the client secret, so rotating the Entra secret silently signs every
    # client out; secretless credential types require it.
    auth_jwt_signing_key: str = ""
    # Shared OAuth-proxy state (client registrations, upstream tokens) for
    # restarts and replicas. Empty = fastmcp's encrypted local file store.
    auth_redis_url: str = ""
    # Fernet key encrypting OAuth-proxy state at rest in Redis.
    auth_storage_encryption_key: str = ""
    # fnmatch-style allow-list of MCP client redirect URIs. Empty = any
    # (fastmcp default; the consent screen is then the only guard).
    auth_allowed_redirect_uris: tuple[str, ...] = ()
    # Entra authority host; sovereign clouds use e.g. login.microsoftonline.us.
    auth_authority_host: str = "login.microsoftonline.com"
    # Shared keys sealing MCP 2026-07-28 multi-round-trip request state
    # (SEP-2322). Empty = a per-process ephemeral key, which is correct for a
    # single worker; every replica behind a load balancer needs the same ring.
    request_state_keys: tuple[str, ...] = ()
    # Per-user requests/second ceiling enforced by our bounded per-user
    # (tid+oid) token-bucket rate limiter (middleware.UserRateLimitMiddleware)
    # in http mode. stdio mode ignores this entirely. 0 (or negative) disables
    # rate limiting.
    rate_limit_rps: float = 10.0
    # Observability shared secret (http mode). Empty = disabled: the /metrics,
    # /stats, and /dashboard routes are not registered at all. When set, those
    # routes require this token (Bearer or HTTP Basic password). Optional — not
    # part of validate_http_config's required set.
    stats_token: str = ""
    # Context-free file uploads via fastmcp's FileUpload app. Tri-state: an
    # explicit MCP_ENABLE_FILE_UPLOAD wins; when unset the feature defaults ON in
    # http mode (remote users have no local disk) and OFF in stdio mode (local
    # users already have local_path). See feature_flags.is_file_upload_enabled.
    enable_file_upload: bool | None = None
    # Maximum size (MB) of any single uploaded file. Must be positive; the
    # feature-flag resolver rejects <= 0. Applies in both transports.
    upload_max_mb: int = 10
    # Global ceiling (MB) on the base64-encoded footprint of ALL uploaded files
    # across every user/scope held in process memory. Bounds total RAM even when
    # many users each stay under their per-user quota. Must be positive; the
    # feature-flag resolver rejects <= 0. See feature_flags.resolve_upload_global_budget_bytes.
    upload_global_budget_mb: int = 1024

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
            default_timezone=os.environ.get("MCP_DEFAULT_TIMEZONE", "").strip(),
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
            auth_credential_type=os.environ.get("MCP_AUTH_CREDENTIAL_TYPE", "secret")
            .strip()
            .lower()
            or "secret",
            auth_client_certificate_path=os.environ.get(
                "MCP_AUTH_CLIENT_CERTIFICATE_PATH", ""
            ).strip(),
            auth_client_certificate_password=os.environ.get(
                "MCP_AUTH_CLIENT_CERTIFICATE_PASSWORD", ""
            ),
            auth_managed_identity_client_id=os.environ.get(
                "MCP_AUTH_MANAGED_IDENTITY_CLIENT_ID", ""
            ).strip(),
            auth_federated_token_file=(
                os.environ.get("MCP_AUTH_FEDERATED_TOKEN_FILE", "").strip()
                or os.environ.get("AZURE_FEDERATED_TOKEN_FILE", "").strip()
            ),
            auth_jwt_signing_key=os.environ.get("MCP_AUTH_JWT_SIGNING_KEY", "").strip(),
            auth_redis_url=os.environ.get("MCP_AUTH_REDIS_URL", "").strip(),
            auth_storage_encryption_key=os.environ.get(
                "MCP_AUTH_STORAGE_ENCRYPTION_KEY", ""
            ).strip(),
            auth_allowed_redirect_uris=_list_env("MCP_AUTH_ALLOWED_REDIRECT_URIS"),
            auth_authority_host=os.environ.get(
                "MCP_AUTH_AUTHORITY_HOST", "login.microsoftonline.com"
            ).strip()
            or "login.microsoftonline.com",
            request_state_keys=_list_env("MCP_REQUEST_STATE_KEYS"),
            rate_limit_rps=_float_env("MCP_RATE_LIMIT_RPS", 10.0),
            stats_token=os.environ.get("MCP_STATS_TOKEN", "").strip(),
            enable_file_upload=env_flag("MCP_ENABLE_FILE_UPLOAD"),
            upload_max_mb=_int_env("MCP_UPLOAD_MAX_MB", 10),
            upload_global_budget_mb=_int_env("MCP_UPLOAD_GLOBAL_BUDGET_MB", 1024),
        )


def env_flag(name: str) -> bool | None:
    value = os.getenv(name, "").strip()
    if not value:
        return None
    return value.lower() in TRUTHY_ENV_VALUES


def _list_env(name: str) -> tuple[str, ...]:
    """Parse a comma-separated env var into a tuple of non-empty entries."""
    raw = os.environ.get(name, "")
    return tuple(item.strip() for item in raw.split(",") if item.strip())


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
    elif config.base_url.startswith("http://") and not _is_loopback_url(
        config.base_url
    ):
        problems.append(
            "MCP_BASE_URL must use https:// unless it points at localhost "
            f"(got {config.base_url!r}). OAuth 2.1 and the MCP authorization "
            "spec require TLS for every non-loopback authorization endpoint; "
            "terminate TLS at your reverse proxy and set the public https URL."
        )

    if not config.auth_client_id:
        problems.append("MCP_AUTH_CLIENT_ID is required in http mode")
    problems.extend(_credential_problems(config))
    if not config.auth_tenant_id:
        problems.append("MCP_AUTH_TENANT_ID is required in http mode")
    elif not _TENANT_GUID_RE.fullmatch(config.auth_tenant_id):
        problems.append(
            "MCP_AUTH_TENANT_ID must be your directory's tenant GUID "
            "(8-4-4-4-12 hexadecimal), not "
            f"{config.auth_tenant_id!r}. Copy the Directory (tenant) ID from "
            "Azure Portal -> Microsoft Entra ID -> Overview, or use "
            f"{CONSUMER_TENANT_ID} for personal Microsoft accounts "
            "(Outlook.com / Live). Pseudo-tenants "
            "('organizations', 'common', 'consumers') and verified domains "
            "('contoso.onmicrosoft.com') are rejected because fastmcp's "
            "AzureProvider validates the token 'iss' claim against a single "
            "literal issuer URL built from this value, and real Entra tokens "
            "always carry the concrete tenant GUID -- so a pseudo-tenant or "
            "domain never matches and every request fails authentication."
        )

    if not 1 <= config.http_port <= 65535:
        problems.append(
            f"MCP_HTTP_PORT must be between 1 and 65535 (got {config.http_port})"
        )

    return problems


_CREDENTIAL_TYPES = ("secret", "certificate", "managed_identity", "workload_identity")

# Minimum size of symmetric key strings this server accepts from the
# environment (JWT signing key, request-state keys). Matches the MCP SDK's own
# request-state floor; both libraries stretch the string with a KDF, so supply
# real randomness (e.g. ``secrets.token_hex(32)``), not a passphrase.
_MIN_KEY_BYTES = 32


def _is_loopback_url(url: str) -> bool:
    host = (urlsplit(url).hostname or "").lower()
    return host in ("localhost", "127.0.0.1", "::1") or host.endswith(".localhost")


def _key_bytes(value: str) -> int:
    """Length of a key string as the SDK and fastmcp consume it (UTF-8 bytes)."""
    return len(value.encode("utf-8"))


def _is_fernet_key(value: str) -> bool:
    try:
        return len(base64.urlsafe_b64decode(value.encode("ascii"))) == 32
    except (binascii.Error, ValueError, UnicodeEncodeError):
        return False


def _credential_problems(config: AppConfig) -> list[str]:
    """Validate how the server authenticates to Entra and protects its state."""
    problems: list[str] = []
    kind = config.auth_credential_type
    secret_hint = (
        "Generate one with: python -c \"import secrets; print(secrets.token_hex(32))\""
    )

    if kind not in _CREDENTIAL_TYPES:
        problems.append(
            "MCP_AUTH_CREDENTIAL_TYPE must be one of "
            f"{', '.join(_CREDENTIAL_TYPES)} (got {kind!r})"
        )
    elif kind == "secret":
        if not config.auth_client_secret:
            problems.append(
                "MCP_AUTH_CLIENT_SECRET is required in http mode when "
                "MCP_AUTH_CREDENTIAL_TYPE=secret (the default). Prefer a "
                "secretless credential type (certificate, managed_identity, "
                "workload_identity) for production."
            )
    else:
        if config.auth_client_secret:
            problems.append(
                f"MCP_AUTH_CLIENT_SECRET must be unset when "
                f"MCP_AUTH_CREDENTIAL_TYPE={kind}: a configured secret would "
                "silently remain a valid way to impersonate this app."
            )
        if not config.auth_jwt_signing_key:
            problems.append(
                f"MCP_AUTH_JWT_SIGNING_KEY is required when "
                f"MCP_AUTH_CREDENTIAL_TYPE={kind} (there is no client secret to "
                f"derive the token-signing key from). {secret_hint}"
            )
        if kind == "certificate" and not config.auth_client_certificate_path:
            problems.append(
                "MCP_AUTH_CLIENT_CERTIFICATE_PATH is required when "
                "MCP_AUTH_CREDENTIAL_TYPE=certificate (a PEM file holding the "
                "private key and the certificate uploaded to the app registration)."
            )
        if kind == "workload_identity" and not config.auth_federated_token_file:
            problems.append(
                "MCP_AUTH_FEDERATED_TOKEN_FILE (or AZURE_FEDERATED_TOKEN_FILE) is "
                "required when MCP_AUTH_CREDENTIAL_TYPE=workload_identity."
            )

    if config.auth_jwt_signing_key and (
        _key_bytes(config.auth_jwt_signing_key) < _MIN_KEY_BYTES
    ):
        problems.append(
            f"MCP_AUTH_JWT_SIGNING_KEY must carry at least {_MIN_KEY_BYTES} bytes "
            f"of key material. {secret_hint}"
        )

    if config.auth_redis_url:
        if not config.auth_redis_url.startswith(("redis://", "rediss://", "unix://")):
            problems.append(
                "MCP_AUTH_REDIS_URL must be a redis://, rediss:// or unix:// URL"
            )
        if not _is_fernet_key(config.auth_storage_encryption_key):
            problems.append(
                "MCP_AUTH_STORAGE_ENCRYPTION_KEY must be a Fernet key when "
                "MCP_AUTH_REDIS_URL is set (OAuth state includes upstream refresh "
                "tokens and must be encrypted at rest). Generate one with: "
                "python -c \"from cryptography.fernet import Fernet; "
                "print(Fernet.generate_key().decode())\""
            )
    elif config.auth_storage_encryption_key:
        problems.append(
            "MCP_AUTH_STORAGE_ENCRYPTION_KEY only applies with MCP_AUTH_REDIS_URL; "
            "the default local store is already encrypted with a key derived "
            "from the JWT signing key."
        )

    host = config.auth_authority_host
    if "/" in host or ":" in host or not host:
        problems.append(
            "MCP_AUTH_AUTHORITY_HOST must be a bare host name such as "
            f"login.microsoftonline.com (got {host!r})"
        )

    for index, key in enumerate(config.request_state_keys):
        if _key_bytes(key) < _MIN_KEY_BYTES:
            problems.append(
                f"MCP_REQUEST_STATE_KEYS entry {index + 1} must carry at least "
                f"{_MIN_KEY_BYTES} bytes of key material. {secret_hint}"
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
