"""
Multi-account profile manager for mcp-microsoft.

Manages named profiles, each with its own Azure App Registration (client_id),
tenant, MSAL token cache, and optional scope overrides.  Provides profile-aware
token acquisition and GraphClient instances.

Profiles are stored in profiles.json.  When no profiles exist, the server starts
with zero profiles and the user must call add_ms_profile() to create the first one.
"""

from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import msal

from mcp_microsoft.config import AppConfig, get_app_config
from mcp_microsoft.feature_flags import resolve_optional_service_enabled

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Default scopes (importable by auth.py facade)
# ---------------------------------------------------------------------------

DEFAULT_SCOPES: list[str] = [
    # Mail
    "Mail.ReadWrite",
    "Mail.Send",
    # Calendar
    "Calendars.ReadWrite",
    # Contacts
    "Contacts.ReadWrite",
    # OneDrive
    "Files.ReadWrite",
]

TEAMS_SCOPES: list[str] = [
    # Teams — channels and channel messages
    "Team.ReadBasic.All",
    "Channel.ReadBasic.All",
    "Channel.Create",
    "ChannelMessage.Read.All",
    "ChannelMessage.Send",
    # Teams — chats (1:1 and group)
    "Chat.ReadWrite",
    "Chat.Create",
    # Teams — online meetings
    "OnlineMeetings.ReadWrite",
]

SHAREPOINT_SCOPES: list[str] = [
    "Sites.ReadWrite.All",
]


def build_default_scopes(profile_name: str | None = None) -> list[str]:
    """Build the consent scope list for profiles without explicit overrides."""
    scopes = list(DEFAULT_SCOPES)
    if resolve_optional_service_enabled("MCP_ENABLE_TEAMS", profile_name):
        scopes.extend(TEAMS_SCOPES)
    if resolve_optional_service_enabled("MCP_ENABLE_SHAREPOINT", profile_name):
        scopes.extend(SHAREPOINT_SCOPES)
    return list(dict.fromkeys(scopes))

# ---------------------------------------------------------------------------
# Profile name validation
# ---------------------------------------------------------------------------

_VALID_NAME_RE = re.compile(r"^[a-zA-Z0-9_-]+$")


def _validate_name(name: str) -> None:
    if not name or not _VALID_NAME_RE.match(name):
        raise ValueError(
            f"Invalid profile name {name!r}. "
            "Use only letters, digits, hyphens, and underscores."
        )


# ---------------------------------------------------------------------------
# ProfileConfig
# ---------------------------------------------------------------------------


@dataclass
class ProfileConfig:
    """Configuration for a single Microsoft 365 profile."""

    name: str
    client_id: str
    tenant_id: str = "common"
    scopes: list[str] | None = None  # None = use DEFAULT_SCOPES
    cache_path: Path = field(default_factory=lambda: Path())

    @property
    def effective_scopes(self) -> list[str]:
        if self.scopes:
            return self.scopes
        return build_default_scopes(self.name)

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}"

    def to_dict(self) -> dict[str, Any]:
        """Serialize for profiles.json (excludes computed fields)."""
        d: dict[str, Any] = {
            "client_id": self.client_id,
            "tenant_id": self.tenant_id,
        }
        if self.scopes is not None:
            d["scopes"] = self.scopes
        return d


# ---------------------------------------------------------------------------
# ProfileManager (singleton)
# ---------------------------------------------------------------------------


class ProfileManager:
    """
    Manages multi-account profiles, MSAL authentication, and GraphClient
    instances keyed by profile name.
    """

    def __init__(self, config: AppConfig | None = None) -> None:
        self._config = config or get_app_config()
        self._profiles: dict[str, ProfileConfig] = {}
        self._default_profile: str = ""
        self._base_dir: Path = self._resolve_base_dir(self._config)
        self._msal_apps: dict[str, msal.PublicClientApplication] = {}
        self._graph_clients: dict[str, Any] = {}
        self._load()

    # --- Base directory ---------------------------------------------------

    @staticmethod
    def _resolve_base_dir(config: AppConfig) -> Path:
        base = config.credentials_dir
        base.mkdir(parents=True, exist_ok=True)
        return base

    # --- Loading / saving ------------------------------------------------

    @property
    def _config_path(self) -> Path:
        return self._base_dir / "profiles.json"

    def _load(self) -> None:
        """Load profiles from profiles.json, or bootstrap from env vars."""
        if self._config_path.exists():
            self._load_from_file()
        else:
            self._bootstrap_from_env()

    def _load_from_file(self) -> None:
        """Parse profiles.json."""
        try:
            data = json.loads(self._config_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError) as exc:
            raise RuntimeError(
                f"Failed to read {self._config_path}: {exc}"
            ) from exc

        raw_profiles = data.get("profiles", {})
        if not raw_profiles:
            self._profiles = {}
            self._default_profile = ""
            return

        for name, cfg in raw_profiles.items():
            _validate_name(name)
            self._profiles[name] = ProfileConfig(
                name=name,
                client_id=cfg["client_id"],
                tenant_id=cfg.get("tenant_id", "common"),
                scopes=cfg.get("scopes"),
                cache_path=self._base_dir / f"msal_cache_{name}.json",
            )

        self._default_profile = data.get("default_profile", "")
        if self._default_profile not in self._profiles:
            # Fall back to first profile
            self._default_profile = next(iter(self._profiles))

    def _bootstrap_from_env(self) -> None:
        """Auto-create a 'default' profile from env vars (set by MCPB user_config).

        If MS365_CLIENT_ID is set, creates and persists a default profile so the
        user is ready to authenticate immediately. If not set, the server starts
        with zero profiles — the user must call add_ms_profile().
        """
        client_id = self._config.bootstrap_client_id
        if not client_id:
            return
        tenant_id = self._config.bootstrap_tenant_id
        self._profiles["default"] = ProfileConfig(
            name="default",
            client_id=client_id,
            tenant_id=tenant_id,
            cache_path=self._base_dir / "msal_cache_default.json",
        )
        self._default_profile = "default"
        self._save()

    def _save(self) -> None:
        """Persist current profiles to profiles.json."""
        data = {
            "default_profile": self._default_profile,
            "profiles": {
                name: cfg.to_dict() for name, cfg in self._profiles.items()
            },
        }
        self._config_path.write_text(
            json.dumps(data, indent=2) + "\n", encoding="utf-8"
        )

    # --- Profile resolution -----------------------------------------------

    def resolve_profile(self, profile: str | None = None) -> ProfileConfig:
        """
        Resolve a profile name to its config.

        None / "" -> default profile.  Raises ValueError if not found.
        """
        if not self._profiles:
            raise ValueError(
                "No profiles configured. Use add_ms_profile to create one."
            )
        name = profile if profile else self._default_profile
        if name not in self._profiles:
            available = ", ".join(sorted(self._profiles.keys()))
            raise ValueError(
                f"Unknown profile {name!r}. Available profiles: {available}"
            )
        return self._profiles[name]

    # --- MSAL auth --------------------------------------------------------

    def _get_msal_app(self, cfg: ProfileConfig) -> msal.PublicClientApplication:
        """Build or return a cached MSAL PublicClientApplication for a profile."""
        if cfg.name in self._msal_apps:
            return self._msal_apps[cfg.name]

        cache = msal.SerializableTokenCache()
        if cfg.cache_path.exists():
            try:
                cache.deserialize(cfg.cache_path.read_text(encoding="utf-8"))
            except Exception:
                pass  # corrupt cache — fall through to interactive auth

        app = msal.PublicClientApplication(
            client_id=cfg.client_id,
            authority=cfg.authority,
            token_cache=cache,
        )
        self._msal_apps[cfg.name] = app
        return app

    def _save_cache(self, cfg: ProfileConfig, app: msal.PublicClientApplication) -> None:
        """Persist the MSAL token cache for a profile if it has changed."""
        cache = app.token_cache  # type: ignore[attr-defined]
        if cache.has_state_changed:
            try:
                cfg.cache_path.write_text(cache.serialize(), encoding="utf-8")
            except OSError as exc:
                logger.warning(
                    "Failed to persist token cache for profile %s at %s: %s",
                    cfg.name,
                    cfg.cache_path,
                    exc,
                )

    def get_token(self, profile: str | None = None) -> str:
        """Acquire a valid access token for the given profile."""
        cfg = self.resolve_profile(profile)
        app = self._get_msal_app(cfg)
        scopes = cfg.effective_scopes

        # Try silent flow
        accounts = app.get_accounts()
        result = None
        if accounts:
            result = app.acquire_token_silent(scopes, account=accounts[0])

        # Fall back to interactive, then device code if interactive fails
        self._last_device_code_message = None
        if not result:
            try:
                result = app.acquire_token_interactive(scopes=scopes)
            except Exception:
                result = None

            if not result or "access_token" not in result:
                # Device code flow — works headless (MCPB, SSH, containers)
                flow = app.initiate_device_flow(scopes=scopes)
                if "user_code" not in flow:
                    raise RuntimeError(
                        f"Could not initiate device code flow for profile {cfg.name!r}: "
                        f"{flow.get('error_description', 'unknown error')}"
                    )
                self._last_device_code_message = flow["message"]
                import logging
                logging.getLogger(__name__).warning(
                    "Interactive auth unavailable. %s", flow["message"]
                )
                result = app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            error = result.get("error", "unknown_error")
            description = result.get("error_description", "No description available.")

            if "AADSTS65001" in description:
                raise RuntimeError(
                    f"Admin consent required for profile {cfg.name!r}: your tenant "
                    f"administrator must pre-approve this application. Share this URL "
                    f"with your IT admin:\n"
                    f"https://login.microsoftonline.com/common/adminconsent"
                    f"?client_id={cfg.client_id}\n"
                    f"Original error: {description}"
                )

            raise RuntimeError(
                f"Token acquisition failed for profile {cfg.name!r} "
                f"[{error}]: {description}"
            )

        self._save_cache(cfg, app)
        return result["access_token"]

    def get_headers(self, profile: str | None = None) -> dict[str, str]:
        """Return authenticated HTTP headers for the given profile."""
        token = self.get_token(profile)
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

    # --- GraphClient access -----------------------------------------------

    def get_graph(self, profile: str | None = None) -> Any:
        """
        Return a GraphClient bound to the resolved profile.

        Instances are cached per profile name for reuse.
        Lazy import to avoid circular dependency with graph.py.
        """
        cfg = self.resolve_profile(profile)
        if cfg.name not in self._graph_clients:
            from mcp_microsoft.graph import GraphClient

            self._graph_clients[cfg.name] = GraphClient(profile=cfg.name)
        return self._graph_clients[cfg.name]

    # --- Profile CRUD (used by management tools) --------------------------

    @property
    def default_profile_name(self) -> str:
        return self._default_profile

    @property
    def profiles(self) -> dict[str, ProfileConfig]:
        return dict(self._profiles)

    def add_profile(
        self,
        name: str,
        client_id: str,
        tenant_id: str = "common",
        scopes: list[str] | None = None,
        set_as_default: bool = False,
    ) -> ProfileConfig:
        """Add a new profile and persist to profiles.json."""
        _validate_name(name)
        if name in self._profiles:
            raise ValueError(f"Profile {name!r} already exists.")
        if not client_id or not client_id.strip():
            raise ValueError("client_id is required and cannot be empty.")
        if not tenant_id or not tenant_id.strip():
            raise ValueError("tenant_id is required and cannot be empty.")

        cfg = ProfileConfig(
            name=name,
            client_id=client_id.strip(),
            tenant_id=tenant_id.strip(),
            scopes=scopes,
            cache_path=self._base_dir / f"msal_cache_{name}.json",
        )
        self._profiles[name] = cfg

        # If this is the first profile, make it the default
        if set_as_default or not self._default_profile:
            self._default_profile = name

        self._save()
        return cfg

    def remove_profile(self, name: str) -> None:
        """Remove a profile and its cached tokens."""
        if name not in self._profiles:
            raise ValueError(f"Profile {name!r} not found.")
        if len(self._profiles) <= 1:
            raise ValueError("Cannot remove the last remaining profile.")

        cfg = self._profiles.pop(name)
        self._msal_apps.pop(name, None)
        self._graph_clients.pop(name, None)

        # Remove token cache file
        if cfg.cache_path.exists():
            try:
                cfg.cache_path.unlink()
            except OSError:
                pass

        # If we removed the default, pick the first remaining
        if self._default_profile == name:
            self._default_profile = next(iter(self._profiles))

        self._save()

    def set_default(self, name: str) -> None:
        """Change the default profile."""
        if name not in self._profiles:
            raise ValueError(f"Profile {name!r} not found.")
        self._default_profile = name
        self._save()

    def is_authenticated(self, profile: str | None = None) -> bool:
        """Check if a profile has cached tokens (without triggering interactive auth)."""
        cfg = self.resolve_profile(profile)
        app = self._get_msal_app(cfg)
        accounts = app.get_accounts()
        if not accounts:
            return False
        result = app.acquire_token_silent(cfg.effective_scopes, account=accounts[0])
        return result is not None and "access_token" in result


# ---------------------------------------------------------------------------
# Corporate-account guard (used by server.py to gate Teams tool registration)
# ---------------------------------------------------------------------------

#: Tenant IDs that are definitively personal / consumer accounts.
_PERSONAL_TENANT_IDS: frozenset[str] = frozenset(
    {
        "consumers",
        # The well-known GUID Microsoft uses as the "consumers" alias
        "9188040d-6c67-4c5b-b112-36a304b66dad",
    }
)


def is_corporate_account(profile: ProfileConfig) -> bool:
    """Return True if *profile* is a work/school (corporate) Microsoft account.

    Teams is an M365 corporate product and its Graph API endpoints always
    return 401 for personal (Outlook.com / Hotmail / Live) accounts.  Use
    this helper to decide whether to expose Teams tools.

    Decision table
    --------------
    tenant_id value          | result | reason
    -------------------------|--------|----------------------------------
    None / ""                | False  | no tenant configured — fail-safe
    "consumers"              | False  | explicit personal-account alias
    "9188040d-..."           | False  | GUID form of "consumers"
    "common"                 | True   | ambiguous sign-in audience — err toward corporate
    "organizations"          | True   | explicit corporate alias
    any GUID / named tenant  | True   | real AAD tenant → corporate
    """
    tid = (profile.tenant_id or "").strip().lower()
    if not tid:
        return False
    return tid not in _PERSONAL_TENANT_IDS


_profile_manager: ProfileManager | None = None


def get_profile_manager(config: AppConfig | None = None) -> ProfileManager:
    global _profile_manager
    if (
        _profile_manager is None
        or (config is not None and _profile_manager._config != config)
    ):
        _profile_manager = ProfileManager(config=config)
    return _profile_manager


def reset_profile_manager() -> None:
    global _profile_manager
    _profile_manager = None
