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
        )


def env_flag(name: str) -> bool | None:
    value = os.getenv(name, "").strip()
    if not value:
        return None
    return value.lower() in TRUTHY_ENV_VALUES


_config_cache: AppConfig | None = None


def get_app_config() -> AppConfig:
    global _config_cache
    if _config_cache is None:
        _config_cache = AppConfig.from_env()
    return _config_cache


def reset_app_config() -> None:
    global _config_cache
    _config_cache = None
