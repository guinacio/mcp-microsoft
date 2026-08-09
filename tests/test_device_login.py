"""Tests for the non-blocking device-code sign-in flow.

Covers ProfileManager.begin_device_login (two-phase device login), the
allow_interactive=False guard in get_token, and the authenticate_ms_profile
tool surface that relays the code to the Claude chat.
"""

from __future__ import annotations

import threading
import time
from pathlib import Path

import pytest

import mcp_microsoft.tools.profiles as tools_profiles
from mcp_microsoft.config import AppConfig
from mcp_microsoft.profiles import ProfileManager
from mcp_microsoft.tools.profiles import AuthenticateProfileInput, authenticate_ms_profile

DEVICE_MESSAGE = (
    "To sign in, use a web browser to open the page "
    "https://microsoft.com/devicelogin and enter the code ABC-123 to authenticate."
)


class FakeMsalApp:
    """Stand-in for msal.PublicClientApplication with a controllable device flow."""

    def __init__(self) -> None:
        self.accounts: list[dict] = []
        self.silent_result: dict | None = None
        self.device_result: dict = {"access_token": "tok-device"}
        self.initiate_calls = 0
        self.flow_started = threading.Event()
        self.release_poll = threading.Event()

    def get_accounts(self):
        return list(self.accounts)

    def acquire_token_silent(self, scopes, account=None):
        return self.silent_result

    def initiate_device_flow(self, scopes):
        self.initiate_calls += 1
        return {
            "user_code": "ABC-123",
            "device_code": "opaque-device-code",
            "verification_uri": "https://microsoft.com/devicelogin",
            "message": DEVICE_MESSAGE,
            "expires_at": time.time() + 900,
            "interval": 5,
        }

    def acquire_token_by_device_flow(self, flow):
        self.flow_started.set()
        if not self.release_poll.wait(timeout=10):
            return {"error": "expired_token", "error_description": "flow timed out"}
        return self.device_result


@pytest.fixture()
def pm(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> ProfileManager:
    config = AppConfig(credentials_dir=tmp_path, bootstrap_client_id="client-123")
    manager = ProfileManager(config=config)
    fake_app = FakeMsalApp()
    monkeypatch.setattr(manager, "_get_msal_app", lambda cfg: fake_app)
    manager._fake_app = fake_app  # test-only handle
    return manager


def _finish_flow(pm: ProfileManager, profile: str = "default") -> None:
    """Let the background poll thread complete and wait for it."""
    app = pm._fake_app
    assert app.flow_started.wait(timeout=5), "device flow thread never started"
    app.release_poll.set()
    session = pm._device_sessions[profile]
    session.thread.join(timeout=5)
    assert not session.thread.is_alive()


# ---------------------------------------------------------------------------
# ProfileManager.begin_device_login
# ---------------------------------------------------------------------------


def test_begin_returns_code_immediately_without_blocking(pm: ProfileManager) -> None:
    start = time.monotonic()
    info = pm.begin_device_login("default")
    elapsed = time.monotonic() - start

    assert info["status"] == "awaiting_user"
    assert info["user_code"] == "ABC-123"
    assert info["verification_uri"] == "https://microsoft.com/devicelogin"
    assert info["message"] == DEVICE_MESSAGE
    assert info["expires_in_seconds"] > 0
    # Never waits for the user (the fake poll blocks for up to 10 s).
    assert elapsed < 2


def test_repeat_calls_reuse_the_pending_flow(pm: ProfileManager) -> None:
    first = pm.begin_device_login("default")
    second = pm.begin_device_login("default")

    assert second["status"] == "awaiting_user"
    assert second["user_code"] == first["user_code"]
    assert pm._fake_app.initiate_calls == 1


def test_completed_flow_reports_authenticated_once(pm: ProfileManager) -> None:
    pm.begin_device_login("default")
    _finish_flow(pm)

    assert pm.begin_device_login("default")["status"] == "authenticated"
    # Session is cleared; the token now lives in the (real) MSAL cache, which
    # the fake simulates via the silent path.
    app = pm._fake_app
    app.accounts = [{"home_account_id": "acc"}]
    app.silent_result = {"access_token": "tok-silent"}
    assert pm.begin_device_login("default")["status"] == "authenticated"
    assert app.initiate_calls == 1


def test_failed_flow_reports_error_then_starts_fresh(pm: ProfileManager) -> None:
    app = pm._fake_app
    app.device_result = {
        "error": "expired_token",
        "error_description": "AADSTS70020: the code expired",
    }
    pm.begin_device_login("default")
    _finish_flow(pm)

    failed = pm.begin_device_login("default")
    assert failed["status"] == "error"
    assert "AADSTS70020" in failed["error"]

    # Next call starts a brand-new flow.
    app.flow_started.clear()
    app.release_poll.clear()
    retry = pm.begin_device_login("default")
    assert retry["status"] == "awaiting_user"
    assert app.initiate_calls == 2


def test_silent_hit_skips_device_flow(pm: ProfileManager) -> None:
    app = pm._fake_app
    app.accounts = [{"home_account_id": "acc"}]
    app.silent_result = {"access_token": "tok-silent"}

    assert pm.begin_device_login("default")["status"] == "authenticated"
    assert app.initiate_calls == 0


# ---------------------------------------------------------------------------
# get_token(allow_interactive=False) guard
# ---------------------------------------------------------------------------


def test_get_token_noninteractive_raises_with_guidance(pm: ProfileManager) -> None:
    with pytest.raises(RuntimeError, match="authenticate_ms_profile"):
        pm.get_token("default", allow_interactive=False)


def test_get_token_noninteractive_surfaces_pending_code(pm: ProfileManager) -> None:
    pm.begin_device_login("default")
    with pytest.raises(RuntimeError, match="ABC-123"):
        pm.get_token("default", allow_interactive=False)


# ---------------------------------------------------------------------------
# authenticate_ms_profile tool
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_tool_relays_code_then_confirms_authentication(
    pm: ProfileManager, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(tools_profiles, "get_profile_manager", lambda: pm)

    first = await authenticate_ms_profile(AuthenticateProfileInput())
    assert first.success is True
    assert first.status == "awaiting_user"
    assert first.user_code == "ABC-123"
    assert first.verification_uri == "https://microsoft.com/devicelogin"
    assert first.device_code_message == DEVICE_MESSAGE
    assert first.instructions and "user" in first.instructions

    _finish_flow(pm)

    second = await authenticate_ms_profile(AuthenticateProfileInput())
    assert second.success is True
    assert second.status == "authenticated"


@pytest.mark.asyncio
async def test_tool_reports_unknown_profile_as_error(
    pm: ProfileManager, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(tools_profiles, "get_profile_manager", lambda: pm)

    resp = await authenticate_ms_profile(AuthenticateProfileInput(profile="nope"))
    assert resp.success is False
    assert resp.status == "error"
    assert resp.error and "nope" in resp.error
