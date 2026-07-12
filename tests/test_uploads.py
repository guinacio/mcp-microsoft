"""Tests for the context-free file-upload provider (uploads.ScopedFileUpload).

Covers per-caller scoping, the bounded store (file/byte quotas, whole-scope LRU
eviction, idle-TTL prune, name-length cap), the resolve_uploaded_file accessor,
and the Graph upload-tool integration (uploaded_file resolves bytes into the
existing PUT path; mutual-exclusion and feature-off errors).
"""

from __future__ import annotations

import base64

import pytest

import mcp_microsoft.config as config_module
from mcp_microsoft.config import AppConfig
from mcp_microsoft.uploads import (
    ScopedFileUpload,
    get_upload_provider,
    reset_upload_provider,
    resolve_uploaded_file,
    set_upload_provider,
)


class _FakeCtx:
    """Duck-typed Context exposing only what _get_scope_key reads."""

    def __init__(self, session_id: str | None = "sess-1") -> None:
        self.session_id = session_id


class _FakeAccessToken:
    def __init__(self, oid: str | None) -> None:
        self.token = "assertion"
        self.claims = {"oid": oid} if oid else {}


def _file(name: str, data: bytes, content_type: str = "application/octet-stream") -> dict:
    return {
        "name": name,
        "size": len(data),
        "type": content_type,
        "data": base64.b64encode(data).decode("ascii"),
    }


@pytest.fixture(autouse=True)
def _reset() -> None:
    reset_upload_provider()
    config_module.reset_app_config()
    yield
    reset_upload_provider()
    config_module.reset_app_config()


def _force_stdio(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        config_module, "get_app_config", lambda: AppConfig(transport="stdio")
    )


def _force_http(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        config_module, "get_app_config", lambda: AppConfig(transport="http")
    )


def _set_ambient_oid(monkeypatch: pytest.MonkeyPatch, oid: str | None) -> None:
    import fastmcp.server.dependencies as deps

    monkeypatch.setattr(deps, "get_access_token", lambda: _FakeAccessToken(oid))


# ---------------------------------------------------------------------------
# quota enforcement
# ---------------------------------------------------------------------------


def test_on_store_enforces_per_scope_file_count(monkeypatch: pytest.MonkeyPatch) -> None:
    _force_stdio(monkeypatch)
    provider = ScopedFileUpload(max_files_per_scope=2)
    ctx = _FakeCtx()

    provider.on_store([_file("a.bin", b"a"), _file("b.bin", b"b")], ctx)
    with pytest.raises(ValueError, match="per-user limit of 2"):
        provider.on_store([_file("c.bin", b"c")], ctx)


def test_on_store_overwrite_does_not_count_against_file_limit(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _force_stdio(monkeypatch)
    provider = ScopedFileUpload(max_files_per_scope=1)
    ctx = _FakeCtx()

    provider.on_store([_file("a.bin", b"a")], ctx)
    # Re-storing the same name overwrites in place (still 1 file), not rejected.
    summaries = provider.on_store([_file("a.bin", b"aaaa")], ctx)
    assert len(summaries) == 1
    assert summaries[0]["size"] == 4


def test_on_store_enforces_per_scope_byte_quota(monkeypatch: pytest.MonkeyPatch) -> None:
    _force_stdio(monkeypatch)
    provider = ScopedFileUpload(max_bytes_per_scope=10)
    ctx = _FakeCtx()

    provider.on_store([_file("a.bin", b"x" * 8)], ctx)
    with pytest.raises(ValueError, match="per-user limit"):
        provider.on_store([_file("b.bin", b"y" * 8)], ctx)


def test_on_store_rejects_overlong_name(monkeypatch: pytest.MonkeyPatch) -> None:
    _force_stdio(monkeypatch)
    provider = ScopedFileUpload(max_name_length=10)
    ctx = _FakeCtx()

    with pytest.raises(ValueError, match="name is too long"):
        provider.on_store([_file("x" * 11, b"data")], ctx)


def test_on_store_rejects_atomically(monkeypatch: pytest.MonkeyPatch) -> None:
    """An over-quota batch stores none of its files (validation precedes commit)."""
    _force_stdio(monkeypatch)
    provider = ScopedFileUpload(max_files_per_scope=1)
    ctx = _FakeCtx()

    with pytest.raises(ValueError):
        provider.on_store([_file("a.bin", b"a"), _file("b.bin", b"b")], ctx)
    assert provider.on_list(ctx) == []


# ---------------------------------------------------------------------------
# scope isolation (http mode, keyed on oid)
# ---------------------------------------------------------------------------


def test_scope_isolation_between_two_oids(monkeypatch: pytest.MonkeyPatch) -> None:
    _force_http(monkeypatch)
    provider = ScopedFileUpload()
    ctx = _FakeCtx(session_id=None)  # http mode ignores session; keys on oid

    _set_ambient_oid(monkeypatch, "user-A")
    provider.on_store([_file("secret.txt", b"A-data")], ctx)
    assert {s["name"] for s in provider.on_list(ctx)} == {"secret.txt"}

    # A different caller sees an empty area, not user-A's file.
    _set_ambient_oid(monkeypatch, "user-B")
    assert provider.on_list(ctx) == []
    with pytest.raises(ValueError, match="not found"):
        provider.on_read("secret.txt", ctx)


def test_scope_key_falls_back_without_oid(monkeypatch: pytest.MonkeyPatch) -> None:
    _force_http(monkeypatch)
    provider = ScopedFileUpload()
    _set_ambient_oid(monkeypatch, None)

    # No oid, no session id -> shared default bucket, never raises.
    assert provider._get_scope_key(_FakeCtx(session_id=None)) == "__default__"
    # No oid but a session id -> session fallback.
    assert provider._get_scope_key(_FakeCtx(session_id="s9")) == "sid:s9"


# ---------------------------------------------------------------------------
# whole-scope LRU eviction + idle-TTL prune (fake clock)
# ---------------------------------------------------------------------------


def test_lru_scope_eviction(monkeypatch: pytest.MonkeyPatch) -> None:
    _force_http(monkeypatch)
    provider = ScopedFileUpload(max_scopes=2)

    for oid in ("u1", "u2", "u3"):
        _set_ambient_oid(monkeypatch, oid)
        provider.on_store([_file("f.bin", b"x")], _FakeCtx(session_id=None))

    # u1 (least recently used) was evicted when u3 pushed over the cap.
    assert provider.scopes_evicted == 1
    _set_ambient_oid(monkeypatch, "u1")
    assert provider.on_list(_FakeCtx(session_id=None)) == []
    _set_ambient_oid(monkeypatch, "u3")
    assert len(provider.on_list(_FakeCtx(session_id=None))) == 1


def test_idle_ttl_prune(monkeypatch: pytest.MonkeyPatch) -> None:
    _force_http(monkeypatch)
    clock = {"t": 1000.0}
    provider = ScopedFileUpload(idle_ttl=100.0, clock=lambda: clock["t"])

    _set_ambient_oid(monkeypatch, "old")
    provider.on_store([_file("f.bin", b"x")], _FakeCtx(session_id=None))

    # Advance past the TTL; a store for a different scope triggers the sweep.
    clock["t"] = 1000.0 + 101.0
    _set_ambient_oid(monkeypatch, "fresh")
    provider.on_store([_file("g.bin", b"y")], _FakeCtx(session_id=None))

    assert provider.scopes_pruned == 1
    _set_ambient_oid(monkeypatch, "old")
    assert provider.on_list(_FakeCtx(session_id=None)) == []


# ---------------------------------------------------------------------------
# resolve_uploaded_file
# ---------------------------------------------------------------------------


def test_resolve_uploaded_file_happy(monkeypatch: pytest.MonkeyPatch) -> None:
    _force_http(monkeypatch)
    _set_ambient_oid(monkeypatch, "resolver-user")
    provider = ScopedFileUpload()
    provider.on_store(
        [_file("doc.pdf", b"PDF-BYTES", content_type="application/pdf")],
        _FakeCtx(session_id=None),
    )
    set_upload_provider(provider)

    data, content_type = resolve_uploaded_file("doc.pdf")
    assert data == b"PDF-BYTES"
    assert content_type == "application/pdf"


def test_resolve_uploaded_file_missing_lists_available(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _force_http(monkeypatch)
    _set_ambient_oid(monkeypatch, "resolver-user")
    provider = ScopedFileUpload()
    provider.on_store([_file("present.txt", b"hi")], _FakeCtx(session_id=None))
    set_upload_provider(provider)

    with pytest.raises(ValueError, match="present.txt"):
        resolve_uploaded_file("absent.txt")


def test_resolve_uploaded_file_feature_off_raises() -> None:
    set_upload_provider(None)
    with pytest.raises(ValueError, match="MCP_ENABLE_FILE_UPLOAD"):
        resolve_uploaded_file("anything")


# ---------------------------------------------------------------------------
# tool integration — upload_file (onedrive)
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_upload_file_resolves_uploaded_file_into_put(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from mcp_microsoft.tools import onedrive

    monkeypatch.setattr(onedrive, "get_app_config", lambda: AppConfig(transport="http"))
    _force_http(monkeypatch)
    _set_ambient_oid(monkeypatch, "uploader")

    provider = ScopedFileUpload()
    provider.on_store([_file("report.txt", b"hello")], _FakeCtx(session_id=None))
    set_upload_provider(provider)

    captured: dict[str, object] = {}

    class DummyGraph:
        async def put(self, path: str, content: bytes) -> dict:
            captured["path"] = path
            captured["content"] = content
            return {"id": "drive-item", "webUrl": "https://example.invalid/f"}

    monkeypatch.setattr(onedrive, "get_graph", lambda _profile: DummyGraph())

    result = await onedrive.upload_file(
        onedrive.UploadFileInput(uploaded_file="report.txt")
    )

    assert result.success is True
    assert result.filename == "report.txt"
    assert captured["content"] == b"hello"
    assert "report.txt" in str(captured["path"])


@pytest.mark.asyncio
async def test_upload_file_uploaded_file_mutually_exclusive(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from mcp_microsoft.tools import onedrive

    monkeypatch.setattr(onedrive, "get_app_config", lambda: AppConfig(transport="http"))

    with pytest.raises(ValueError, match="cannot be combined"):
        await onedrive.upload_file(
            onedrive.UploadFileInput(
                uploaded_file="report.txt",
                content_base64=base64.b64encode(b"x").decode("ascii"),
            )
        )


@pytest.mark.asyncio
async def test_upload_file_uploaded_file_feature_off(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from mcp_microsoft.tools import onedrive

    monkeypatch.setattr(onedrive, "get_app_config", lambda: AppConfig(transport="http"))
    set_upload_provider(None)

    with pytest.raises(ValueError, match="not enabled"):
        await onedrive.upload_file(onedrive.UploadFileInput(uploaded_file="report.txt"))


@pytest.mark.asyncio
async def test_upload_to_site_resolves_uploaded_file_into_put(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from mcp_microsoft.tools import sharepoint

    monkeypatch.setattr(
        sharepoint, "get_app_config", lambda: AppConfig(transport="http")
    )
    _force_http(monkeypatch)
    _set_ambient_oid(monkeypatch, "uploader")

    provider = ScopedFileUpload()
    provider.on_store([_file("deck.pptx", b"SLIDES")], _FakeCtx(session_id=None))
    set_upload_provider(provider)

    captured: dict[str, object] = {}

    class DummyGraph:
        async def put(self, path: str, content: bytes) -> dict:
            captured["path"] = path
            captured["content"] = content
            return {"id": "sp-item", "webUrl": "https://example.invalid/d"}

    monkeypatch.setattr(sharepoint, "_get_sharepoint_graph", lambda _p: DummyGraph())

    result = await sharepoint.upload_to_site(
        sharepoint.UploadToSiteInput(
            site_id="site", drive_id="drive", uploaded_file="deck.pptx"
        )
    )

    assert result.success is True
    assert result.filename == "deck.pptx"
    assert captured["content"] == b"SLIDES"


# ---------------------------------------------------------------------------
# feature-flag resolution
# ---------------------------------------------------------------------------


def test_is_file_upload_enabled_defaults_by_transport() -> None:
    from mcp_microsoft.feature_flags import is_file_upload_enabled

    assert is_file_upload_enabled(config=AppConfig(transport="http")) is True
    assert is_file_upload_enabled(config=AppConfig(transport="stdio")) is False


def test_is_file_upload_enabled_explicit_flag_wins() -> None:
    from mcp_microsoft.feature_flags import is_file_upload_enabled

    # Explicit off in http, explicit on in stdio.
    assert (
        is_file_upload_enabled(
            config=AppConfig(transport="http", enable_file_upload=False)
        )
        is False
    )
    assert (
        is_file_upload_enabled(
            config=AppConfig(transport="stdio", enable_file_upload=True)
        )
        is True
    )


def test_resolve_upload_max_bytes_rejects_non_positive() -> None:
    from mcp_microsoft.feature_flags import resolve_upload_max_bytes

    assert resolve_upload_max_bytes(config=AppConfig(upload_max_mb=25)) == 25 * 1024 * 1024
    with pytest.raises(ValueError, match="positive"):
        resolve_upload_max_bytes(config=AppConfig(upload_max_mb=0))
