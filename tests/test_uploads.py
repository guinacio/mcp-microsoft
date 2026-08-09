"""Tests for the context-free file-upload provider (uploads.ScopedFileUpload).

Covers per-caller scoping, the bounded store (file/byte quotas, whole-scope LRU
eviction, idle-TTL prune, name-length cap), the resolve_uploaded_file accessor,
and the Graph upload-tool integration (uploaded_file resolves bytes into the
existing PUT path; mutual-exclusion and feature-off errors).
"""

from __future__ import annotations

import base64

import pytest
from fastmcp.exceptions import ToolError

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
    def __init__(self, oid: str | None, sub: str | None = None) -> None:
        self.token = "assertion"
        claims: dict[str, str] = {}
        if oid:
            claims["oid"] = oid
        if sub:
            claims["sub"] = sub
        self.claims = claims


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


def _set_ambient_claims(
    monkeypatch: pytest.MonkeyPatch,
    *,
    oid: str | None = None,
    sub: str | None = None,
    token: bool = True,
) -> None:
    """Install an ambient token with the given claims (or no token at all)."""
    import fastmcp.server.dependencies as deps

    tok = _FakeAccessToken(oid, sub) if token else None
    monkeypatch.setattr(deps, "get_access_token", lambda: tok)


# ---------------------------------------------------------------------------
# quota enforcement
# ---------------------------------------------------------------------------


def test_on_store_enforces_per_scope_file_count(monkeypatch: pytest.MonkeyPatch) -> None:
    _force_stdio(monkeypatch)
    provider = ScopedFileUpload(max_files_per_scope=2)
    ctx = _FakeCtx()

    provider.on_store([_file("a.bin", b"a"), _file("b.bin", b"b")], ctx)
    with pytest.raises(ToolError, match="per-user limit of 2"):
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
    with pytest.raises(ToolError, match="per-user limit"):
        provider.on_store([_file("b.bin", b"y" * 8)], ctx)


def test_on_store_rejects_overlong_name(monkeypatch: pytest.MonkeyPatch) -> None:
    _force_stdio(monkeypatch)
    provider = ScopedFileUpload(max_name_length=10)
    ctx = _FakeCtx()

    with pytest.raises(ToolError, match="name is too long"):
        provider.on_store([_file("x" * 11, b"data")], ctx)


def test_on_store_rejects_atomically(monkeypatch: pytest.MonkeyPatch) -> None:
    """An over-quota batch stores none of its files (validation precedes commit)."""
    _force_stdio(monkeypatch)
    provider = ScopedFileUpload(max_files_per_scope=1)
    ctx = _FakeCtx()

    with pytest.raises(ToolError):
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
    with pytest.raises(ToolError, match="not found"):
        provider.on_read("secret.txt", ctx)


def test_scope_key_http_falls_back_oid_then_sub_then_refuses(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """http mode: oid wins, else sub, else REFUSE (never a shared bucket)."""
    _force_http(monkeypatch)
    provider = ScopedFileUpload()

    # oid present -> oid key.
    _set_ambient_claims(monkeypatch, oid="o1", sub="s1")
    assert provider._get_scope_key(_FakeCtx(session_id=None)) == "oid:o1"

    # no oid but sub present -> sub key (a session id must NOT be used in http).
    _set_ambient_claims(monkeypatch, sub="s2")
    assert provider._get_scope_key(_FakeCtx(session_id="ignored")) == "sub:s2"

    # neither oid nor sub -> refuse rather than share a bucket.
    _set_ambient_claims(monkeypatch)  # token present, no identity claims
    with pytest.raises(ToolError, match="stable user identity"):
        provider._get_scope_key(_FakeCtx(session_id="ignored"))

    # no ambient token at all -> also refuse.
    _set_ambient_claims(monkeypatch, token=False)
    with pytest.raises(ToolError, match="stable user identity"):
        provider._get_scope_key(_FakeCtx(session_id="ignored"))


def test_scope_key_stdio_uses_session_then_default(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """stdio mode keeps session-id scoping with a shared default fallback."""
    _force_stdio(monkeypatch)
    provider = ScopedFileUpload()

    assert provider._get_scope_key(_FakeCtx(session_id="s9")) == "sid:s9"
    assert provider._get_scope_key(_FakeCtx(session_id=None)) == "__default__"


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

    with pytest.raises(ToolError, match="present.txt"):
        resolve_uploaded_file("absent.txt")


def test_resolve_uploaded_file_feature_off_raises() -> None:
    set_upload_provider(None)
    with pytest.raises(ToolError, match="MCP_ENABLE_FILE_UPLOAD"):
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

    with pytest.raises(ToolError, match="cannot be combined"):
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

    with pytest.raises(ToolError, match="not enabled"):
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


def test_resolve_upload_global_budget_bytes_rejects_non_positive() -> None:
    from mcp_microsoft.feature_flags import resolve_upload_global_budget_bytes

    assert (
        resolve_upload_global_budget_bytes(config=AppConfig(upload_global_budget_mb=2))
        == 2 * 1024 * 1024
    )
    with pytest.raises(ValueError, match="positive"):
        resolve_upload_global_budget_bytes(config=AppConfig(upload_global_budget_mb=0))


# ---------------------------------------------------------------------------
# F5 — global encoded-byte budget across all scopes
# ---------------------------------------------------------------------------


def test_global_budget_rejects_across_scopes(monkeypatch: pytest.MonkeyPatch) -> None:
    """The global budget counts encoded bytes across ALL scopes, not per-scope."""
    _force_http(monkeypatch)
    # Budget of ~600 encoded bytes; each 300-byte file encodes to ~400 base64.
    provider = ScopedFileUpload(global_budget_bytes=600)

    _set_ambient_oid(monkeypatch, "user-A")
    provider.on_store([_file("a.bin", b"x" * 300)], _FakeCtx(session_id=None))
    encoded_a = provider._global_encoded_bytes
    assert encoded_a > 300  # base64 expands ~4/3

    # A DIFFERENT user, well under their own per-user quota, is still rejected
    # because the global encoded budget is already nearly exhausted.
    _set_ambient_oid(monkeypatch, "user-B")
    with pytest.raises(ToolError, match="server upload storage is full"):
        provider.on_store([_file("b.bin", b"y" * 300)], _FakeCtx(session_id=None))

    # The rejected store left the global total untouched (atomic).
    assert provider._global_encoded_bytes == encoded_a


def test_global_budget_accounting_adjusts_on_overwrite_and_evict(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Global total follows overwrites (in place) and whole-scope eviction."""
    _force_http(monkeypatch)
    provider = ScopedFileUpload(max_scopes=1, global_budget_bytes=10 * 1024 * 1024)

    _set_ambient_oid(monkeypatch, "u1")
    provider.on_store([_file("f.bin", b"a" * 100)], _FakeCtx(session_id=None))
    after_first = provider._global_encoded_bytes

    # Overwrite the same name with a bigger payload -> total reflects the delta,
    # not the sum of both versions.
    provider.on_store([_file("f.bin", b"a" * 400)], _FakeCtx(session_id=None))
    assert provider._global_encoded_bytes > after_first
    only_scope_bytes = provider._global_encoded_bytes

    # A second distinct scope evicts u1 (max_scopes=1); the global total drops the
    # evicted scope's contribution and reflects only the survivor.
    _set_ambient_oid(monkeypatch, "u2")
    provider.on_store([_file("g.bin", b"z" * 50)], _FakeCtx(session_id=None))
    assert provider.scopes_evicted == 1
    assert provider._global_encoded_bytes < only_scope_bytes


# ---------------------------------------------------------------------------
# F2 — read_file caps returned text content
# ---------------------------------------------------------------------------


def test_read_file_caps_large_text_content(monkeypatch: pytest.MonkeyPatch) -> None:
    from mcp_microsoft.uploads import _READ_CONTENT_CAP

    _force_stdio(monkeypatch)
    provider = ScopedFileUpload(
        max_file_size=10 * 1024 * 1024, max_bytes_per_scope=10 * 1024 * 1024
    )
    ctx = _FakeCtx()

    big = ("A" * (_READ_CONTENT_CAP + 5000)).encode("ascii")
    provider.on_store([_file("big.txt", big, content_type="text/plain")], ctx)

    result = provider.on_read("big.txt", ctx)
    assert result["truncated"] is True
    assert len(result["content"]) == _READ_CONTENT_CAP
    assert "note" in result and "upload_file" in result["note"]
    # Full size is still reported honestly.
    assert result["size"] == len(big)


def test_read_file_small_text_unchanged(monkeypatch: pytest.MonkeyPatch) -> None:
    _force_stdio(monkeypatch)
    provider = ScopedFileUpload()
    ctx = _FakeCtx()

    provider.on_store([_file("small.txt", b"hello world", content_type="text/plain")], ctx)
    result = provider.on_read("small.txt", ctx)
    assert result["content"] == "hello world"
    assert "truncated" not in result
    assert "note" not in result


# ---------------------------------------------------------------------------
# F6 — identity fallback hardening (sub isolation; refuse when no identity)
# ---------------------------------------------------------------------------


def test_two_users_without_oid_isolated_by_sub(monkeypatch: pytest.MonkeyPatch) -> None:
    """Two authenticated callers WITHOUT oid but distinct sub stay isolated."""
    _force_http(monkeypatch)
    provider = ScopedFileUpload()
    ctx = _FakeCtx(session_id=None)

    _set_ambient_claims(monkeypatch, sub="sub-A")
    provider.on_store([_file("a-secret.txt", b"A")], ctx)
    assert {s["name"] for s in provider.on_list(ctx)} == {"a-secret.txt"}

    _set_ambient_claims(monkeypatch, sub="sub-B")
    assert provider.on_list(ctx) == []  # cannot see sub-A's file
    with pytest.raises(ToolError, match="not found"):
        provider.on_read("a-secret.txt", ctx)


def test_http_no_identity_refuses_storage(monkeypatch: pytest.MonkeyPatch) -> None:
    """http mode with no oid/sub (or no token) refuses every storage op."""
    _force_http(monkeypatch)
    provider = ScopedFileUpload()
    ctx = _FakeCtx(session_id="ignored")

    _set_ambient_claims(monkeypatch, token=False)  # no ambient token at all
    with pytest.raises(ToolError, match="stable user identity"):
        provider.on_store([_file("x.bin", b"x")], ctx)
    with pytest.raises(ToolError, match="stable user identity"):
        provider.on_list(ctx)


# ---------------------------------------------------------------------------
# F1 — thread-safety under real concurrent workers
# ---------------------------------------------------------------------------


def test_concurrent_store_list_prune_is_consistent(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Hammer on_store/on_list/_prune from many threads: no errors, consistent state."""
    import threading
    from concurrent.futures import ThreadPoolExecutor

    _force_stdio(monkeypatch)  # stdio scoping is thread-agnostic (session/default)
    provider = ScopedFileUpload(
        max_files_per_scope=1000,
        max_bytes_per_scope=50 * 1024 * 1024,
        global_budget_bytes=50 * 1024 * 1024,
        max_scopes=64,
    )

    errors: list[BaseException] = []
    barrier = threading.Barrier(8)

    def worker(wid: int) -> None:
        barrier.wait()
        try:
            for i in range(50):
                # Mix of a shared scope (session None -> __default__) and distinct
                # per-worker session scopes.
                shared = _FakeCtx(session_id=None)
                mine = _FakeCtx(session_id=f"w{wid}")
                provider.on_store([_file(f"w{wid}-f{i}.bin", b"x" * 8)], mine)
                provider.on_store([_file(f"shared-{wid}-{i}.bin", b"y" * 4)], shared)
                provider.on_list(mine)
                provider.on_list(shared)
                provider._clock()
        except BaseException as exc:  # noqa: BLE001 - surface any thread failure
            errors.append(exc)

    with ThreadPoolExecutor(max_workers=8) as pool:
        list(pool.map(worker, range(8)))

    assert errors == [], f"threads raised: {errors!r}"

    # Accounting is internally consistent: the global total equals the sum of the
    # surviving scopes' encoded_bytes, and each scope's encoded_bytes equals the
    # sum of its files' encoded_size.
    total = 0
    for entry in provider._store.values():
        scope_encoded = sum(fe["encoded_size"] for fe in entry["files"].values())
        assert entry["encoded_bytes"] == scope_encoded
        total += entry["encoded_bytes"]
    assert provider._global_encoded_bytes == total


# ---------------------------------------------------------------------------
# F3 — actionable upload errors survive http-mode masking (in-memory client)
# ---------------------------------------------------------------------------


def _error_text(result: object) -> str:
    """Flatten a CallToolResult's error payload to a single searchable string."""
    parts = "".join(
        str(getattr(block, "text", block)) for block in (getattr(result, "content", None) or [])
    )
    return parts + str(getattr(result, "structured_content", "") or "")


@pytest.mark.asyncio
async def test_over_quota_message_survives_masking(monkeypatch: pytest.MonkeyPatch) -> None:
    """A ToolError from on_store reaches the client unmasked on a masking server."""
    from fastmcp import Client, FastMCP
    from fastmcp.server.providers.addressing import hashed_backend_name

    _force_http(monkeypatch)
    _set_ambient_oid(monkeypatch, "masking-user")

    mcp = FastMCP("mask-test", mask_error_details=True)
    assert mcp._mask_error_details is True
    provider = ScopedFileUpload(max_files_per_scope=1)
    mcp.add_provider(provider)

    # store_files is not model-visible; address it by its hashed backend name.
    store_name = hashed_backend_name(provider.name, "store_files")

    async with Client(mcp) as client:
        await client.call_tool(store_name, {"files": [_file("a.bin", b"a")]})
        result = await client.call_tool(
            store_name, {"files": [_file("b.bin", b"b")]}, raise_on_error=False
        )
    assert result.is_error is True
    # The actionable limit message is visible, not a generic masked string.
    assert "per-user limit of 1" in _error_text(result)


@pytest.mark.asyncio
async def test_mutual_exclusion_message_survives_masking(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """upload_file's mutual-exclusion ToolError reaches the client unmasked."""
    from fastmcp import Client, FastMCP

    from mcp_microsoft.tools import onedrive

    monkeypatch.setattr(onedrive, "get_app_config", lambda: AppConfig(transport="http"))
    _force_http(monkeypatch)

    mcp = FastMCP("mask-test", mask_error_details=True)
    onedrive.register(mcp)

    async with Client(mcp) as client:
        result = await client.call_tool(
            "upload_file",
            {
                "params": {
                    "uploaded_file": "r.txt",
                    "content_base64": base64.b64encode(b"x").decode("ascii"),
                }
            },
            raise_on_error=False,
        )
    assert result.is_error is True
    assert "cannot be combined" in _error_text(result)


# ---------------------------------------------------------------------------
# F7 — temp-file cleanup on a write failure (no orphan left behind)
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_upload_file_write_failure_cleans_up_temp(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import tempfile as _tempfile

    from mcp_microsoft.tools import onedrive

    monkeypatch.setattr(onedrive, "get_app_config", lambda: AppConfig(transport="http"))
    _force_http(monkeypatch)
    _set_ambient_oid(monkeypatch, "uploader")

    provider = ScopedFileUpload()
    provider.on_store([_file("report.txt", b"hello")], _FakeCtx(session_id=None))
    set_upload_provider(provider)

    created: list[str] = []
    real_ntf = _tempfile.NamedTemporaryFile

    class _BoomFile:
        def __init__(self, inner: object) -> None:
            self._inner = inner

        def __enter__(self) -> "_BoomFile":
            self._inner.__enter__()
            created.append(self.name)
            return self

        def __exit__(self, *exc: object) -> object:
            return self._inner.__exit__(*exc)

        @property
        def name(self) -> str:
            return self._inner.name

        def write(self, _data: bytes) -> int:
            raise OSError("disk full")

    def _fake_ntf(*args: object, **kwargs: object) -> _BoomFile:
        return _BoomFile(real_ntf(*args, **kwargs))

    monkeypatch.setattr(_tempfile, "NamedTemporaryFile", _fake_ntf)

    with pytest.raises(OSError, match="disk full"):
        await onedrive.upload_file(onedrive.UploadFileInput(uploaded_file="report.txt"))

    # The temp file was created but the write failed; the finally block unlinked it.
    assert created, "expected a temp file to have been created"
    import os

    for path in created:
        assert not os.path.exists(path), f"orphan temp file left behind: {path}"
