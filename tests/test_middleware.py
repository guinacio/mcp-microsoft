"""Unit tests for the http-mode audit logging middleware."""

from __future__ import annotations

import logging

import pytest
from fastmcp.server.middleware import MiddlewareContext
from mcp import types as mt

from mcp_microsoft.middleware import AuditLoggingMiddleware


class _FakeAccessToken:
    """Duck-typed stand-in for fastmcp's AccessToken."""

    def __init__(self, claims: dict) -> None:
        self.claims = claims


def _tool_call_context(tool_name: str = "send_email") -> MiddlewareContext:
    params = mt.CallToolRequestParams(
        name=tool_name,
        arguments={"secret_argument": "sensitive-value"},
    )
    return MiddlewareContext(message=params, method="tools/call", type="request")


@pytest.mark.asyncio
async def test_audit_middleware_logs_success_with_identity(
    monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    import fastmcp.server.dependencies as deps

    token = _FakeAccessToken({"oid": "user-oid-1", "preferred_username": "alice@example.com"})
    monkeypatch.setattr(deps, "get_access_token", lambda: token)

    middleware = AuditLoggingMiddleware()

    async def call_next(_context: MiddlewareContext) -> str:
        return "ok"

    caplog.set_level(logging.INFO, logger="mcp_microsoft.middleware")
    result = await middleware.on_call_tool(_tool_call_context("send_email"), call_next)

    assert result == "ok"
    assert len(caplog.records) == 1
    message = caplog.records[0].getMessage()
    assert "tool=send_email" in message
    assert "oid=user-oid-1" in message
    assert "user=alice@example.com" in message
    assert "outcome=success" in message
    # Never logs tool arguments, results, or the raw token.
    assert "secret_argument" not in message
    assert "sensitive-value" not in message


@pytest.mark.asyncio
async def test_audit_middleware_logs_error_outcome_and_reraises(
    monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    import fastmcp.server.dependencies as deps

    monkeypatch.setattr(deps, "get_access_token", lambda: None)

    middleware = AuditLoggingMiddleware()

    async def call_next(_context: MiddlewareContext) -> str:
        raise RuntimeError("boom")

    caplog.set_level(logging.INFO, logger="mcp_microsoft.middleware")
    with pytest.raises(RuntimeError, match="boom"):
        await middleware.on_call_tool(_tool_call_context("delete_email"), call_next)

    message = caplog.records[-1].getMessage()
    assert "tool=delete_email" in message
    assert "oid=-" in message
    assert "user=-" in message
    assert "outcome=error" in message


@pytest.mark.asyncio
async def test_audit_middleware_never_raises_on_identity_lookup_failure(
    monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    """A broken get_access_token() must not break the tool call it's auditing."""
    import fastmcp.server.dependencies as deps

    def _boom() -> None:
        raise RuntimeError("no request context")

    monkeypatch.setattr(deps, "get_access_token", _boom)

    middleware = AuditLoggingMiddleware()

    async def call_next(_context: MiddlewareContext) -> str:
        return "ok"

    caplog.set_level(logging.INFO, logger="mcp_microsoft.middleware")
    result = await middleware.on_call_tool(_tool_call_context(), call_next)

    assert result == "ok"
    assert "oid=-" in caplog.records[0].getMessage()


def test_caller_identity_missing_claims_degrades_gracefully(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import fastmcp.server.dependencies as deps

    token = _FakeAccessToken({})  # no oid / preferred_username
    monkeypatch.setattr(deps, "get_access_token", lambda: token)

    assert AuditLoggingMiddleware._caller_identity() == ("-", "-")


def test_caller_identity_no_ambient_token(monkeypatch: pytest.MonkeyPatch) -> None:
    import fastmcp.server.dependencies as deps

    monkeypatch.setattr(deps, "get_access_token", lambda: None)

    assert AuditLoggingMiddleware._caller_identity() == ("-", "-")


# --------------------------------------------------------------------------
# UserRateLimitMiddleware — bounded per-user token-bucket rate limiter
# --------------------------------------------------------------------------


def _set_token(monkeypatch: pytest.MonkeyPatch, claims: dict | None) -> None:
    """Point the ambient get_access_token at a token with *claims* (or None)."""
    import fastmcp.server.dependencies as deps

    token = _FakeAccessToken(claims) if claims is not None else None
    monkeypatch.setattr(deps, "get_access_token", lambda: token)


def _fixed_clock(monkeypatch: pytest.MonkeyPatch, holder: list[float]) -> None:
    """Make ``time.monotonic()`` return holder[0] (mutate it to advance time)."""
    import mcp_microsoft.middleware as mw

    monkeypatch.setattr(mw.time, "monotonic", lambda: holder[0])


async def _call_next_ok(_context: MiddlewareContext) -> str:
    return "ok"


def _req_context() -> MiddlewareContext:
    params = mt.CallToolRequestParams(name="send_email", arguments={})
    return MiddlewareContext(message=params, method="tools/call", type="request")


def test_rate_limit_key_composition_and_fallbacks(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from mcp_microsoft.middleware import _UNAUTHENTICATED_KEY, _rate_limit_key

    # tid + oid -> composite key.
    _set_token(monkeypatch, {"tid": "tenant-1", "oid": "user-1"})
    assert _rate_limit_key() == "tenant-1:user-1"

    # oid only -> oid alone.
    _set_token(monkeypatch, {"oid": "user-1"})
    assert _rate_limit_key() == "user-1"

    # neither tid nor oid, but sub present -> sub.
    _set_token(monkeypatch, {"sub": "subject-9"})
    assert _rate_limit_key() == "subject-9"

    # tid without oid falls through to sub (composite needs both).
    _set_token(monkeypatch, {"tid": "tenant-1", "sub": "subject-9"})
    assert _rate_limit_key() == "subject-9"

    # no identifying claim -> shared unauthenticated bucket.
    _set_token(monkeypatch, {})
    assert _rate_limit_key() == _UNAUTHENTICATED_KEY

    # no ambient token -> shared unauthenticated bucket.
    _set_token(monkeypatch, None)
    assert _rate_limit_key() == _UNAUTHENTICATED_KEY


def test_rate_limit_key_never_raises_on_broken_dependency(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import fastmcp.server.dependencies as deps

    from mcp_microsoft.middleware import _UNAUTHENTICATED_KEY, _rate_limit_key

    def _boom() -> object:
        raise RuntimeError("no request context")

    monkeypatch.setattr(deps, "get_access_token", _boom)
    assert _rate_limit_key() == _UNAUTHENTICATED_KEY


@pytest.mark.asyncio
async def test_rate_limit_per_key_isolation_and_over_limit_exception(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """One key can be throttled without affecting another; over-limit raises
    fastmcp's RateLimitError (McpError, code -32000) unchanged."""
    from mcp import McpError
    from fastmcp.server.middleware.rate_limiting import RateLimitError

    from mcp_microsoft.middleware import UserRateLimitMiddleware

    clock = [1000.0]
    _fixed_clock(monkeypatch, clock)  # frozen time -> no refill between calls

    # rps=1 -> burst_capacity=2: two calls allowed, third throttled.
    mw = UserRateLimitMiddleware(max_requests_per_second=1)

    _set_token(monkeypatch, {"tid": "t", "oid": "A"})
    assert await mw.on_request(_req_context(), _call_next_ok) == "ok"
    assert await mw.on_request(_req_context(), _call_next_ok) == "ok"

    with pytest.raises(RateLimitError) as exc_info:
        await mw.on_request(_req_context(), _call_next_ok)
    # Exact client-visible behavior: McpError subclass, JSON-RPC code -32000.
    assert isinstance(exc_info.value, McpError)
    assert exc_info.value.error.code == -32000

    # A different key is unaffected (independent bucket).
    _set_token(monkeypatch, {"tid": "t", "oid": "B"})
    assert await mw.on_request(_req_context(), _call_next_ok) == "ok"


@pytest.mark.asyncio
async def test_rate_limit_refills_over_time(monkeypatch: pytest.MonkeyPatch) -> None:
    from fastmcp.server.middleware.rate_limiting import RateLimitError

    from mcp_microsoft.middleware import UserRateLimitMiddleware

    clock = [1000.0]
    _fixed_clock(monkeypatch, clock)
    mw = UserRateLimitMiddleware(max_requests_per_second=1)  # burst 2
    _set_token(monkeypatch, {"tid": "t", "oid": "A"})

    assert await mw.on_request(_req_context(), _call_next_ok) == "ok"
    assert await mw.on_request(_req_context(), _call_next_ok) == "ok"
    with pytest.raises(RateLimitError):
        await mw.on_request(_req_context(), _call_next_ok)

    # Advance 1s at 1 token/s -> one more request allowed.
    clock[0] += 1.0
    assert await mw.on_request(_req_context(), _call_next_ok) == "ok"


@pytest.mark.asyncio
async def test_rate_limit_lru_cap_evicts_oldest(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import mcp_microsoft.middleware as mw_mod
    from mcp_microsoft.middleware import UserRateLimitMiddleware

    clock = [1000.0]
    _fixed_clock(monkeypatch, clock)  # frozen -> idle TTL never triggers here
    monkeypatch.setattr(mw_mod, "_LIMITER_CAP", 2)

    mw = UserRateLimitMiddleware(max_requests_per_second=100)

    for oid in ("A", "B", "C"):
        _set_token(monkeypatch, {"tid": "t", "oid": oid})
        await mw.on_request(_req_context(), _call_next_ok)

    # Cap is 2; least-recently-used key ("t:A") was evicted, newest two remain.
    assert set(mw._buckets.keys()) == {"t:B", "t:C"}


@pytest.mark.asyncio
async def test_rate_limit_prunes_idle_buckets(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import mcp_microsoft.middleware as mw_mod
    from mcp_microsoft.middleware import UserRateLimitMiddleware

    clock = [1000.0]
    _fixed_clock(monkeypatch, clock)
    mw = UserRateLimitMiddleware(max_requests_per_second=100)

    # Touch key A at t=1000.
    _set_token(monkeypatch, {"tid": "t", "oid": "A"})
    await mw.on_request(_req_context(), _call_next_ok)
    assert "t:A" in mw._buckets

    # Advance past the idle TTL, then touch key B: the lazy sweep drops A.
    clock[0] += mw_mod._LIMITER_IDLE_TTL + 1.0
    _set_token(monkeypatch, {"tid": "t", "oid": "B"})
    await mw.on_request(_req_context(), _call_next_ok)

    assert "t:A" not in mw._buckets
    assert "t:B" in mw._buckets
