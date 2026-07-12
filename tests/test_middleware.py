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
