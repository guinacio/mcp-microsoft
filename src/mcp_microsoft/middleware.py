"""HTTP-mode-only middleware for mcp-microsoft.

Rate limiting reuses fastmcp's own ``RateLimitingMiddleware`` (see
``server._build_http_middleware``); this module holds the two pieces that
need repo-specific behavior: per-call audit logging and metrics recording.

Never wired up in stdio mode — see server.py's ``create_mcp_server``.
"""

from __future__ import annotations

import logging
import time
from typing import Any

from fastmcp.exceptions import NotFoundError
from fastmcp.server.middleware import CallNext, Middleware, MiddlewareContext

_log = logging.getLogger(__name__)


def _caller_identity() -> tuple[str, str]:
    """Return ``(oid, preferred_username)`` from the ambient bearer token.

    Shared by :class:`AuditLoggingMiddleware` and :class:`MetricsMiddleware`
    so both attribute a call to the same identity. Falls back to ``"-"`` for
    either field (or both) when unauthenticated, when the claim is absent, or
    when the dependency itself raises — reading identity for observability must
    never break the tool call it observes. The bearer token itself is never
    read into the return value.
    """
    try:
        from fastmcp.server.dependencies import get_access_token

        token = get_access_token()
    except Exception:
        return "-", "-"

    if token is None:
        return "-", "-"

    claims = getattr(token, "claims", None) or {}
    oid = claims.get("oid") or "-"
    username = claims.get("preferred_username") or "-"
    return str(oid), str(username)


class AuditLoggingMiddleware(Middleware):
    """Logs a one-line audit record for every tool call in http (multi-user) mode.

    Records the tool name, caller identity (the ``oid`` and
    ``preferred_username`` claims from the ambient bearer token — never the
    token itself), duration, and outcome (success/error). Arguments and
    result payloads are deliberately never logged: they can carry message
    bodies, attachment content, or other user data.
    """

    def __init__(self, logger: logging.Logger | None = None) -> None:
        self._logger = logger or _log

    @staticmethod
    def _caller_identity() -> tuple[str, str]:
        """Backwards-compatible alias for the module-level helper."""
        return _caller_identity()

    async def on_call_tool(
        self,
        context: MiddlewareContext,
        call_next: CallNext,
    ) -> Any:
        tool_name = getattr(context.message, "name", "unknown")
        oid, username = _caller_identity()
        start = time.perf_counter()

        try:
            result = await call_next(context)
        except Exception:
            duration_ms = (time.perf_counter() - start) * 1000
            self._logger.info(
                "tool_call tool=%s oid=%s user=%s duration_ms=%.1f outcome=error",
                tool_name,
                oid,
                username,
                duration_ms,
            )
            raise

        duration_ms = (time.perf_counter() - start) * 1000
        self._logger.info(
            "tool_call tool=%s oid=%s user=%s duration_ms=%.1f outcome=success",
            tool_name,
            oid,
            username,
            duration_ms,
        )
        return result


class MetricsMiddleware(Middleware):
    """Feeds every tool call into the in-process :class:`MetricsRegistry`.

    Sibling of :class:`AuditLoggingMiddleware`, registered after it in the
    http-mode stack (so it wraps the tool call most tightly and times the tool
    itself). Records the call in both the success and exception paths, timing
    the call in each; the exception is always re-raised unchanged.

    Recording must never affect the observed call: the registry write is
    wrapped so any failure inside it is swallowed (and debug-logged) rather
    than propagated. Identity comes from the same helper the audit middleware
    uses, so both attribute a call to the same ``oid``/``username``.
    """

    def __init__(self, registry: Any | None = None, logger: logging.Logger | None = None) -> None:
        # A ``None`` registry means "resolve the process-wide singleton lazily
        # on each call", which keeps the middleware in step with the registry
        # the stats routes read even across a reset. Tests may inject one.
        self._registry = registry
        self._logger = logger or _log

    def _record(
        self, tool: str, oid: str, username: str, duration_ms: float, ok: bool
    ) -> None:
        try:
            from mcp_microsoft.metrics import get_metrics_registry

            registry = self._registry or get_metrics_registry()
            registry.record(tool, oid, username, duration_ms, ok)
        except Exception:  # pragma: no cover - defensive; must never break a call
            self._logger.debug("metrics recording failed", exc_info=True)

    def _record_unknown(self) -> None:
        """Bump the aggregate unknown-tool counter, swallowing any failure."""
        try:
            from mcp_microsoft.metrics import get_metrics_registry

            registry = self._registry or get_metrics_registry()
            registry.record_unknown_tool()
        except Exception:  # pragma: no cover - defensive; must never break a call
            self._logger.debug("metrics recording failed", exc_info=True)

    async def on_call_tool(
        self,
        context: MiddlewareContext,
        call_next: CallNext,
    ) -> Any:
        tool_name = getattr(context.message, "name", "unknown")
        oid, username = _caller_identity()
        start = time.perf_counter()
        try:
            result = await call_next(context)
        except NotFoundError:
            # fastmcp resolves the tool name INSIDE call_next (after the
            # middleware chain has already started), so a tools/call for a name
            # that does not exist surfaces here as a NotFoundError carrying the
            # attacker-controlled name. Recording it under that name would let
            # any authenticated caller grow the per-tool dict without bound,
            # poison the error/traffic counters, and explode Prometheus
            # tool-label cardinality. Count it in a single aggregate bucket
            # instead and re-raise unchanged. (NotFoundError is a bare
            # Exception subclass, distinct from fastmcp's ValidationError -- so
            # argument-validation failures on a REAL tool fall through to the
            # branch below and are still recorded under their real name.)
            self._record_unknown()
            raise
        except Exception:
            duration_ms = (time.perf_counter() - start) * 1000
            self._record(tool_name, oid, username, duration_ms, ok=False)
            raise
        duration_ms = (time.perf_counter() - start) * 1000
        self._record(tool_name, oid, username, duration_ms, ok=True)
        return result
