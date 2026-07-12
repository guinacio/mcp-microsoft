"""HTTP-mode-only middleware for mcp-microsoft.

Rate limiting reuses fastmcp's own ``RateLimitingMiddleware`` (see
``server._build_http_middleware``); this module holds the one piece that
needs repo-specific behavior: per-call audit logging.

Never wired up in stdio mode — see server.py's ``create_mcp_server``.
"""

from __future__ import annotations

import logging
import time
from typing import Any

from fastmcp.server.middleware import CallNext, Middleware, MiddlewareContext

_log = logging.getLogger(__name__)


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
        """Return ``(oid, preferred_username)`` from the ambient token.

        Falls back to ``"-"`` for either field (or both) when unauthenticated,
        when the claim is absent, or when the dependency itself raises —
        an audit-logging failure must never break the tool call it observes.
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

    async def on_call_tool(
        self,
        context: MiddlewareContext,
        call_next: CallNext,
    ) -> Any:
        tool_name = getattr(context.message, "name", "unknown")
        oid, username = self._caller_identity()
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
