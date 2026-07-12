"""HTTP-mode-only middleware for mcp-microsoft.

Holds the repo-specific pieces of the http-mode middleware stack: a bounded
per-user token-bucket rate limiter, per-call audit logging, and metrics
recording.

Never wired up in stdio mode — see server.py's ``create_mcp_server``.
"""

from __future__ import annotations

import logging
import time
from collections import OrderedDict
from typing import Any

from fastmcp.exceptions import NotFoundError
from fastmcp.server.middleware import CallNext, Middleware, MiddlewareContext
from fastmcp.server.middleware.rate_limiting import RateLimitError

_log = logging.getLogger(__name__)

# Shared rate-limit bucket key for any request whose caller identity cannot be
# resolved. Kept separate from every authenticated user's per-identity bucket so
# an unauthenticated caller can only ever starve other unauthenticated callers.
_UNAUTHENTICATED_KEY = "unauthenticated"

# Bounds for UserRateLimitMiddleware's per-key bucket store. fastmcp's own
# RateLimitingMiddleware backs its per-client limiters with an unbounded
# defaultdict that never evicts (entries live for the whole process lifetime);
# these caps keep our store bounded in a long-running multi-user server.
_LIMITER_CAP = 10_000  # max distinct keys retained (LRU eviction beyond this)
_LIMITER_IDLE_TTL = 900.0  # seconds a bucket may sit idle before it is pruned


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


def _caller_tid_oid_sub() -> tuple[str | None, str | None, str | None]:
    """Return ``(tid, oid, sub)`` from the ambient bearer token, else ``None``s.

    Sibling of :func:`_caller_identity` (which returns ``(oid, username)``);
    added for per-user rate-limit keying without changing that helper's
    signature. Like it, this never raises — an absent claim, no ambient token,
    or a raising dependency all collapse to ``None`` for the affected field so
    keying a rate-limit bucket can never break the request it is limiting. The
    bearer token itself is never read into the return value.
    """
    try:
        from fastmcp.server.dependencies import get_access_token

        token = get_access_token()
    except Exception:
        return None, None, None

    if token is None:
        return None, None, None

    claims = getattr(token, "claims", None) or {}
    tid = claims.get("tid")
    oid = claims.get("oid")
    sub = claims.get("sub")
    return (
        str(tid) if tid else None,
        str(oid) if oid else None,
        str(sub) if sub else None,
    )


def _rate_limit_key() -> str:
    """Return the per-user rate-limit bucket key for the current request.

    Prefers ``f"{tid}:{oid}"`` (tenant + object id — globally unique per user),
    falling back to ``oid`` alone, then ``sub``, then the shared
    ``"unauthenticated"`` bucket. Never raises — see :func:`_caller_tid_oid_sub`.
    """
    tid, oid, sub = _caller_tid_oid_sub()
    if tid and oid:
        return f"{tid}:{oid}"
    if oid:
        return oid
    if sub:
        return sub
    return _UNAUTHENTICATED_KEY


class _TokenBucket:
    """A single token bucket with a last-seen stamp for idle eviction.

    Mirrors fastmcp ``TokenBucketRateLimiter`` semantics (capacity tokens,
    refilled at ``refill_rate`` tokens/second, consume one per request) but is
    driven by an externally supplied monotonic clock and carries ``last_seen``
    so the owning store can prune idle buckets. The refill/consume arithmetic is
    synchronous and does no ``await``, so it is atomic on the single event loop.
    """

    __slots__ = ("tokens", "last_refill", "last_seen")

    def __init__(self, capacity: float, now: float) -> None:
        self.tokens = float(capacity)
        self.last_refill = now
        self.last_seen = now


class UserRateLimitMiddleware(Middleware):
    """Bounded per-user token-bucket rate limiter for http (multi-user) mode.

    Replaces fastmcp's ``RateLimitingMiddleware`` to fix two problems:

    * **Keying.** fastmcp keys every request under a single literal ``"global"``
      bucket unless a ``get_client_id`` callable is supplied, letting one user
      throttle everyone. We key on the caller's validated Entra identity
      (:func:`_rate_limit_key`: ``tid:oid``, falling back to ``oid`` / ``sub`` /
      the shared ``"unauthenticated"`` bucket).
    * **Unbounded state.** fastmcp's per-client ``defaultdict`` of limiters never
      evicts, so entries accumulate for the whole process lifetime. We store the
      buckets in an ``OrderedDict`` capped at ``_LIMITER_CAP`` (LRU eviction) and
      additionally prune any bucket idle longer than ``_LIMITER_IDLE_TTL`` on a
      lazy sweep from the LRU head.

    User-facing semantics are unchanged from fastmcp: same token-bucket algorithm
    (``burst_capacity`` defaults to ``2 * max_requests_per_second`` as fastmcp's
    does), applied in ``on_request`` (so it covers all MCP requests, not just
    tool calls), and over-limit raises the same ``RateLimitError`` (an
    ``McpError`` with JSON-RPC code ``-32000``) so client-visible behavior is
    identical.
    """

    def __init__(
        self,
        max_requests_per_second: float = 10.0,
        burst_capacity: int | None = None,
    ) -> None:
        self.max_requests_per_second = max_requests_per_second
        self.burst_capacity = (
            burst_capacity
            if burst_capacity is not None
            else int(max_requests_per_second * 2)
        )
        self._buckets: OrderedDict[str, _TokenBucket] = OrderedDict()

    def _prune_idle(self, now: float) -> None:
        """Evict buckets idle longer than the TTL, sweeping from the LRU head.

        Buckets are ordered by last access (``move_to_end`` on every hit), so the
        head is the least-recently-seen: once it is within the TTL, every later
        bucket is too, and the sweep stops. O(1) amortized per request.
        """
        while self._buckets:
            _key, bucket = next(iter(self._buckets.items()))
            if now - bucket.last_seen <= _LIMITER_IDLE_TTL:
                break
            self._buckets.popitem(last=False)

    def _allow(self, key: str, now: float) -> bool:
        """Refill and consume one token for *key*; True if a token was available.

        Purely synchronous (no ``await``) so the read-modify-write is atomic on
        the event loop.
        """
        bucket = self._buckets.get(key)
        if bucket is None:
            bucket = _TokenBucket(self.burst_capacity, now)
            self._buckets[key] = bucket
        else:
            self._buckets.move_to_end(key)
            elapsed = now - bucket.last_refill
            bucket.tokens = min(
                self.burst_capacity,
                bucket.tokens + elapsed * self.max_requests_per_second,
            )
            bucket.last_refill = now
        bucket.last_seen = now

        # Enforce the LRU cap. The just-touched key sits at the tail, so only
        # genuinely least-recently-used keys are evicted.
        while len(self._buckets) > _LIMITER_CAP:
            self._buckets.popitem(last=False)

        if bucket.tokens >= 1:
            bucket.tokens -= 1
            return True
        return False

    async def on_request(
        self,
        context: MiddlewareContext,
        call_next: CallNext,
    ) -> Any:
        now = time.monotonic()
        self._prune_idle(now)
        key = _rate_limit_key()
        if not self._allow(key, now):
            raise RateLimitError(f"Rate limit exceeded for client: {key}")
        return await call_next(context)


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
