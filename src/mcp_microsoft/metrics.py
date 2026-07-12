"""In-process observability metrics for http (multi-user) transport.

:class:`MetricsRegistry` aggregates tool-call traffic, per-tool latency, and
per-user activity entirely in memory. It is designed for the **single-worker**
http deployment mcp-microsoft targets (see the plan doc's single-worker
constraint): there is exactly one registry per process and it is never shared
across workers.

**Concurrency invariant.** Every mutating method (:meth:`MetricsRegistry.record`)
and every reader (:meth:`MetricsRegistry.snapshot`,
:meth:`MetricsRegistry.render_prometheus`) runs to completion synchronously on
the asyncio event loop — there is no ``await`` anywhere inside an update path,
so no two updates can interleave and no lock is required. Callers on the event
loop (the metrics middleware, the stats routes) therefore see a consistent
view without any synchronization primitive. Do **not** add ``await`` inside
these methods without also introducing a lock.

Nothing here is wired up in stdio mode — the metrics middleware and the stats
routes that read this registry are registered only in http mode, and only when
``MCP_STATS_TOKEN`` is set (see ``server.py``).
"""

from __future__ import annotations

import collections
import math
import time
from datetime import datetime, timezone
from typing import Any

# Number of one-minute buckets kept in the rolling traffic timeline.
_TIMELINE_MINUTES = 60
# Per-tool latency samples retained for percentile math. Bounded so a hot tool
# can't grow this without limit; percentiles are computed over the most recent
# window of this many calls.
_DURATION_SAMPLES = 256
# Hard cap on distinct caller identities tracked at once. On overflow the
# least-recently-seen identity is evicted (and counted). Bounds memory against
# a large or hostile user population; the "top" list surfaced to operators is
# far smaller anyway.
_USER_CAP = 1000
# Cap on the number of users included in a snapshot's ``top`` list.
_USER_TOP = 100


def _iso(ts: float) -> str:
    """Return *ts* (epoch seconds) as an ISO-8601 UTC string."""
    return datetime.fromtimestamp(ts, tz=timezone.utc).isoformat()


def _percentile(sorted_values: list[float], pct: float) -> float:
    """Linear-interpolation percentile of an already-sorted list.

    Matches the common "linear interpolation between closest ranks" method
    (the numpy default). Returns 0.0 for an empty input.
    """
    if not sorted_values:
        return 0.0
    if len(sorted_values) == 1:
        return sorted_values[0]
    k = (len(sorted_values) - 1) * (pct / 100.0)
    lo = math.floor(k)
    hi = math.ceil(k)
    if lo == hi:
        return sorted_values[int(k)]
    return sorted_values[lo] * (hi - k) + sorted_values[hi] * (k - lo)


class MetricsRegistry:
    """In-process aggregation of tool-call metrics for one server process."""

    def __init__(self) -> None:
        self._started_at = time.time()
        self._total_calls = 0
        self._total_errors = 0
        # Rolling per-minute traffic, keyed by minute index (int(ts // 60)).
        # Pruned to the last _TIMELINE_MINUTES entries on write.
        self._minutes: dict[int, dict[str, int]] = {}
        # Per-tool aggregates. Cardinality is naturally bounded by the ~95
        # registered tools, so no cap is needed here.
        self._tools: dict[str, dict[str, Any]] = {}
        # Per-user aggregates, ordered by recency of last activity so the
        # least-recently-seen identity sits at the front for O(1) eviction.
        self._users: "collections.OrderedDict[str, dict[str, Any]]" = (
            collections.OrderedDict()
        )
        self._users_evicted = 0

    # -- write path ---------------------------------------------------------

    def _minute_bucket(self, now: float) -> dict[str, int]:
        """Return the traffic bucket for *now*, creating & pruning as needed."""
        minute = int(now // 60)
        bucket = self._minutes.get(minute)
        if bucket is None:
            bucket = {"calls": 0, "errors": 0}
            self._minutes[minute] = bucket
            cutoff = minute - (_TIMELINE_MINUTES - 1)
            if len(self._minutes) > _TIMELINE_MINUTES:
                for stale in [m for m in self._minutes if m < cutoff]:
                    del self._minutes[stale]
        return bucket

    def record(
        self,
        tool: str,
        oid: str,
        username: str,
        duration_ms: float,
        ok: bool,
    ) -> None:
        """Record a single tool call. The one write API.

        Runs synchronously on the event loop with no ``await`` — see the
        module docstring's concurrency invariant.
        """
        now = time.time()

        # Global + traffic timeline.
        self._total_calls += 1
        bucket = self._minute_bucket(now)
        bucket["calls"] += 1
        if not ok:
            self._total_errors += 1
            bucket["errors"] += 1

        # Per-tool.
        stats = self._tools.get(tool)
        if stats is None:
            stats = {
                "calls": 0,
                "errors": 0,
                "durations": collections.deque(maxlen=_DURATION_SAMPLES),
            }
            self._tools[tool] = stats
        stats["calls"] += 1
        if not ok:
            stats["errors"] += 1
        stats["durations"].append(float(duration_ms))

        # Per-user (unauthenticated / unknown collapses to the "-" key).
        key = oid or "-"
        user = self._users.get(key)
        if user is None:
            if len(self._users) >= _USER_CAP:
                # Evict the least-recently-seen identity (front of the
                # OrderedDict) and count it.
                self._users.popitem(last=False)
                self._users_evicted += 1
            user = {
                "username": username or "-",
                "calls": 0,
                "errors": 0,
                "first_seen": now,
                "last_seen": now,
            }
            self._users[key] = user
        user["username"] = username or "-"
        user["calls"] += 1
        if not ok:
            user["errors"] += 1
        user["last_seen"] = now
        # Mark most-recently-seen so eviction order tracks last_seen.
        self._users.move_to_end(key)

    # -- read path ----------------------------------------------------------

    def _minute_series(self, now: float) -> list[dict[str, Any]]:
        """Return the last _TIMELINE_MINUTES buckets, oldest first, zero-filled."""
        current = int(now // 60)
        series: list[dict[str, Any]] = []
        for offset in range(_TIMELINE_MINUTES - 1, -1, -1):
            minute = current - offset
            bucket = self._minutes.get(minute, {"calls": 0, "errors": 0})
            series.append(
                {
                    "t": _iso(minute * 60),
                    "calls": bucket["calls"],
                    "errors": bucket["errors"],
                }
            )
        return series

    def snapshot(self) -> dict[str, Any]:
        """Return a JSON-serializable point-in-time view of all metrics."""
        now = time.time()
        series = self._minute_series(now)

        def _window(entries: list[dict[str, Any]]) -> dict[str, int]:
            return {
                "calls": sum(e["calls"] for e in entries),
                "errors": sum(e["errors"] for e in entries),
            }

        tools: list[dict[str, Any]] = []
        for name, stats in self._tools.items():
            durations = sorted(stats["durations"])
            avg = sum(durations) / len(durations) if durations else 0.0
            tools.append(
                {
                    "name": name,
                    "calls": stats["calls"],
                    "errors": stats["errors"],
                    "avg_ms": round(avg, 2),
                    "p50_ms": round(_percentile(durations, 50), 2),
                    "p95_ms": round(_percentile(durations, 95), 2),
                }
            )
        tools.sort(key=lambda t: t["calls"], reverse=True)

        users_sorted = sorted(
            self._users.items(),
            key=lambda kv: kv[1]["last_seen"],
            reverse=True,
        )
        top_users = [
            {
                "oid": oid,
                "username": data["username"],
                "calls": data["calls"],
                "errors": data["errors"],
                "first_seen_iso": _iso(data["first_seen"]),
                "last_seen_iso": _iso(data["last_seen"]),
            }
            for oid, data in users_sorted[:_USER_TOP]
        ]

        return {
            "server": {
                "uptime_s": round(now - self._started_at, 3),
                "started_at_iso": _iso(self._started_at),
                "total_calls": self._total_calls,
                "total_errors": self._total_errors,
            },
            "traffic": {
                "last_5m": _window(series[-5:]),
                "last_60m": _window(series),
                "per_minute": series,
            },
            "tools": tools,
            "users": {
                "count": len(self._users),
                "evicted": self._users_evicted,
                "top": top_users,
            },
        }

    def render_prometheus(self) -> str:
        """Render metrics in the Prometheus text exposition format (0.0.4).

        Hand-rolled (no client library). Emits process-global counters/gauges
        and per-tool series. There is deliberately **NO per-user label series**:
        caller identities are unbounded and high-cardinality, which is a
        well-known Prometheus anti-pattern (it explodes the time-series
        database). Per-user detail lives in ``/stats`` and ``/dashboard``
        instead; only an aggregate ``mcp_users_tracked`` gauge is exposed here.
        """
        now = time.time()
        lines: list[str] = []

        lines.append("# HELP mcp_uptime_seconds Seconds since the metrics registry started.")
        lines.append("# TYPE mcp_uptime_seconds gauge")
        lines.append(f"mcp_uptime_seconds {now - self._started_at:.3f}")

        lines.append("# HELP mcp_calls_total Total tool calls recorded.")
        lines.append("# TYPE mcp_calls_total counter")
        lines.append(f"mcp_calls_total {self._total_calls}")

        lines.append("# HELP mcp_errors_total Total tool calls that ended in error.")
        lines.append("# TYPE mcp_errors_total counter")
        lines.append(f"mcp_errors_total {self._total_errors}")

        lines.append("# HELP mcp_users_tracked Distinct caller identities currently tracked.")
        lines.append("# TYPE mcp_users_tracked gauge")
        lines.append(f"mcp_users_tracked {len(self._users)}")

        lines.append(
            "# HELP mcp_users_evicted_total Caller identities evicted due to the tracking cap."
        )
        lines.append("# TYPE mcp_users_evicted_total counter")
        lines.append(f"mcp_users_evicted_total {self._users_evicted}")

        # Per-tool series (sorted by name for stable, diff-friendly output).
        ordered_tools = sorted(self._tools.items())

        lines.append("# HELP mcp_tool_calls_total Tool calls, labelled by tool.")
        lines.append("# TYPE mcp_tool_calls_total counter")
        for name, stats in ordered_tools:
            lines.append(
                f'mcp_tool_calls_total{{tool="{_escape_label(name)}"}} {stats["calls"]}'
            )

        lines.append("# HELP mcp_tool_errors_total Tool calls that errored, labelled by tool.")
        lines.append("# TYPE mcp_tool_errors_total counter")
        for name, stats in ordered_tools:
            lines.append(
                f'mcp_tool_errors_total{{tool="{_escape_label(name)}"}} {stats["errors"]}'
            )

        lines.append(
            "# HELP mcp_tool_duration_ms Tool call duration in milliseconds, by statistic."
        )
        lines.append("# TYPE mcp_tool_duration_ms gauge")
        for name, stats in ordered_tools:
            durations = sorted(stats["durations"])
            if not durations:
                continue
            escaped = _escape_label(name)
            avg = sum(durations) / len(durations)
            p50 = _percentile(durations, 50)
            p95 = _percentile(durations, 95)
            lines.append(
                f'mcp_tool_duration_ms{{tool="{escaped}",stat="p50"}} {p50:.3f}'
            )
            lines.append(
                f'mcp_tool_duration_ms{{tool="{escaped}",stat="p95"}} {p95:.3f}'
            )
            lines.append(
                f'mcp_tool_duration_ms{{tool="{escaped}",stat="avg"}} {avg:.3f}'
            )

        return "\n".join(lines) + "\n"


def _escape_label(value: str) -> str:
    """Escape a Prometheus label value per the text exposition format.

    Backslash, double-quote, and newline are the three characters that must be
    escaped in a label value.
    """
    return value.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")


_registry: MetricsRegistry | None = None


def get_metrics_registry() -> MetricsRegistry:
    """Return the process-wide metrics registry, creating it on first use."""
    global _registry
    if _registry is None:
        _registry = MetricsRegistry()
    return _registry


def reset_metrics_registry() -> None:
    """Drop the cached registry so the next access rebuilds it.

    Wired into ``runtime.reset_runtime_state`` so tests (and a config reload)
    start from a clean, zeroed registry.
    """
    global _registry
    _registry = None
