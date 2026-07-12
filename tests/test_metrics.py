"""Unit tests for the observability metrics registry and middleware.

Covers the in-process :class:`MetricsRegistry` aggregation math and the
:class:`MetricsMiddleware` that feeds it. Route-level (ASGI) coverage lives in
tests/test_observability.py.
"""

from __future__ import annotations

import time

import pytest
from fastmcp.server.middleware import MiddlewareContext
from mcp import types as mt

import mcp_microsoft.metrics as metrics
from mcp_microsoft.metrics import (
    MetricsRegistry,
    _percentile,
    get_metrics_registry,
    reset_metrics_registry,
)
from mcp_microsoft.middleware import MetricsMiddleware


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


class _FakeAccessToken:
    def __init__(self, claims: dict) -> None:
        self.claims = claims


def _ctx(tool_name: str = "send_email") -> MiddlewareContext:
    params = mt.CallToolRequestParams(
        name=tool_name, arguments={"secret_argument": "sensitive-value"}
    )
    return MiddlewareContext(message=params, method="tools/call", type="request")


# ---------------------------------------------------------------------------
# registry: counters & error counts
# ---------------------------------------------------------------------------


def test_records_calls_and_errors() -> None:
    reg = MetricsRegistry()
    reg.record("list_emails", "oid-1", "alice@x.com", 12.0, True)
    reg.record("list_emails", "oid-1", "alice@x.com", 34.0, False)
    reg.record("send_email", "oid-2", "bob@x.com", 5.0, True)

    snap = reg.snapshot()
    assert snap["server"]["total_calls"] == 3
    assert snap["server"]["total_errors"] == 1

    tools = {t["name"]: t for t in snap["tools"]}
    assert tools["list_emails"]["calls"] == 2
    assert tools["list_emails"]["errors"] == 1
    assert tools["send_email"]["calls"] == 1
    assert tools["send_email"]["errors"] == 0
    # Sorted by calls desc.
    assert snap["tools"][0]["name"] == "list_emails"

    users = {u["oid"]: u for u in snap["users"]["top"]}
    assert users["oid-1"]["calls"] == 2
    assert users["oid-1"]["errors"] == 1
    assert users["oid-2"]["calls"] == 1


def test_unknown_identity_collapses_to_dash() -> None:
    reg = MetricsRegistry()
    reg.record("list_emails", "", "", 1.0, True)
    reg.record("list_emails", "-", "-", 1.0, True)

    snap = reg.snapshot()
    assert snap["users"]["count"] == 1
    assert snap["users"]["top"][0]["oid"] == "-"
    assert snap["users"]["top"][0]["calls"] == 2


# ---------------------------------------------------------------------------
# registry: percentile math on a known sequence
# ---------------------------------------------------------------------------


def test_percentile_helper_linear_interpolation() -> None:
    seq = [10.0, 20.0, 30.0, 40.0, 50.0, 60.0, 70.0, 80.0, 90.0, 100.0]
    assert _percentile(seq, 50) == 55.0
    assert _percentile(seq, 95) == 95.5
    assert _percentile([], 50) == 0.0
    assert _percentile([42.0], 95) == 42.0


def test_snapshot_computes_p50_p95_avg() -> None:
    reg = MetricsRegistry()
    for ms in (10, 20, 30, 40, 50, 60, 70, 80, 90, 100):
        reg.record("t", "oid", "u", float(ms), True)

    tool = reg.snapshot()["tools"][0]
    assert tool["avg_ms"] == 55.0
    assert tool["p50_ms"] == 55.0
    assert tool["p95_ms"] == 95.5


# ---------------------------------------------------------------------------
# registry: minute-bucket rollover
# ---------------------------------------------------------------------------


def test_minute_bucket_rollover(monkeypatch: pytest.MonkeyPatch) -> None:
    clock = {"t": 1_000_000.0}
    monkeypatch.setattr(time, "time", lambda: clock["t"])

    reg = MetricsRegistry()
    reg.record("t", "u", "u", 1.0, True)

    snap = reg.snapshot()
    pm = snap["traffic"]["per_minute"]
    assert len(pm) == 60
    assert pm[-1]["calls"] == 1
    assert snap["traffic"]["last_5m"]["calls"] == 1
    assert snap["traffic"]["last_60m"]["calls"] == 1

    # Advance three whole minutes and record again.
    clock["t"] += 180
    reg.record("t", "u", "u", 1.0, True)
    snap = reg.snapshot()
    pm = snap["traffic"]["per_minute"]
    assert pm[-1]["calls"] == 1  # newest minute
    assert pm[-4]["calls"] == 1  # three minutes earlier
    assert snap["traffic"]["last_5m"]["calls"] == 2
    assert snap["traffic"]["last_60m"]["calls"] == 2

    # Advance past the whole 60-minute window: both calls fall out of the
    # rolling timeline, but cumulative totals are unaffected.
    clock["t"] += 65 * 60
    snap = reg.snapshot()
    assert snap["traffic"]["last_60m"]["calls"] == 0
    assert snap["server"]["total_calls"] == 2


# ---------------------------------------------------------------------------
# registry: user cap eviction
# ---------------------------------------------------------------------------


def test_user_cap_evicts_least_recently_seen(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(metrics, "_USER_CAP", 3)
    reg = MetricsRegistry()

    for i in range(3):
        reg.record("t", f"user-{i}", f"u{i}", 1.0, True)
    assert reg.snapshot()["users"]["count"] == 3
    assert reg.snapshot()["users"]["evicted"] == 0

    # Re-touch user-0 so user-1 is now the least-recently-seen identity.
    reg.record("t", "user-0", "u0", 1.0, True)
    # A brand-new user overflows the cap and evicts the LRU (user-1).
    reg.record("t", "user-3", "u3", 1.0, True)

    snap = reg.snapshot()
    assert snap["users"]["count"] == 3
    assert snap["users"]["evicted"] == 1
    oids = {u["oid"] for u in snap["users"]["top"]}
    assert oids == {"user-0", "user-2", "user-3"}
    assert "user-1" not in oids


def test_users_top_sorted_by_last_seen_desc(monkeypatch: pytest.MonkeyPatch) -> None:
    clock = {"t": 1_000.0}
    monkeypatch.setattr(time, "time", lambda: clock["t"])
    reg = MetricsRegistry()

    reg.record("t", "early", "e", 1.0, True)
    clock["t"] += 10
    reg.record("t", "late", "l", 1.0, True)

    top = reg.snapshot()["users"]["top"]
    assert [u["oid"] for u in top] == ["late", "early"]


# ---------------------------------------------------------------------------
# registry: unknown-tool aggregate counter (poisoning resistance)
# ---------------------------------------------------------------------------


def test_record_unknown_tool_is_aggregate_only() -> None:
    reg = MetricsRegistry()
    # A couple of legitimate calls to establish a baseline.
    reg.record("list_emails", "oid-1", "alice@x.com", 5.0, True)
    reg.record("list_emails", "oid-1", "alice@x.com", 5.0, False)

    for _ in range(7):
        reg.record_unknown_tool()

    snap = reg.snapshot()
    # The aggregate counter caught every unknown call...
    assert snap["server"]["unknown_tool_calls"] == 7
    # ...without touching the global totals, the per-tool dict, or the minute
    # error buckets (poisoning resistance).
    assert snap["server"]["total_calls"] == 2
    assert snap["server"]["total_errors"] == 1
    assert {t["name"] for t in snap["tools"]} == {"list_emails"}
    assert snap["traffic"]["last_60m"]["errors"] == 1


# ---------------------------------------------------------------------------
# registry: per-tool LRU cap (backstop against name injection)
# ---------------------------------------------------------------------------


def test_tool_cap_evicts_least_recently_seen(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(metrics, "_TOOL_CAP", 3)
    reg = MetricsRegistry()

    for i in range(3):
        reg.record(f"tool-{i}", "oid", "u", 1.0, True)
    snap = reg.snapshot()
    assert len(snap["tools"]) == 3
    assert snap["server"]["tools_evicted"] == 0

    # Re-touch tool-0 so tool-1 becomes the least-recently-seen entry.
    reg.record("tool-0", "oid", "u", 1.0, True)
    # A brand-new tool overflows the cap and evicts the LRU (tool-1).
    reg.record("tool-3", "oid", "u", 1.0, True)

    snap = reg.snapshot()
    names = {t["name"] for t in snap["tools"]}
    assert len(names) == 3
    assert names == {"tool-0", "tool-2", "tool-3"}
    assert "tool-1" not in names
    assert snap["server"]["tools_evicted"] == 1


def test_tool_cap_bounds_size_under_many_distinct_names(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(metrics, "_TOOL_CAP", 8)
    reg = MetricsRegistry()

    for i in range(100):
        reg.record(f"tool-{i}", "oid", "u", 1.0, True)

    snap = reg.snapshot()
    assert len(snap["tools"]) == 8
    assert snap["server"]["tools_evicted"] == 92


# ---------------------------------------------------------------------------
# registry: prometheus rendering
# ---------------------------------------------------------------------------


def test_prometheus_has_per_tool_series_and_escapes_labels() -> None:
    reg = MetricsRegistry()
    reg.record('weird"\\tool\nname', "oid-1", "u", 5.0, True)

    out = reg.render_prometheus()

    # Per-tool counter series present, with the label value escaped per the
    # exposition format (\" for quote, \\ for backslash, \n for newline).
    assert 'mcp_tool_calls_total{tool="weird\\"\\\\tool\\nname"} 1' in out
    assert 'mcp_tool_errors_total{tool="weird\\"\\\\tool\\nname"} 0' in out
    assert 'stat="p50"' in out and 'stat="p95"' in out and 'stat="avg"' in out


def test_prometheus_has_no_per_user_series() -> None:
    reg = MetricsRegistry()
    reg.record("send_email", "oid-secret-value", "victim@example.com", 5.0, True)

    out = reg.render_prometheus()

    # Aggregate user gauges are fine...
    assert "mcp_users_tracked 1" in out
    assert "mcp_users_evicted_total 0" in out
    # ...but no per-user label series, and the identity never leaks into the
    # exposition (high-cardinality anti-pattern; see render_prometheus docstring).
    assert 'oid="' not in out
    assert "oid-secret-value" not in out
    assert "victim@example.com" not in out


def test_prometheus_totals_and_types() -> None:
    reg = MetricsRegistry()
    reg.record("send_email", "oid-1", "u", 5.0, True)
    reg.record("send_email", "oid-1", "u", 5.0, False)

    out = reg.render_prometheus()
    assert "# TYPE mcp_calls_total counter" in out
    assert "mcp_calls_total 2" in out
    assert "mcp_errors_total 1" in out
    assert "# TYPE mcp_uptime_seconds gauge" in out
    assert out.endswith("\n")


def test_prometheus_exposes_unknown_tool_and_eviction_counters(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(metrics, "_TOOL_CAP", 1)
    reg = MetricsRegistry()
    reg.record_unknown_tool()
    reg.record_unknown_tool()
    reg.record("tool-a", "oid", "u", 1.0, True)
    reg.record("tool-b", "oid", "u", 1.0, True)  # evicts tool-a (cap == 1)

    out = reg.render_prometheus()
    assert "# TYPE mcp_unknown_tool_calls_total counter" in out
    assert "mcp_unknown_tool_calls_total 2" in out
    assert "# TYPE mcp_tools_evicted_total counter" in out
    assert "mcp_tools_evicted_total 1" in out


def test_escape_label_escapes_carriage_return() -> None:
    # \r must be escaped so a label value can never break the one-sample-per-
    # line framing of the exposition format.
    assert metrics._escape_label("a\rb") == "a\\rb"
    # All four special characters, backslash first.
    assert metrics._escape_label('x\\y"z\na\rb') == 'x\\\\y\\"z\\na\\rb'

    reg = MetricsRegistry()
    reg.record("carriage\rreturn", "oid", "u", 1.0, True)
    out = reg.render_prometheus()
    assert 'mcp_tool_calls_total{tool="carriage\\rreturn"} 1' in out
    # The raw CR never reaches the wire.
    assert "\r" not in out


# ---------------------------------------------------------------------------
# registry: snapshot shape
# ---------------------------------------------------------------------------


def test_snapshot_shape() -> None:
    reg = MetricsRegistry()
    reg.record("send_email", "oid-1", "alice@x.com", 5.0, True)

    snap = reg.snapshot()
    assert set(snap) == {"server", "traffic", "tools", "users"}
    assert set(snap["server"]) == {
        "uptime_s",
        "started_at_iso",
        "total_calls",
        "total_errors",
        "unknown_tool_calls",
        "tools_evicted",
    }
    assert set(snap["traffic"]) == {"last_5m", "last_60m", "per_minute"}
    assert set(snap["traffic"]["last_5m"]) == {"calls", "errors"}
    assert set(snap["tools"][0]) == {
        "name",
        "calls",
        "errors",
        "avg_ms",
        "p50_ms",
        "p95_ms",
    }
    assert set(snap["users"]) == {"count", "evicted", "top"}
    assert set(snap["users"]["top"][0]) == {
        "oid",
        "username",
        "calls",
        "errors",
        "first_seen_iso",
        "last_seen_iso",
    }


def test_singleton_reset() -> None:
    first = get_metrics_registry()
    assert get_metrics_registry() is first
    reset_metrics_registry()
    assert get_metrics_registry() is not first


# ---------------------------------------------------------------------------
# MetricsMiddleware
# ---------------------------------------------------------------------------


class _FakeRegistry:
    def __init__(self, explode: bool = False) -> None:
        self.calls: list[tuple] = []
        self.unknown_calls = 0
        self._explode = explode

    def record(self, tool, oid, username, duration_ms, ok) -> None:
        if self._explode:
            raise RuntimeError("registry boom")
        self.calls.append((tool, oid, username, ok))

    def record_unknown_tool(self) -> None:
        if self._explode:
            raise RuntimeError("registry boom")
        self.unknown_calls += 1


@pytest.mark.asyncio
async def test_middleware_records_success(monkeypatch: pytest.MonkeyPatch) -> None:
    import fastmcp.server.dependencies as deps

    token = _FakeAccessToken({"oid": "o1", "preferred_username": "a@x.com"})
    monkeypatch.setattr(deps, "get_access_token", lambda: token)

    reg = _FakeRegistry()
    mw = MetricsMiddleware(registry=reg)

    async def call_next(_ctx: MiddlewareContext) -> str:
        return "ok"

    result = await mw.on_call_tool(_ctx("send_email"), call_next)

    assert result == "ok"
    assert len(reg.calls) == 1
    tool, oid, username, ok = reg.calls[0]
    assert (tool, oid, username, ok) == ("send_email", "o1", "a@x.com", True)


@pytest.mark.asyncio
async def test_middleware_records_error_and_reraises(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import fastmcp.server.dependencies as deps

    monkeypatch.setattr(deps, "get_access_token", lambda: None)

    reg = _FakeRegistry()
    mw = MetricsMiddleware(registry=reg)

    async def call_next(_ctx: MiddlewareContext) -> str:
        raise RuntimeError("boom")

    with pytest.raises(RuntimeError, match="boom"):
        await mw.on_call_tool(_ctx("delete_email"), call_next)

    # Still recorded, with ok=False and the unauthenticated "-" identity.
    assert reg.calls == [("delete_email", "-", "-", False)]


@pytest.mark.asyncio
async def test_middleware_never_raises_when_registry_explodes_on_success(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import fastmcp.server.dependencies as deps

    monkeypatch.setattr(deps, "get_access_token", lambda: None)

    mw = MetricsMiddleware(registry=_FakeRegistry(explode=True))

    async def call_next(_ctx: MiddlewareContext) -> str:
        return "ok"

    # A broken registry must not break the tool call it observes.
    assert await mw.on_call_tool(_ctx(), call_next) == "ok"


@pytest.mark.asyncio
async def test_middleware_registry_failure_does_not_mask_tool_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import fastmcp.server.dependencies as deps

    monkeypatch.setattr(deps, "get_access_token", lambda: None)

    mw = MetricsMiddleware(registry=_FakeRegistry(explode=True))

    async def call_next(_ctx: MiddlewareContext) -> str:
        raise ValueError("tool failure")

    # The tool's own exception propagates unchanged, not the registry's.
    with pytest.raises(ValueError, match="tool failure"):
        await mw.on_call_tool(_ctx(), call_next)


@pytest.mark.asyncio
async def test_middleware_unknown_tool_is_not_recorded_per_name(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A NotFoundError (unknown tool name) is counted in aggregate only."""
    import fastmcp.server.dependencies as deps
    from fastmcp.exceptions import NotFoundError

    monkeypatch.setattr(deps, "get_access_token", lambda: None)

    reg = _FakeRegistry()
    mw = MetricsMiddleware(registry=reg)

    async def call_next(_ctx: MiddlewareContext):
        raise NotFoundError("Unknown tool: 'attacker-chosen-name'")

    with pytest.raises(NotFoundError):
        await mw.on_call_tool(_ctx("attacker-chosen-name"), call_next)

    # The arbitrary name never entered the per-tool record path...
    assert reg.calls == []
    # ...only the aggregate unknown-tool counter moved.
    assert reg.unknown_calls == 1


@pytest.mark.asyncio
async def test_middleware_known_tool_validation_error_recorded_under_real_name(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Argument-validation failures on a REAL tool are still recorded as errors.

    fastmcp raises ``ValidationError`` (a ``FastMCPError``, distinct from
    ``NotFoundError``) for bad arguments on a resolved tool, so the specific
    NotFoundError branch must not swallow it.
    """
    import fastmcp.server.dependencies as deps
    from fastmcp.exceptions import ValidationError

    monkeypatch.setattr(deps, "get_access_token", lambda: None)

    reg = _FakeRegistry()
    mw = MetricsMiddleware(registry=reg)

    async def call_next(_ctx: MiddlewareContext):
        raise ValidationError("1 validation error for call[list_emails]")

    with pytest.raises(ValidationError):
        await mw.on_call_tool(_ctx("list_emails"), call_next)

    # Recorded as an error under its real name; not filtered as an unknown tool.
    assert reg.calls == [("list_emails", "-", "-", False)]
    assert reg.unknown_calls == 0


@pytest.mark.asyncio
async def test_middleware_defaults_to_singleton_registry(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import fastmcp.server.dependencies as deps

    token = _FakeAccessToken({"oid": "o9", "preferred_username": "z@x.com"})
    monkeypatch.setattr(deps, "get_access_token", lambda: token)
    reset_metrics_registry()

    mw = MetricsMiddleware()  # no explicit registry -> process singleton

    async def call_next(_ctx: MiddlewareContext) -> str:
        return "ok"

    await mw.on_call_tool(_ctx("list_emails"), call_next)

    snap = get_metrics_registry().snapshot()
    assert snap["server"]["total_calls"] == 1
    assert snap["tools"][0]["name"] == "list_emails"
    reset_metrics_registry()
