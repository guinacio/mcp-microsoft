"""ASGI-level tests for the token-gated observability routes.

Exercises ``GET /metrics``, ``GET /stats``, and ``GET /dashboard`` end to end
via ``httpx.ASGITransport`` (the same pattern as tests/test_integration_http.py),
covering the token gating (absent when unset; 401/200 for both credential
schemes when set), content types, and the stdio-mode / no-token absence cases.
Registry aggregation math is unit-tested separately in tests/test_metrics.py.
"""

from __future__ import annotations

import base64

import httpx
import pytest

from mcp_microsoft.metrics import get_metrics_registry
from mcp_microsoft.runtime import reset_runtime_state

_HTTP_ENV_VARS = (
    "MCP_TRANSPORT",
    "MCP_HTTP_HOST",
    "MCP_HTTP_PORT",
    "MCP_HTTP_STATELESS",
    "MCP_BASE_URL",
    "MCP_AUTH_CLIENT_ID",
    "MCP_AUTH_CLIENT_SECRET",
    "MCP_AUTH_TENANT_ID",
    "MCP_AUTH_REQUIRED_SCOPE",
    "MCP_RATE_LIMIT_RPS",
    "MCP_STATS_TOKEN",
)

_FULL_HTTP_ENV = {
    "MCP_TRANSPORT": "http",
    "MCP_BASE_URL": "https://mcp.example.com",
    "MCP_AUTH_CLIENT_ID": "cid",
    "MCP_AUTH_CLIENT_SECRET": "secret",
    "MCP_AUTH_TENANT_ID": "organizations",
}

_TOKEN = "s3cr3t-stats-token"


@pytest.fixture(autouse=True)
def _reset_state():
    reset_runtime_state()
    yield
    reset_runtime_state()


def _build_server(monkeypatch: pytest.MonkeyPatch, **overrides: str):
    import mcp_microsoft.server as server_mod

    for name in _HTTP_ENV_VARS:
        monkeypatch.delenv(name, raising=False)
    for name, value in {**_FULL_HTTP_ENV, **overrides}.items():
        monkeypatch.setenv(name, value)
    reset_runtime_state()
    return server_mod.get_mcp_server(reset=True)


def _client(mcp) -> httpx.AsyncClient:
    transport = httpx.ASGITransport(app=mcp.http_app())
    return httpx.AsyncClient(transport=transport, base_url="http://test")


def _basic(user: str, password: str) -> dict[str, str]:
    raw = base64.b64encode(f"{user}:{password}".encode()).decode()
    return {"Authorization": f"Basic {raw}"}


# ---------------------------------------------------------------------------
# routes absent unless MCP_STATS_TOKEN is set
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_routes_absent_when_token_unset(monkeypatch: pytest.MonkeyPatch) -> None:
    mcp = _build_server(monkeypatch)  # no MCP_STATS_TOKEN
    async with _client(mcp) as client:
        for path in ("/metrics", "/stats", "/dashboard"):
            resp = await client.get(path)
            assert resp.status_code == 404, f"{path} should be unregistered"
        # /health is always on in http mode, unauthenticated.
        health = await client.get("/health")
        assert health.status_code == 200


@pytest.mark.asyncio
async def test_routes_absent_in_stdio_mode(
    monkeypatch: pytest.MonkeyPatch, tmp_path
) -> None:
    """Even with a token set, stdio mode registers no custom http routes."""
    import mcp_microsoft.server as server_mod

    for name in _HTTP_ENV_VARS:
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    monkeypatch.setenv("MCP_STATS_TOKEN", _TOKEN)  # ignored in stdio mode
    reset_runtime_state()

    mcp = server_mod.get_mcp_server(reset=True)
    assert mcp._additional_http_routes == []


# ---------------------------------------------------------------------------
# auth gating when the token is set
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_401_without_credentials(monkeypatch: pytest.MonkeyPatch) -> None:
    mcp = _build_server(monkeypatch, MCP_STATS_TOKEN=_TOKEN)
    async with _client(mcp) as client:
        for path in ("/metrics", "/stats", "/dashboard"):
            resp = await client.get(path)
            assert resp.status_code == 401
            www = resp.headers.get("www-authenticate", "")
            assert www == 'Basic realm="mcp-microsoft stats", Bearer'
            # No detail leaked in the body.
            assert _TOKEN not in resp.text


@pytest.mark.asyncio
async def test_401_with_wrong_bearer(monkeypatch: pytest.MonkeyPatch) -> None:
    mcp = _build_server(monkeypatch, MCP_STATS_TOKEN=_TOKEN)
    async with _client(mcp) as client:
        resp = await client.get(
            "/stats", headers={"Authorization": "Bearer not-the-token"}
        )
        assert resp.status_code == 401


@pytest.mark.asyncio
async def test_200_with_correct_bearer(monkeypatch: pytest.MonkeyPatch) -> None:
    mcp = _build_server(monkeypatch, MCP_STATS_TOKEN=_TOKEN)
    async with _client(mcp) as client:
        resp = await client.get(
            "/stats", headers={"Authorization": f"Bearer {_TOKEN}"}
        )
        assert resp.status_code == 200


@pytest.mark.asyncio
async def test_200_with_correct_basic(monkeypatch: pytest.MonkeyPatch) -> None:
    mcp = _build_server(monkeypatch, MCP_STATS_TOKEN=_TOKEN)
    async with _client(mcp) as client:
        # Any username, password == token.
        resp = await client.get("/stats", headers=_basic("operator", _TOKEN))
        assert resp.status_code == 200


@pytest.mark.asyncio
async def test_401_with_wrong_basic_password(monkeypatch: pytest.MonkeyPatch) -> None:
    mcp = _build_server(monkeypatch, MCP_STATS_TOKEN=_TOKEN)
    async with _client(mcp) as client:
        resp = await client.get("/stats", headers=_basic("operator", "wrong"))
        assert resp.status_code == 401


# ---------------------------------------------------------------------------
# content of each route
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_no_store_header_on_all_three_routes(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Live observability responses must forbid caching (Cache-Control: no-store)."""
    mcp = _build_server(monkeypatch, MCP_STATS_TOKEN=_TOKEN)
    async with _client(mcp) as client:
        for path in ("/metrics", "/stats", "/dashboard"):
            resp = await client.get(
                path, headers={"Authorization": f"Bearer {_TOKEN}"}
            )
            assert resp.status_code == 200, path
            assert resp.headers.get("cache-control") == "no-store", path


@pytest.mark.asyncio
async def test_unknown_tool_calls_do_not_poison_metrics(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The confirmed HIGH finding's repro: unknown tool names sent by an
    authenticated caller must NOT grow per-tool cardinality, inflate the
    global/error counters or minute buckets, or leak into the Prometheus
    ``tool`` label. They are absorbed by a single aggregate counter.

    Driven through a real http-mode server via an in-memory ``fastmcp.Client``
    (which is where fastmcp raises ``NotFoundError`` inside the middleware
    chain), then verified against both the registry snapshot and the scraped
    /metrics wire output (both read the same process-singleton registry).
    """
    from fastmcp import Client

    mcp = _build_server(monkeypatch, MCP_STATS_TOKEN=_TOKEN)
    registry = get_metrics_registry()

    bogus_names = [f"bogus_tool_{i}" for i in range(6)]
    async with Client(mcp) as client:
        for name in bogus_names:
            # Each unknown name surfaces client-side as an error; the server
            # raised NotFoundError inside the metrics middleware.
            with pytest.raises(Exception):
                await client.call_tool(name, {})

    snap = registry.snapshot()
    tool_names = {t["name"] for t in snap["tools"]}
    for name in bogus_names:
        assert name not in tool_names, f"{name} poisoned the per-tool registry"

    # Aggregate counter absorbed every unknown call; nothing else moved.
    assert snap["server"]["unknown_tool_calls"] == len(bogus_names)
    assert snap["server"]["total_calls"] == 0
    assert snap["server"]["total_errors"] == 0
    assert snap["traffic"]["last_60m"]["calls"] == 0
    assert snap["traffic"]["last_60m"]["errors"] == 0

    # Prometheus: the aggregate line is present and carries the count; no bogus
    # per-tool series leaked onto the wire.
    async with _client(mcp) as http:
        resp = await http.get(
            "/metrics", headers={"Authorization": f"Bearer {_TOKEN}"}
        )
    body = resp.text
    assert f"mcp_unknown_tool_calls_total {len(bogus_names)}" in body
    for name in bogus_names:
        assert name not in body, f"{name} leaked into the Prometheus output"


@pytest.mark.asyncio
async def test_metrics_content_type_and_known_line(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    mcp = _build_server(monkeypatch, MCP_STATS_TOKEN=_TOKEN)
    # Record into the same process singleton the route reads.
    get_metrics_registry().record("send_email", "oid-1", "alice@x.com", 5.0, True)

    async with _client(mcp) as client:
        resp = await client.get(
            "/metrics", headers={"Authorization": f"Bearer {_TOKEN}"}
        )

    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith("text/plain")
    assert "version=0.0.4" in resp.headers["content-type"]
    body = resp.text
    assert "# TYPE mcp_calls_total counter" in body
    assert 'mcp_tool_calls_total{tool="send_email"} 1' in body


@pytest.mark.asyncio
async def test_stats_json_shape(monkeypatch: pytest.MonkeyPatch) -> None:
    mcp = _build_server(monkeypatch, MCP_STATS_TOKEN=_TOKEN)
    async with _client(mcp) as client:
        resp = await client.get(
            "/stats", headers={"Authorization": f"Bearer {_TOKEN}"}
        )

    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith("application/json")
    body = resp.json()
    assert set(body) == {"server", "traffic", "tools", "users"}
    assert "total_calls" in body["server"]
    assert len(body["traffic"]["per_minute"]) == 60


@pytest.mark.asyncio
async def test_dashboard_is_self_contained_html(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    mcp = _build_server(monkeypatch, MCP_STATS_TOKEN=_TOKEN)
    async with _client(mcp) as client:
        resp = await client.get("/dashboard", headers=_basic("op", _TOKEN))

    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith("text/html")
    html = resp.text
    # Fetches the same-origin /stats endpoint (browser re-sends Basic creds).
    assert 'fetch("/stats"' in html
    # Self-contained: no external network dependencies.
    assert "http://" not in html.replace("http://test", "")
    assert "https://" not in html
    assert "<script>" in html


# ---------------------------------------------------------------------------
# /health remains unauthenticated even with the stats token set
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_health_still_unauthenticated_with_token_set(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    mcp = _build_server(monkeypatch, MCP_STATS_TOKEN=_TOKEN)
    async with _client(mcp) as client:
        resp = await client.get("/health")
    assert resp.status_code == 200
    assert resp.json() == {"status": "ok", "transport": "http"}
