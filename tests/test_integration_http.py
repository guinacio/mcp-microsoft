"""
End-to-end integration tests for http (multi-user) transport.

These go beyond the unit-level coverage in test_identity.py (token-provider
recipes in isolation) and the registration-matrix coverage in
test_http_transport.py (config, tool gating, middleware wiring). Here we
exercise real seams end-to-end:

1. A full authenticated tool call over an in-memory ``fastmcp.Client``,
   proving the tool -> get_graph -> OboTokenProvider -> Graph pipeline.
2. The ASGI-level protocol/auth surface (401 challenge + RFC 9728 protected
   resource metadata) via ``mcp.http_app()`` + ``httpx.ASGITransport``.
3. Per-user isolation at the OboTokenProvider/GraphClient seam: two
   sequential "requests" from different ambient identities must never bleed
   Graph bearer tokens across users.
4. A cheap stdio-mode regression canary.

``GET /health`` is deliberately not repeated here: it is already covered by
``test_http_mode_health_endpoint_is_unauthenticated`` in
tests/test_http_transport.py using the identical ASGITransport pattern.
"""

from __future__ import annotations

import httpx
import pytest
from fastmcp import Client

import mcp_microsoft.graph as graph_module
import mcp_microsoft.identity as identity_module
from mcp_microsoft.graph import GraphClient
from mcp_microsoft.identity import OboTokenProvider
from mcp_microsoft.models import ListEmailsResponse
from mcp_microsoft.runtime import reset_runtime_state

# ---------------------------------------------------------------------------
# Shared fixtures / helpers
# ---------------------------------------------------------------------------

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
)

# Structurally valid dummy tenant GUID (http mode requires a concrete GUID).
_TENANT_GUID = "11111111-1111-1111-1111-111111111111"

_FULL_HTTP_ENV = {
    "MCP_TRANSPORT": "http",
    "MCP_BASE_URL": "https://mcp.example.com",
    "MCP_AUTH_CLIENT_ID": "cid",
    "MCP_AUTH_CLIENT_SECRET": "secret",
    "MCP_AUTH_TENANT_ID": _TENANT_GUID,
}

_GRAPH_MESSAGES_PAGE = {
    "value": [
        {
            "id": "msg-1",
            "subject": "Hello",
            "from": {"emailAddress": {"name": "Alice", "address": "alice@example.com"}},
            "receivedDateTime": "2026-07-01T12:00:00Z",
            "isRead": True,
            "bodyPreview": "Hi there",
            "hasAttachments": False,
            "importance": "normal",
        }
    ]
}


@pytest.fixture(autouse=True)
def _reset_state():
    """Every test in this file starts and ends with a clean runtime.

    ``reset_runtime_state`` clears the cached AppConfig/ProfileManager/
    FastMCP server singletons; the shared http-mode GraphClient
    (``graph._obo_graph_client``) is a separate module-level cache not
    covered by that helper, so it is reset explicitly to stop a client bound
    to one test's (closed) mock transport from leaking into the next test.
    """
    reset_runtime_state()
    graph_module._obo_graph_client = None
    yield
    reset_runtime_state()
    graph_module._obo_graph_client = None


def _build_http_server(monkeypatch: pytest.MonkeyPatch):
    """Build a real http-mode FastMCP server via the env-driven bootstrap path.

    Mirrors the fixture pattern in tests/test_http_transport.py.
    """
    import mcp_microsoft.server as server_mod

    for name in _HTTP_ENV_VARS:
        monkeypatch.delenv(name, raising=False)
    for name, value in _FULL_HTTP_ENV.items():
        monkeypatch.setenv(name, value)
    reset_runtime_state()
    return server_mod.get_mcp_server(reset=True)


# ---------------------------------------------------------------------------
# 1. Full-stack authenticated tool call (in-memory fastmcp.Client)
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_authenticated_tool_call_reaches_graph_with_obo_token(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """tool -> get_graph -> OboTokenProvider -> Graph, wired end-to-end.

    The in-memory client transport bypasses HTTP-layer bearer enforcement
    (that surface is covered by the ASGI tests below), so identity here is
    injected by monkeypatching ``OboTokenProvider.get_access_token``
    directly -- exactly the seam ``get_graph()`` hands the real request
    context through in production.
    """
    mcp = _build_http_server(monkeypatch)

    captured: dict[str, str | None] = {}

    def handler(request: httpx.Request) -> httpx.Response:
        captured["auth"] = request.headers.get("Authorization")
        captured["url"] = str(request.url)
        return httpx.Response(200, json=_GRAPH_MESSAGES_PAGE)

    mock_client = httpx.AsyncClient(transport=httpx.MockTransport(handler), timeout=5.0)
    monkeypatch.setattr(graph_module, "_request_client", mock_client)

    async def fake_get_access_token(self: OboTokenProvider) -> str:
        return "fake-graph-token"

    monkeypatch.setattr(OboTokenProvider, "get_access_token", fake_get_access_token)

    try:
        async with Client(mcp) as client:
            result = await client.call_tool(
                "list_emails", {"params": {"folder": "inbox", "max_results": 5}}
            )
    finally:
        await mock_client.aclose()

    assert result.is_error is False
    # Round-trips through the tool's declared Pydantic response model.
    parsed = ListEmailsResponse.model_validate(result.structured_content)
    assert parsed.count == 1
    assert parsed.messages[0].subject == "Hello"
    assert parsed.messages[0].from_.address == "alice@example.com"

    assert captured["auth"] == "Bearer fake-graph-token"
    assert "/me/mailFolders/inbox/messages" in captured["url"]


@pytest.mark.asyncio
async def test_error_masking_hides_sensitive_details_from_client(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """http mode's mask_error_details keeps internal exception text off the wire.

    A tool driven to raise deep in the OBO path (a RuntimeError carrying a
    sensitive marker) must surface only a generic error to the client -- the
    marker must never reach the caller. Mirrors the fixture pattern of the
    authenticated-call test above.
    """
    mcp = _build_http_server(monkeypatch)

    async def boom(self: OboTokenProvider) -> str:
        raise RuntimeError("SENSITIVE_XYZ")

    monkeypatch.setattr(OboTokenProvider, "get_access_token", boom)

    async with Client(mcp) as client:
        result = await client.call_tool(
            "list_emails",
            {"params": {"folder": "inbox", "max_results": 5}},
            raise_on_error=False,
        )

    assert result.is_error is True
    serialized = "".join(
        str(getattr(block, "text", block)) for block in (result.content or [])
    )
    serialized += str(getattr(result, "structured_content", "") or "")
    assert "SENSITIVE_XYZ" not in serialized


# ---------------------------------------------------------------------------
# 1b. Real registration-path OBO resolution (no _find_azure_provider /
#     OboTokenProvider.get_access_token patching)
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_tool_call_resolves_real_azure_provider_through_server_auth(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Prove ``get_server().auth`` resolves to the REAL AzureProvider.

    Unlike test #1 (which stubs ``OboTokenProvider.get_access_token`` wholesale)
    and the isolation test below (which stubs ``_find_azure_provider``), this
    exercises the entire un-stubbed OBO recipe against the real server build:

      * the real ``OboTokenProvider.get_access_token`` runs,
      * it calls the real ``get_server()`` -> ``.auth`` (the AzureProvider the
        real ``create_mcp_server`` attached in http mode),
      * the real ``_find_azure_provider`` unwraps it,
      * only ``AzureProvider.get_obo_credential`` is patched (class-level) to
        skip the live Entra network exchange and hand back a fake credential
        whose Graph token is the literal ``"fake"``.

    The ambient bearer is injected by patching
    ``fastmcp.server.dependencies.get_access_token`` (OboTokenProvider imports
    it inside the method, so the fastmcp module is the patch point) -- the
    in-memory transport does no HTTP-layer bearer validation.

    A "Bearer fake" on the outbound Graph request is only possible if every one
    of those real seams lined up.
    """
    from fastmcp.server.auth.providers.azure import AzureProvider

    mcp = _build_http_server(monkeypatch)

    # Sanity: the real server really did attach an AzureProvider as its auth.
    assert isinstance(mcp.auth, AzureProvider)

    import fastmcp.server.dependencies as deps

    monkeypatch.setattr(
        deps,
        "get_access_token",
        lambda: _FakeAccessToken("user-assertion", claims={"oid": "user-1"}),
    )

    fake_credential = _FakeCredential("fake")

    async def fake_get_obo_credential(self: AzureProvider, user_assertion: str):
        return fake_credential

    # Patch the class method (not _find_azure_provider): the real provider
    # instance reached via get_server().auth is what gets this call.
    monkeypatch.setattr(
        AzureProvider, "get_obo_credential", fake_get_obo_credential
    )

    captured: dict[str, str | None] = {}

    def handler(request: httpx.Request) -> httpx.Response:
        captured["auth"] = request.headers.get("Authorization")
        return httpx.Response(200, json=_GRAPH_MESSAGES_PAGE)

    mock_client = httpx.AsyncClient(transport=httpx.MockTransport(handler), timeout=5.0)
    monkeypatch.setattr(graph_module, "_request_client", mock_client)

    try:
        async with Client(mcp) as client:
            result = await client.call_tool(
                "list_emails", {"params": {"folder": "inbox", "max_results": 5}}
            )
    finally:
        await mock_client.aclose()

    assert result.is_error is False
    assert captured["auth"] == "Bearer fake"
    assert fake_credential.get_token_calls == 1


# ---------------------------------------------------------------------------
# 2. ASGI-level protocol/auth surface
# ---------------------------------------------------------------------------


async def _fetch_protected_resource_metadata(client: httpx.AsyncClient) -> httpx.Response:
    """Fetch RFC 9728 protected-resource metadata, probing both plausible paths.

    Per RFC 9728 sec 3.1 and ``mcp.server.auth.routes.build_resource_metadata_url``,
    fastmcp registers this route relative to the *resource's own path*
    (``/mcp`` here), not the bare well-known path. Verified directly against
    the installed fastmcp 3.4.4: ``GET /.well-known/oauth-protected-resource/mcp``
    -> 200; the bare path 404s. Both are probed so this test does not
    silently rot if a future fastmcp version changes the mount convention.
    """
    response = await client.get("/.well-known/oauth-protected-resource/mcp")
    if response.status_code == 404:
        response = await client.get("/.well-known/oauth-protected-resource")
    return response


@pytest.mark.asyncio
async def test_unauthenticated_post_to_mcp_endpoint_returns_401_with_metadata_challenge(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    mcp = _build_http_server(monkeypatch)
    app = mcp.http_app()
    transport = httpx.ASGITransport(app=app)

    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        response = await client.post(
            "/mcp",
            json={"jsonrpc": "2.0", "method": "tools/list", "id": 1},
            headers={
                "Accept": "application/json, text/event-stream",
                "Content-Type": "application/json",
            },
        )

    assert response.status_code == 401
    www_authenticate = response.headers.get("www-authenticate", "")
    assert www_authenticate.startswith("Bearer")
    # References the RFC 9728 discovery document so MCP clients can locate
    # the authorization server(s) that can issue a usable token.
    assert "resource_metadata=" in www_authenticate
    assert "/.well-known/oauth-protected-resource" in www_authenticate


@pytest.mark.asyncio
async def test_protected_resource_metadata_advertises_this_server_as_the_proxy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """RFC 9728 metadata is served and describes the OAuth-proxy shape.

    NOTE (fastmcp API surprise vs. a naive reading of the plan): the
    ``authorization_servers`` entry is this server's OWN ``base_url``, not
    ``login.microsoftonline.com``. That is correct for AzureProvider's
    OAuth-proxy pattern (see docs/plans/2026-07-12-multi-user-streamable-http.md,
    "Why AzureProvider + OBO"): MCP clients never talk to Entra directly --
    they authenticate against this proxy, which performs the Entra exchange
    server-side. A direct reference to login.microsoftonline.com would
    actually be wrong here. Verified against the installed fastmcp 3.4.4.
    """
    mcp = _build_http_server(monkeypatch)
    app = mcp.http_app()
    transport = httpx.ASGITransport(app=app)

    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        response = await _fetch_protected_resource_metadata(client)

    assert response.status_code == 200
    body = response.json()
    assert body["resource"] == "https://mcp.example.com/mcp"
    assert body["authorization_servers"] == ["https://mcp.example.com/"]
    assert "header" in body.get("bearer_methods_supported", [])


# ---------------------------------------------------------------------------
# 3. Per-user isolation at the OboTokenProvider / GraphClient seam
# ---------------------------------------------------------------------------


class _FakeAccessToken:
    """Duck-typed stand-in for fastmcp's AccessToken."""

    def __init__(self, token: str, claims: dict | None = None) -> None:
        self.token = token
        self.claims = claims or {}


class _FakeCredential:
    """Fake OnBehalfOfCredential bound to one user's Graph token."""

    def __init__(self, graph_token: str) -> None:
        self._graph_token = graph_token
        self.get_token_calls = 0

    async def get_token(self, *scopes: str) -> _FakeAccessToken:
        self.get_token_calls += 1
        return _FakeAccessToken(self._graph_token)


class _MultiUserAzureProvider:
    """Fake AzureProvider whose OBO exchange is keyed by user assertion.

    Mirrors the real ``AzureProvider.get_obo_credential``, which is an LRU of
    ``OnBehalfOfCredential`` keyed by the (hashed) user assertion -- i.e. one
    process-wide provider correctly hands back a *different* credential per
    caller instead of ever sharing one across users.
    """

    def __init__(self, credentials_by_assertion: dict[str, _FakeCredential]) -> None:
        self._by_assertion = credentials_by_assertion
        self.obo_assertions: list[str] = []

    async def get_obo_credential(self, user_assertion: str) -> _FakeCredential:
        self.obo_assertions.append(user_assertion)
        return self._by_assertion[user_assertion]


class _FakeServer:
    def __init__(self, auth: object) -> None:
        self.auth = auth


@pytest.mark.asyncio
async def test_shared_graph_client_never_bleeds_tokens_across_users(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """One shared GraphClient (as used by http mode's ``get_graph()``) must
    derive a fresh, correctly-scoped Bearer token per request from whichever
    identity is ambient *at that moment* -- never reusing or crossing a
    previous caller's token.
    """
    import fastmcp.server.dependencies as deps

    credential_a = _FakeCredential("graph-token-user-a")
    credential_b = _FakeCredential("graph-token-user-b")
    provider = _MultiUserAzureProvider(
        {
            "assertion-user-a": credential_a,
            "assertion-user-b": credential_b,
        }
    )
    server = _FakeServer(provider)

    ambient: dict[str, _FakeAccessToken] = {
        "token": _FakeAccessToken("assertion-user-a", claims={"oid": "user-a"})
    }
    monkeypatch.setattr(deps, "get_access_token", lambda: ambient["token"])
    monkeypatch.setattr(deps, "get_server", lambda: server)
    monkeypatch.setattr(identity_module, "_find_azure_provider", lambda auth: provider)

    seen_auth: list[str | None] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen_auth.append(request.headers.get("Authorization"))
        return httpx.Response(200, json={"value": []})

    mock_client = httpx.AsyncClient(transport=httpx.MockTransport(handler), timeout=5.0)
    monkeypatch.setattr(graph_module, "_request_client", mock_client)

    # A single shared instance, exactly like the process-wide
    # `_obo_graph_client` that every user's request funnels through in http
    # mode.
    shared_graph = GraphClient(token_provider=OboTokenProvider())

    try:
        await shared_graph.get("/me/mailFolders/inbox/messages")  # user A's request
        ambient["token"] = _FakeAccessToken("assertion-user-b", claims={"oid": "user-b"})
        await shared_graph.get("/me/mailFolders/inbox/messages")  # user B's request
    finally:
        await mock_client.aclose()

    assert seen_auth == ["Bearer graph-token-user-a", "Bearer graph-token-user-b"]
    assert provider.obo_assertions == ["assertion-user-a", "assertion-user-b"]
    # Each user's OBO credential was exercised exactly once -- no reuse of
    # the other user's cached credential.
    assert credential_a.get_token_calls == 1
    assert credential_b.get_token_calls == 1


# ---------------------------------------------------------------------------
# 4. stdio regression canary
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_stdio_mode_still_registers_profile_tools_and_wires_get_graph(
    monkeypatch: pytest.MonkeyPatch, tmp_path
) -> None:
    """Cheap end-to-end sanity check that the default (stdio) build is
    untouched by http-mode work: profile tools register, and
    ``graph.get_graph`` resolves through the same cached GraphClient
    ``ProfileManager.get_graph`` would hand back.

    Deeper coverage already exists elsewhere and is intentionally not
    repeated: the full stdio-vs-http tool-registration matrix lives in
    test_http_transport.py, and the ``get_graph`` dispatch logic in
    isolation (mocked ProfileManager) is unit-tested in
    test_identity.py::test_get_graph_stdio_mode_delegates_to_profile_manager.
    This test only proves those two units still integrate correctly against
    the real, unmocked default server build.
    """
    import mcp_microsoft.server as server_mod
    from mcp_microsoft.profiles import get_profile_manager

    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    monkeypatch.setenv("MS365_CLIENT_ID", "boot-client")
    monkeypatch.setenv("MS365_TENANT_ID", "common")
    for name in _HTTP_ENV_VARS:
        monkeypatch.delenv(name, raising=False)
    reset_runtime_state()

    mcp = server_mod.get_mcp_server(reset=True)
    names = {tool.name for tool in await mcp.list_tools(run_middleware=False)}

    assert "authenticate_ms_profile" in names
    assert "list_emails" in names

    assert graph_module.get_graph(None) is get_profile_manager().get_graph(None)
