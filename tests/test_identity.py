from __future__ import annotations

import httpx
import pytest

import mcp_microsoft.config as config_module
import mcp_microsoft.graph as graph_module
import mcp_microsoft.identity as identity_module
import mcp_microsoft.profiles as profiles_module
from mcp_microsoft.config import AppConfig
from mcp_microsoft.graph import GraphClient
from mcp_microsoft.identity import OboTokenProvider, ProfileTokenProvider


class _ConstantProvider:
    """Fake TokenProvider that always returns the same token."""

    def __init__(self, token: str) -> None:
        self._token = token
        self.calls = 0

    async def get_access_token(self) -> str:
        self.calls += 1
        return self._token


class _SequenceProvider:
    """Fake TokenProvider that returns a fresh token on each call."""

    def __init__(self, *tokens: str) -> None:
        self._tokens = list(tokens)
        self.calls = 0

    async def get_access_token(self) -> str:
        token = self._tokens[self.calls]
        self.calls += 1
        return token


@pytest.mark.asyncio
async def test_profile_token_provider_delegates_to_profile_manager(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """ProfileTokenProvider forwards the profile name to ProfileManager.get_token."""
    seen_calls: list[tuple[str | None, bool]] = []

    class FakeManager:
        def get_token(
            self, profile: str | None = None, *, allow_interactive: bool = True
        ) -> str:
            seen_calls.append((profile, allow_interactive))
            return f"token-for-{profile}"

    monkeypatch.setattr(
        profiles_module, "get_profile_manager", lambda config=None: FakeManager()
    )

    assert await ProfileTokenProvider("work").get_access_token() == "token-for-work"
    # Default (None) profile is passed through unchanged.
    assert await ProfileTokenProvider().get_access_token() == "token-for-None"
    # Tool-call paths must never fall into interactive/device-code auth: an
    # expired token raises immediately instead of blocking the tool call.
    assert seen_calls == [("work", False), (None, False)]


@pytest.mark.asyncio
async def test_graph_client_sends_bearer_from_injected_provider(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """An injected TokenProvider's token appears as the Authorization header."""
    captured: dict[str, str | None] = {}

    def handler(request: httpx.Request) -> httpx.Response:
        captured["auth"] = request.headers.get("Authorization")
        return httpx.Response(200, json={"ok": True})

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler), timeout=5.0)
    monkeypatch.setattr(graph_module, "_request_client", client)

    provider = _ConstantProvider("tok123")
    g = GraphClient(token_provider=provider)
    try:
        result = await g.get("/me")
    finally:
        await client.aclose()

    assert result == {"ok": True}
    assert captured["auth"] == "Bearer tok123"
    assert provider.calls == 1


@pytest.mark.asyncio
async def test_headers_rederived_per_request_attempt(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Headers are re-derived per attempt, so a mid-retry token refresh is used."""
    seen_auth: list[str | None] = []
    responses = iter(
        [
            httpx.Response(429, headers={"Retry-After": "0"}),
            httpx.Response(200, json={"ok": True}),
        ]
    )

    def handler(request: httpx.Request) -> httpx.Response:
        seen_auth.append(request.headers.get("Authorization"))
        return next(responses)

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler), timeout=5.0)
    monkeypatch.setattr(graph_module, "_request_client", client)

    provider = _SequenceProvider("first-token", "second-token")
    g = GraphClient(token_provider=provider)
    try:
        result = await g.get("/me")
    finally:
        await client.aclose()

    assert result == {"ok": True}
    # Provider invoked once per attempt (initial 429 + retry), fresh token each time.
    assert provider.calls == 2
    assert seen_auth == ["Bearer first-token", "Bearer second-token"]


# ---------------------------------------------------------------------------
# OboTokenProvider — per-user On-Behalf-Of exchange (http mode)
# ---------------------------------------------------------------------------


class _FakeAccessToken:
    """Duck-typed stand-in for fastmcp's AccessToken."""

    def __init__(self, token: str, claims: dict | None = None) -> None:
        self.token = token
        self.claims = claims or {}


class _FakeCredential:
    """Fake OnBehalfOfCredential: records scopes and yields a fixed Graph token."""

    def __init__(self, graph_token: str) -> None:
        self._graph_token = graph_token
        self.get_token_calls: list[tuple[str, ...]] = []

    async def get_token(self, *scopes: str) -> _FakeAccessToken:
        self.get_token_calls.append(scopes)
        return _FakeAccessToken(self._graph_token)


class _FakeAzureProvider:
    """Duck-typed AzureProvider exposing only get_obo_credential.

    The isinstance-based unwrap in ``_find_azure_provider`` is exercised
    separately in ``test_find_azure_provider_*``; the happy path patches
    ``_find_azure_provider`` so it focuses purely on the OBO recipe.
    """

    def __init__(self, credential: _FakeCredential) -> None:
        self._credential = credential
        self.obo_assertions: list[str] = []

    async def get_obo_credential(self, user_assertion: str) -> _FakeCredential:
        self.obo_assertions.append(user_assertion)
        return self._credential


class _FakeServer:
    def __init__(self, auth: object) -> None:
        self.auth = auth


def _patch_fastmcp_deps(monkeypatch, *, access_token, server) -> None:
    """Patch the fastmcp dependency functions OboTokenProvider imports lazily."""
    import fastmcp.server.dependencies as deps

    monkeypatch.setattr(deps, "get_access_token", lambda: access_token)
    monkeypatch.setattr(deps, "get_server", lambda: server)


@pytest.mark.asyncio
async def test_obo_provider_exchanges_ambient_token_for_graph_token(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Happy path: ambient bearer -> OBO user assertion -> Graph .default token."""
    credential = _FakeCredential("graph-token-xyz")
    provider = _FakeAzureProvider(credential)
    token = _FakeAccessToken("user-assertion-abc", claims={"oid": "user-1"})

    _patch_fastmcp_deps(
        monkeypatch, access_token=token, server=_FakeServer(provider)
    )
    # Unwrap logic is covered by test_find_azure_provider_*; isolate the recipe.
    monkeypatch.setattr(identity_module, "_find_azure_provider", lambda auth: provider)

    result = await OboTokenProvider().get_access_token()

    assert result == "graph-token-xyz"
    # The caller's bearer token is handed to Azure as the OBO user assertion.
    assert provider.obo_assertions == ["user-assertion-abc"]
    # The Graph token is requested with the .default resource scope.
    assert credential.get_token_calls == [
        ("https://graph.microsoft.com/.default",)
    ]


@pytest.mark.asyncio
async def test_obo_provider_raises_when_no_ambient_token(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """No validated access token -> actionable RuntimeError (no token leaked)."""
    _patch_fastmcp_deps(
        monkeypatch, access_token=None, server=_FakeServer(object())
    )

    with pytest.raises(RuntimeError, match="not authenticated"):
        await OboTokenProvider().get_access_token()


class _NotAzureAuth:
    """Stand-in auth provider that is not an AzureProvider."""


@pytest.mark.asyncio
async def test_obo_provider_raises_when_provider_not_azure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Non-Azure auth provider -> RuntimeError naming the actual type."""
    token = _FakeAccessToken("user-assertion")
    # Real _find_azure_provider runs here (not patched): _NotAzureAuth fails the
    # isinstance check and the error names the concrete provider type.
    _patch_fastmcp_deps(
        monkeypatch, access_token=token, server=_FakeServer(_NotAzureAuth())
    )

    with pytest.raises(RuntimeError, match="_NotAzureAuth"):
        await OboTokenProvider().get_access_token()


def _build_real_azure_provider():
    from fastmcp.server.auth.providers.azure import AzureProvider

    return AzureProvider(
        client_id="cid",
        client_secret="secret",
        tenant_id="11111111-1111-1111-1111-111111111111",
        base_url="https://mcp.example.com",
        required_scopes=["mcp-access"],
    )


def test_find_azure_provider_direct_multiauth_and_none() -> None:
    """_find_azure_provider returns the provider directly and unwraps MultiAuth."""
    from fastmcp.server.auth.auth import MultiAuth

    from mcp_microsoft.identity import _find_azure_provider

    azure = _build_real_azure_provider()

    # Direct AzureProvider is returned as-is.
    assert _find_azure_provider(azure) is azure

    # AzureProvider composed as MultiAuth.server is unwrapped.
    multi = MultiAuth(server=azure)
    assert _find_azure_provider(multi) is azure

    # Anything else yields None so the caller can raise a clear error.
    assert _find_azure_provider(object()) is None
    assert _find_azure_provider(None) is None


# ---------------------------------------------------------------------------
# get_graph — transport-mode dispatch
# ---------------------------------------------------------------------------


def test_get_graph_http_mode_returns_shared_obo_client(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """http mode returns one shared OBO-backed client; profile arg is ignored."""
    monkeypatch.setattr(
        config_module, "get_app_config", lambda: AppConfig(transport="http")
    )
    # ProfileManager must never be consulted in http mode.
    monkeypatch.setattr(
        profiles_module,
        "get_profile_manager",
        lambda *a, **k: (_ for _ in ()).throw(
            AssertionError("ProfileManager must not be used in http mode")
        ),
    )
    # Reset the module-level shared client for a deterministic first creation.
    monkeypatch.setattr(graph_module, "_obo_graph_client", None)

    first = graph_module.get_graph("some-profile")
    second = graph_module.get_graph("a-different-profile")

    # Same shared instance regardless of the (ignored) profile argument.
    assert first is second
    assert isinstance(first._token_provider, OboTokenProvider)


def test_get_graph_stdio_mode_delegates_to_profile_manager(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """stdio mode is unchanged: it delegates to ProfileManager.get_graph."""
    sentinel = object()
    seen: list[str | None] = []

    class FakeManager:
        def get_graph(self, profile: str | None = None) -> object:
            seen.append(profile)
            return sentinel

    monkeypatch.setattr(
        config_module, "get_app_config", lambda: AppConfig(transport="stdio")
    )
    monkeypatch.setattr(
        profiles_module, "get_profile_manager", lambda *a, **k: FakeManager()
    )

    result = graph_module.get_graph("work")

    assert result is sentinel
    assert seen == ["work"]
