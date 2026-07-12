from __future__ import annotations

import httpx
import pytest

import mcp_microsoft.graph as graph_module
import mcp_microsoft.profiles as profiles_module
from mcp_microsoft.graph import GraphClient
from mcp_microsoft.identity import ProfileTokenProvider


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
    seen_profiles: list[str | None] = []

    class FakeManager:
        def get_token(self, profile: str | None = None) -> str:
            seen_profiles.append(profile)
            return f"token-for-{profile}"

    monkeypatch.setattr(
        profiles_module, "get_profile_manager", lambda config=None: FakeManager()
    )

    assert await ProfileTokenProvider("work").get_access_token() == "token-for-work"
    # Default (None) profile is passed through unchanged.
    assert await ProfileTokenProvider().get_access_token() == "token-for-None"
    assert seen_profiles == ["work", None]


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
