"""
Async httpx client wrapper for the Microsoft Graph API.

Each GraphClient obtains its Bearer token from a TokenProvider, re-fetched
per request (and per retry attempt) so MSAL can handle transparent token
refresh.  The default provider wraps ProfileManager, preserving the original
profile-bound behavior.

Usage:
    from mcp_microsoft.graph import get_graph
    g = get_graph("work")
    data = await g.get("/me/mailFolders/inbox/messages", params={"$top": 10})
"""

from __future__ import annotations

import asyncio
import http.cookiejar
import logging
from collections.abc import Awaitable, Callable
from typing import Any

import httpx
from fastmcp.exceptions import ToolError

from mcp_microsoft.identity import OboTokenProvider, ProfileTokenProvider, TokenProvider

logger = logging.getLogger(__name__)


class GraphAuthorizationError(ToolError, httpx.HTTPStatusError):
    """A Graph 401/403 whose message must stay visible to the caller.

    In http mode the server runs with ``mask_error_details=True``, which
    replaces every non-``ToolError`` exception with a generic message — leaving
    a user whose consent expired or whose account lacks a permission with no
    way to self-serve. Auth denials are safe to surface (the body is Entra's
    own error text about the caller's request), so this class is *both* a
    ``ToolError`` (survives masking) and an ``httpx.HTTPStatusError`` (existing
    handlers, e.g. the Teams meeting-artifact translators, still catch it).
    """

    def __init__(
        self, message: str, *, request: httpx.Request, response: httpx.Response
    ) -> None:
        Exception.__init__(self, message)
        # FastMCPError contract (fastmcp reads this when logging the error).
        self.log_level = logging.ERROR
        # httpx.HTTPStatusError contract.
        self.request = request
        self.response = response


_AUTH_DENIED_HINTS = {
    401: (
        "Microsoft rejected the sign-in for this request (expired or revoked "
        "session). Disconnect and reconnect the Microsoft integration, then try "
        "again."
    ),
    403: (
        "The signed-in account is missing a required permission or admin "
        "consent for this operation. Ask your Microsoft 365 administrator to "
        "grant the missing Graph permission to the app registration."
    ),
}

_RETRY_STATUSES = frozenset({429, 503})
_MAX_RETRIES = 3
_DEFAULT_RETRY_AFTER = 5  # seconds, used when Retry-After header is absent
_MAX_RETRY_AFTER = 60     # cap so a misbehaving server cannot stall forever

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_REQUEST_TIMEOUT = 30.0
_TRANSFER_TIMEOUT = 120.0

_request_client: httpx.AsyncClient | None = None
_transfer_client: httpx.AsyncClient | None = None

# Shared, per-process GraphClient for http (multi-user) transport. Safe to
# share across all users because identity is derived per-request from the
# ambient auth context inside OboTokenProvider — nothing user-specific is
# stored on the client. Created lazily on first http-mode get_graph() call.
_obo_graph_client: GraphClient | None = None


class _NullCookieJar(http.cookiejar.CookieJar):
    """Cookie jar that silently drops every ``Set-Cookie``.

    Graph API calls never need cookies, and in http (multi-user) mode the two
    AsyncClients below are shared across every user. A default httpx client
    keeps a mutable cookie jar, so a ``Set-Cookie`` from Graph or a redirected
    download host would be stored once and replayed on a *different* user's
    request. Persisting nothing keeps the shared clients free of cross-user
    cookie state (harmless in stdio mode, so applied unconditionally).
    """

    def set_cookie(self, cookie: http.cookiejar.Cookie) -> None:
        return None


async def initialize_http_clients() -> None:
    """Initialize shared HTTP clients for Graph API traffic."""
    global _request_client, _transfer_client

    if _request_client is None:
        _request_client = httpx.AsyncClient(
            timeout=_REQUEST_TIMEOUT,
            cookies=_NullCookieJar(),
        )
    if _transfer_client is None:
        _transfer_client = httpx.AsyncClient(
            timeout=_TRANSFER_TIMEOUT,
            follow_redirects=True,
            cookies=_NullCookieJar(),
        )


async def close_http_clients() -> None:
    """Close any shared HTTP clients created for Graph API traffic."""
    global _request_client, _transfer_client

    request_client = _request_client
    transfer_client = _transfer_client
    _request_client = None
    _transfer_client = None

    if request_client is not None:
        await request_client.aclose()
    if transfer_client is not None:
        await transfer_client.aclose()


def get_request_http_client() -> httpx.AsyncClient | None:
    """Return the shared request client when the server lifespan is active."""
    return _request_client


def get_transfer_http_client() -> httpx.AsyncClient | None:
    """Return the shared transfer client when the server lifespan is active."""
    return _transfer_client


def _raise_for_status_enriched(response: httpx.Response) -> None:
    """Raise for a 4xx/5xx response with the Graph error body appended.

    401/403 raise :class:`GraphAuthorizationError` (a ``ToolError`` that
    survives http-mode error masking, still catchable as
    ``httpx.HTTPStatusError``) with an actionable hint appended; every other
    error status raises plain ``httpx.HTTPStatusError`` exactly as before.
    """
    try:
        response.raise_for_status()
    except httpx.HTTPStatusError as exc:
        # Append Graph error body for better diagnostics
        try:
            body = response.json()
            error_info = body.get("error", {})
            msg = f"{exc} | Graph error: {error_info.get('code')} — {error_info.get('message')}"
        except Exception:
            msg = str(exc)
        hint = _AUTH_DENIED_HINTS.get(response.status_code)
        if hint is not None:
            raise GraphAuthorizationError(
                f"{msg} | {hint}", request=exc.request, response=exc.response
            ) from None
        raise httpx.HTTPStatusError(msg, request=exc.request, response=exc.response) from None


class GraphClient:
    """
    Thin async wrapper around httpx for Microsoft Graph REST API calls.

    The Bearer token is supplied by a TokenProvider and re-fetched on every
    request.  When no provider is given, the client defaults to a
    ProfileTokenProvider bound to *profile*, preserving the original
    ProfileManager-backed behavior.
    """

    def __init__(
        self,
        profile: str | None = None,
        *,
        token_provider: TokenProvider | None = None,
    ) -> None:
        self._profile = profile
        self._token_provider: TokenProvider = token_provider or ProfileTokenProvider(profile)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    async def _get_headers(self) -> dict[str, str]:
        """Build authenticated headers from a freshly acquired access token."""
        return {
            "Authorization": f"Bearer {await self._token_provider.get_access_token()}",
            "Content-Type": "application/json",
        }

    async def _send_with_retry(
        self,
        send: Callable[[dict[str, str]], Awaitable[httpx.Response]],
        *,
        operation_name: str = "request",
    ) -> httpx.Response:
        """Execute an authenticated HTTP call with shared retry handling."""
        for attempt in range(1, _MAX_RETRIES + 1):
            response = await send(await self._get_headers())

            if response.status_code in _RETRY_STATUSES:
                raw_retry_after = response.headers.get("Retry-After")
                try:
                    wait = min(int(raw_retry_after), _MAX_RETRY_AFTER) if raw_retry_after else _DEFAULT_RETRY_AFTER
                except ValueError:
                    wait = _DEFAULT_RETRY_AFTER

                if attempt < _MAX_RETRIES:
                    logger.warning(
                        "Graph API returned %s on %s (attempt %d/%d); retrying in %ds.",
                        response.status_code,
                        operation_name,
                        attempt,
                        _MAX_RETRIES,
                        wait,
                    )
                    await asyncio.sleep(wait)
                    continue

                raise httpx.HTTPStatusError(
                    f"Graph API returned {response.status_code} after {_MAX_RETRIES} attempts."
                    f" Last Retry-After: {raw_retry_after or 'not provided'}.",
                    request=response.request,
                    response=response,
                )

            return response

        raise RuntimeError("Retry loop exited without returning or raising.")

    # ------------------------------------------------------------------
    # Internal request helper
    # ------------------------------------------------------------------

    async def _request(
        self,
        method: str,
        path: str,
        **kwargs: Any,
    ) -> Any:
        """
        Execute an authenticated HTTP request against the Graph API.

        Args:
            method: HTTP method string (GET, POST, PATCH, DELETE, PUT, ...).
            path: Graph API path starting with '/' (e.g. '/me/messages').
            **kwargs: Passed directly to httpx.AsyncClient.request().

        Returns:
            Parsed JSON response as a Python dict/list, or None for 204 No Content.

        Raises:
            httpx.HTTPStatusError: on 4xx/5xx responses (includes response body
                in the message so callers can surface Graph error details).
            RuntimeError: if auth token acquisition fails.
        """
        url = f"{GRAPH_BASE}{path}"

        # Merge caller-supplied headers (e.g. Content-Type for PUT).
        # Re-fetched on every attempt so a token refresh mid-retry is picked up.
        extra_headers: dict[str, str] = kwargs.pop("headers", {})

        client = get_request_http_client()

        async def _send(headers_from_auth: dict[str, str]) -> httpx.Response:
            headers = {**headers_from_auth, **extra_headers}
            if client is None:
                async with httpx.AsyncClient(timeout=_REQUEST_TIMEOUT) as ephemeral_client:
                    return await ephemeral_client.request(
                        method=method,
                        url=url,
                        headers=headers,
                        **kwargs,
                    )
            return await client.request(
                method=method,
                url=url,
                headers=headers,
                **kwargs,
            )

        response = await self._send_with_retry(_send, operation_name=method.upper())

        _raise_for_status_enriched(response)

        # 204 No Content — nothing to parse
        if response.status_code == 204:
            return None

        # 202 Accepted — used by copy operations. Return body + Location header.
        if response.status_code == 202:
            result: dict = {}
            try:
                result = response.json()
            except Exception:
                pass
            # Preserve the Location header (monitor URL for async operations)
            location = response.headers.get("Location", "")
            if location:
                result["_monitor_url"] = location
            return result or None

        return response.json()

    # ------------------------------------------------------------------
    # Convenience methods
    # ------------------------------------------------------------------

    async def get(
        self,
        path: str,
        params: dict[str, Any] | None = None,
        headers: dict[str, str] | None = None,
    ) -> Any:
        """GET request. Returns parsed JSON."""
        kwargs: dict[str, Any] = {"params": params}
        if headers:
            kwargs["headers"] = headers
        return await self._request("GET", path, **kwargs)

    async def post(
        self,
        path: str,
        json: Any = None,
        headers: dict[str, str] | None = None,
    ) -> Any:
        """POST request with JSON body. Returns parsed JSON."""
        kwargs: dict[str, Any] = {"json": json}
        if headers:
            kwargs["headers"] = headers
        return await self._request("POST", path, **kwargs)

    async def patch(
        self,
        path: str,
        json: Any = None,
        headers: dict[str, str] | None = None,
    ) -> Any:
        """PATCH request with JSON body. Returns parsed JSON."""
        kwargs: dict[str, Any] = {"json": json}
        if headers:
            kwargs["headers"] = headers
        return await self._request("PATCH", path, **kwargs)

    async def delete(self, path: str) -> None:
        """DELETE request. Returns None (Graph returns 204 on success)."""
        await self._request("DELETE", path)

    async def put(
        self,
        path: str,
        content: bytes,
        content_type: str = "application/octet-stream",
    ) -> Any:
        """PUT request with raw bytes body (used for OneDrive file uploads)."""
        return await self._request(
            "PUT",
            path,
            content=content,
            headers={"Content-Type": content_type},
        )

    async def batch(self, requests: list[dict]) -> list[dict]:
        """
        Execute Graph API requests via the $batch endpoint.

        Accepts up to any number of requests and automatically chunks them into
        groups of 20 (the Graph API limit per batch call).  All chunks are sent
        sequentially and their responses are combined into a single flat list.

        Each request dict must contain:
            "id"     - caller-assigned string identifier (must be unique within
                       the full list; duplicate IDs across chunks are safe because
                       each chunk is a separate HTTP call, but callers should keep
                       them unique for easy result look-up).
            "method" - HTTP method string, e.g. "GET", "POST", "DELETE".
            "url"    - Relative Graph path, e.g. "/me/messages/{id}/move".
                       Do NOT include the base URL.

        Optional keys per request dict:
            "body"    - dict to send as the JSON request body.
            "headers" - dict of additional per-request headers (e.g.
                        {"Content-Type": "application/json"} when body is set).

        Returns:
            A flat list of response dicts, one per input request, in the order
            they were returned by Graph (which may differ from input order).
            Each response dict has at minimum:
                "id"     - matches the "id" from the originating request.
                "status" - HTTP status code (int).
                "body"   - parsed JSON body dict, or {} when Graph returns none.

        Raises:
            httpx.HTTPStatusError: if the $batch POST itself fails (network or
                auth error).  Per-item failures inside a successful batch call
                are NOT raised — callers must inspect the "status" field of each
                response dict to detect them.
        """
        _BATCH_CHUNK = 20
        responses: list[dict] = []

        for offset in range(0, len(requests), _BATCH_CHUNK):
            chunk = requests[offset : offset + _BATCH_CHUNK]
            result = await self.post("/$batch", json={"requests": chunk})
            for resp in (result or {}).get("responses", []):
                responses.append({
                    "id": resp.get("id", ""),
                    "status": resp.get("status", 0),
                    "body": resp.get("body") or {},
                })

        return responses

    async def get_raw(self, path: str) -> bytes:
        """GET request that returns raw bytes (used for file downloads)."""
        url = f"{GRAPH_BASE}{path}"
        client = get_transfer_http_client()

        async def _send(headers_from_auth: dict[str, str]) -> httpx.Response:
            if client is None:
                async with httpx.AsyncClient(
                    timeout=_TRANSFER_TIMEOUT,
                    follow_redirects=True,
                ) as ephemeral_client:
                    return await ephemeral_client.get(url, headers=headers_from_auth)
            return await client.get(url, headers=headers_from_auth)

        response = await self._send_with_retry(_send, operation_name="get_raw")
        _raise_for_status_enriched(response)
        return response.content


def get_graph(profile: str | None = None) -> GraphClient:
    """
    Return a GraphClient for the current transport mode.

    stdio mode: uses ProfileManager's cached instances so each *profile* gets a
    single reusable GraphClient (unchanged behavior).

    http (multi-user) mode: returns a shared, lazily created GraphClient backed
    by :class:`~mcp_microsoft.identity.OboTokenProvider`. The *profile* argument
    is IGNORED — identity always comes from the caller's bearer token via the
    per-request On-Behalf-Of exchange, so one shared client serves every user.
    """
    from mcp_microsoft.config import get_app_config

    if get_app_config().transport == "http":
        global _obo_graph_client
        if _obo_graph_client is None:
            logger.debug(
                "http transport: using shared OBO-backed GraphClient; the "
                "profile argument is ignored (identity comes from the token)."
            )
            _obo_graph_client = GraphClient(token_provider=OboTokenProvider())
        return _obo_graph_client

    from mcp_microsoft.profiles import get_profile_manager

    return get_profile_manager().get_graph(profile)
