"""
Async httpx client wrapper for the Microsoft Graph API.

Each GraphClient instance is bound to a profile name.  Authentication
headers are fetched per-request via ProfileManager so MSAL can handle
transparent token refresh.

Usage:
    from mcp_microsoft.graph import get_graph
    g = get_graph("work")
    data = await g.get("/me/mailFolders/inbox/messages", params={"$top": 10})
"""

from __future__ import annotations

from typing import Any

import httpx

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class GraphClient:
    """
    Thin async wrapper around httpx for Microsoft Graph REST API calls.

    Each instance is optionally bound to a named profile.  The Bearer token
    is injected from ProfileManager.get_headers(profile) on every request.
    """

    def __init__(self, profile: str | None = None) -> None:
        self._profile = profile

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _get_headers(self) -> dict[str, str]:
        """Fetch authenticated headers for this client's profile."""
        from mcp_microsoft.profiles import ProfileManager

        return ProfileManager.get().get_headers(self._profile)

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
        headers = self._get_headers()

        # Merge caller-supplied headers (e.g. Content-Type for PUT)
        if "headers" in kwargs:
            merged = {**headers, **kwargs.pop("headers")}
        else:
            merged = headers

        async with httpx.AsyncClient(timeout=30.0) as client:
            response = await client.request(
                method=method,
                url=url,
                headers=merged,
                **kwargs,
            )

        # Surface rate-limit hints clearly
        if response.status_code == 429:
            retry_after = response.headers.get("Retry-After", "unknown")
            raise httpx.HTTPStatusError(
                f"Graph API rate limit exceeded. Retry after {retry_after} seconds.",
                request=response.request,
                response=response,
            )

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
            raise httpx.HTTPStatusError(msg, request=exc.request, response=exc.response) from None

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

    async def get_raw(self, path: str) -> bytes:
        """GET request that returns raw bytes (used for file downloads)."""
        url = f"{GRAPH_BASE}{path}"
        headers = self._get_headers()

        async with httpx.AsyncClient(timeout=120.0, follow_redirects=True) as client:
            response = await client.get(url, headers=headers)

        if response.status_code == 429:
            retry_after = response.headers.get("Retry-After", "unknown")
            raise httpx.HTTPStatusError(
                f"Graph API rate limit exceeded. Retry after {retry_after} seconds.",
                request=response.request,
                response=response,
            )

        response.raise_for_status()
        return response.content


def get_graph(profile: str | None = None) -> GraphClient:
    """
    Return a GraphClient for the given profile.

    Uses ProfileManager's cached instances so each profile
    gets a single reusable GraphClient.
    """
    from mcp_microsoft.profiles import ProfileManager

    return ProfileManager.get().get_graph(profile)
