"""
Token-provider abstraction for Microsoft Graph authentication.

A :class:`TokenProvider` yields the OAuth access token a ``GraphClient``
attaches to each outbound Graph request.  Decoupling token acquisition from
the Graph client lets the same client serve different identity models:

* stdio mode uses :class:`ProfileTokenProvider`, which wraps the existing
  ``ProfileManager`` singleton (MSAL public client, encrypted disk cache).
* HTTP multi-user mode uses :class:`OboTokenProvider`, which mints a per-user
  Graph token via an Entra On-Behalf-Of exchange derived from the caller's
  validated bearer token.  It plugs into this same seam with no changes to
  ``GraphClient`` or any tool.
"""

from __future__ import annotations

import asyncio
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Protocol, runtime_checkable

from fastmcp.exceptions import ToolError

if TYPE_CHECKING:
    from fastmcp.server.auth.providers.azure import AzureProvider

# Microsoft Graph resource scope for the OBO exchange. ``.default`` resolves to
# every delegated Graph permission granted to the app registration for the
# calling user, matching the scopes advertised via ``additional_authorize_scopes``.
_GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default"

# Serializes first-time token acquisition per profile (None = default profile).
# Without this, two concurrent initial calls for the same profile could each
# launch an interactive / device-code prompt. Module-level: one lock per
# profile for the process lifetime.
_profile_token_locks: dict[str | None, asyncio.Lock] = {}


def _profile_token_lock(profile: str | None) -> asyncio.Lock:
    """Return the process-wide token-acquisition lock for *profile*."""
    lock = _profile_token_locks.get(profile)
    if lock is None:
        lock = asyncio.Lock()
        _profile_token_locks[profile] = lock
    return lock


@runtime_checkable
class TokenProvider(Protocol):
    """Supplies a Graph API access token for a single outbound request."""

    async def get_access_token(self) -> str:
        """Return a valid Bearer access token for Microsoft Graph."""
        ...


@dataclass
class ProfileTokenProvider:
    """Token provider backed by the ProfileManager singleton (stdio mode).

    Acquires tokens through the profile's MSAL ``PublicClientApplication``.
    ``ProfileManager.get_token`` is synchronous and may block on interactive
    or device-code auth, so it runs in a worker thread to keep the event loop
    responsive.
    """

    profile: str | None = None

    async def get_access_token(self) -> str:
        """Fetch a token for this provider's profile without blocking the loop.

        Passes ``allow_interactive=False`` so that an expired or missing token
        raises immediately with a user-facing message rather than blocking the
        tool call on a device code prompt that would only appear in the log.
        The user is directed to call ``authenticate_ms_profile`` explicitly,
        which runs the interactive / device code flow and surfaces the sign-in
        URL directly in the Claude chat.
        """
        # Lazy import mirrors graph.py's circular-import avoidance.
        from mcp_microsoft.profiles import get_profile_manager

        # Serialize per profile so concurrent first-time calls can't each
        # trigger interactive/device-code auth. Once a profile is authenticated
        # the MSAL silent (cached) path is fast, so the lock is barely held.
        async with _profile_token_lock(self.profile):
            pm = get_profile_manager()
            return await asyncio.to_thread(
                lambda: pm.get_token(self.profile, allow_interactive=False)
            )


def _find_azure_provider(auth: object) -> "AzureProvider | None":
    """Return the ``AzureProvider`` behind *auth*, unwrapping ``MultiAuth``.

    Replicates fastmcp's private ``_find_azure_provider`` using only public
    types.  The server's auth provider may be an ``AzureProvider`` directly, or
    an ``AzureProvider`` composed as the ``server`` of a ``MultiAuth`` (used when
    a server accepts tokens from several sources).  Any other shape yields
    ``None``, letting the caller raise an actionable error.
    """
    from fastmcp.server.auth.auth import MultiAuth
    from fastmcp.server.auth.providers.azure import AzureProvider

    if isinstance(auth, AzureProvider):
        return auth
    if isinstance(auth, MultiAuth) and isinstance(auth.server, AzureProvider):
        return auth.server
    return None


class OboTokenProvider:
    """Per-user Graph token provider via Entra On-Behalf-Of (http mode).

    In HTTP multi-user mode every Graph call must run under the identity of the
    caller whose bearer token FastMCP validated for the current request.  This
    provider reads that ambient access token, hands it to the server's
    ``AzureProvider`` as the OBO *user assertion*, and exchanges it for a
    Microsoft Graph token scoped to ``https://graph.microsoft.com/.default``
    (all delegated permissions configured on the app registration).

    A single shared instance is safe to reuse across every user and request:
    the identity is resolved fresh from the request context on each call and is
    never stored on the instance.

    This class holds no token cache of its own.  Per-user credential caching
    lives in ``AzureProvider.get_obo_credential`` (an LRU of
    ``OnBehalfOfCredential`` keyed by the SHA-256 of the user assertion), and
    azure-identity keeps its own internal access-token cache inside each
    credential — so repeated calls by the same user reuse cached Graph tokens
    without re-hitting Entra.
    """

    async def get_access_token(self) -> str:
        """Exchange the caller's ambient token for a Microsoft Graph token.

        Raises:
            RuntimeError: if the request carries no validated access token, or
                the configured auth provider is not an ``AzureProvider``.  The
                token itself is never included in any error message or log.
            ToolError: if the OBO exchange itself is refused by Entra (expired
                or revoked session, missing consent, Conditional Access) — a
                ``ToolError`` so the sanitized, actionable message survives
                http-mode error masking and reaches the caller.
        """
        # Imported here (not at module scope) so stdio mode never pays the
        # fastmcp[azure] import cost, and so tests can monkeypatch these
        # dependency functions on the fastmcp module.
        from fastmcp.server.dependencies import get_access_token, get_server

        access_token = get_access_token()
        if access_token is None:
            raise RuntimeError(
                "Cannot perform an On-Behalf-Of exchange: the request is not "
                "authenticated (no access token in the current context). Graph "
                "tools require a valid bearer token in http transport."
            )

        server_auth = get_server().auth
        azure_provider = _find_azure_provider(server_auth)
        if azure_provider is None:
            raise RuntimeError(
                "On-Behalf-Of exchange requires an AzureProvider as the server "
                f"auth provider; got {type(server_auth).__name__}."
            )

        credential = await azure_provider.get_obo_credential(
            user_assertion=access_token.token,
        )
        try:
            result = await credential.get_token(_GRAPH_DEFAULT_SCOPE)
        except Exception as exc:
            raise ToolError(_obo_failure_message(exc)) from exc
        return result.token


# Matches Entra error codes (e.g. "AADSTS65001") in exchange failure text.
_AADSTS_CODE_RE = re.compile(r"AADSTS\d+")

# Cap on the sanitized upstream detail included in an OBO failure message.
_OBO_DETAIL_MAX = 500


def _obo_failure_message(exc: Exception) -> str:
    """Build the user-facing message for a failed On-Behalf-Of exchange.

    The raw azure-identity error is Entra's own description of why *this
    caller's* exchange was refused — safe to relay (it never contains the user
    assertion or the client secret), but it is normalized here anyway: control
    characters collapse to spaces (audit-log hygiene when fastmcp logs the
    message) and the detail is capped at ``_OBO_DETAIL_MAX`` chars. A known
    ``AADSTS`` code is called out explicitly so operators can look it up.
    """
    detail = re.sub(r"[\x00-\x1f\x7f]+", " ", str(exc)).strip()[:_OBO_DETAIL_MAX]
    code_match = _AADSTS_CODE_RE.search(detail)
    code_part = f" (Entra error {code_match.group(0)})" if code_match else ""
    return (
        f"Could not obtain a Microsoft Graph token for your account{code_part}. "
        "Common causes: your session expired or was revoked, the app is missing "
        "admin consent for a Graph permission, or a Conditional Access policy "
        "requires re-authentication. Disconnect and reconnect the Microsoft "
        f"integration, then try again. Details: {detail}"
    )
