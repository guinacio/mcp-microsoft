"""
Token-provider abstraction for Microsoft Graph authentication.

A :class:`TokenProvider` yields the OAuth access token a ``GraphClient``
attaches to each outbound Graph request.  Decoupling token acquisition from
the Graph client lets the same client serve different identity models:

* stdio mode uses :class:`ProfileTokenProvider`, which wraps the existing
  ``ProfileManager`` singleton (MSAL public client, encrypted disk cache).
* HTTP multi-user mode will introduce a per-request, On-Behalf-Of provider
  in a later phase; it plugs into this same seam with no changes to
  ``GraphClient`` or any tool.
"""

from __future__ import annotations

import asyncio
from dataclasses import dataclass
from typing import Protocol, runtime_checkable


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
        """Fetch a token for this provider's profile without blocking the loop."""
        # Lazy import mirrors graph.py's circular-import avoidance.
        from mcp_microsoft.profiles import get_profile_manager

        return await asyncio.to_thread(get_profile_manager().get_token, self.profile)
