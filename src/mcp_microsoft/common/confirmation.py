"""
Era-aware user confirmation for side-effecting tools.

MCP 2026-07-28 removed server-initiated requests from the wire (SEP-2322,
statelessness from SEP-2575), so ``ctx.elicit()`` only works on handshake-era
connections (protocol <= 2025-11-25). On the modern era a tool asks for input
by *returning* an ``InputRequiredResult``; the client collects the answer and
re-issues the same ``tools/call`` with it attached in ``inputResponses``.

:func:`request_confirmation` hides that split behind one call:

* handshake era -> ``ctx.elicit()`` with a boolean confirmation form;
* modern era    -> an ``InputRequiredResult`` on the first round, and the
  client's answer on the retry.

Both paths fail closed: a client that did not declare form elicitation, a
decline/cancel, a malformed answer, or an answer that does not belong to the
exact action being confirmed never yields :class:`ConfirmationGranted`.

Modern-era binding: the ``request_state`` this module mints carries a digest of
a caller-supplied *binding* string describing the concrete action (recipients,
message IDs, the number of messages a folder-wide operation will touch, ...).
The SDK seals ``request_state`` (AES-GCM) and binds it to the tool name,
arguments, authenticated principal, audience, and a TTL before it leaves the
server, so a client cannot mint or replay one. Re-checking the digest on the
answer round adds what the SDK cannot know: that server-side state the prompt
described (for example a folder's message count) has not changed since the
user saw it. A mismatch re-asks with the current details instead of
proceeding.
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from typing import Any

from mcp import types as mcp_types
from mcp.shared.exceptions import MCPError
from mcp_types.version import MODERN_PROTOCOL_VERSIONS

# Key of the single input request a confirmation round mints.
CONFIRMATION_INPUT_KEY = "confirm"

# Prefix + version of the request_state this module mints, so a future format
# change can never be confused with this one.
_STATE_PREFIX = "mcp-microsoft/confirm/v1:"

_CONFIRMATION_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "confirmed": {
            "type": "boolean",
            "title": "Confirm",
            "description": "Approve this action?",
        }
    },
    "required": ["confirmed"],
}

UNSUPPORTED_CLIENT_ERROR = (
    "confirm=True requires an MCP host that negotiated form elicitation. "
    "The action was not performed."
)
CANCELLED_ERROR = "Cancelled by user."


@dataclass(frozen=True)
class ConfirmationGranted:
    """The user explicitly approved this exact action."""


@dataclass(frozen=True)
class ConfirmationDenied:
    """The action must not run; ``reason`` is safe to return to the caller."""

    reason: str


ConfirmationOutcome = (
    ConfirmationGranted | ConfirmationDenied | mcp_types.InputRequiredResult
)


def _protocol_version(ctx: Any) -> str | None:
    """Return the protocol version this request negotiated, or None if unknown."""
    try:
        session = ctx.session
    except (AttributeError, RuntimeError):
        return None
    version = getattr(session, "protocol_version", None)
    return version if isinstance(version, str) else None


def is_modern_protocol(ctx: Any) -> bool:
    """Whether *ctx* is a sessionless (2026-07-28+) request."""
    return _protocol_version(ctx) in MODERN_PROTOCOL_VERSIONS


def supports_form_elicitation(ctx: Any) -> bool:
    """Whether the client declared form elicitation for this request.

    Reads ``session.client_capabilities``, which the SDK populates from the
    ``initialize`` handshake on older eras and from the per-request envelope
    (``_meta`` client capabilities) on 2026-07-28. An ``elicitation`` object
    with neither ``form`` nor ``url`` means form-only support per the
    2025-11-25 schema, so it counts as form support.
    """
    try:
        capabilities = ctx.session.client_capabilities
    except (AttributeError, RuntimeError):
        return False
    elicitation = getattr(capabilities, "elicitation", None)
    if elicitation is None:
        return False
    form = getattr(elicitation, "form", None)
    url = getattr(elicitation, "url", None)
    return form is not None or url is None


def _binding_state(binding: str) -> str:
    digest = hashlib.sha256(binding.encode("utf-8")).hexdigest()
    return f"{_STATE_PREFIX}{digest}"


def _ask(message: str, binding: str) -> mcp_types.InputRequiredResult:
    params = mcp_types.ElicitRequestFormParams(
        message=message,
        requested_schema=_CONFIRMATION_SCHEMA,
    )
    return mcp_types.InputRequiredResult(
        result_type="input_required",
        input_requests={
            CONFIRMATION_INPUT_KEY: mcp_types.ElicitRequest(
                method="elicitation/create", params=params
            )
        },
        request_state=_binding_state(binding),
    )


def _confirmed_content(answer: Any) -> bool:
    """True only for an accepted elicitation whose ``confirmed`` is literally True."""
    if not isinstance(answer, mcp_types.ElicitResult):
        return False
    if answer.action != "accept" or not isinstance(answer.content, dict):
        return False
    return answer.content.get("confirmed") is True


async def request_confirmation(
    ctx: Any | None,
    *,
    message: str,
    binding: str,
) -> ConfirmationOutcome:
    """Ask the user to approve an action, on whichever protocol era is live.

    Args:
        ctx: The tool's FastMCP ``Context`` (``None`` outside a request).
        message: The prompt shown to the user. Include every detail the user
            needs to decide (recipients, counts, irreversibility).
        binding: A string that changes whenever the action the prompt
            describes changes. Only its SHA-256 digest leaves the server.

    Returns:
        :class:`ConfirmationGranted` to proceed, :class:`ConfirmationDenied`
        to stop with a caller-safe reason, or an ``InputRequiredResult`` that
        the tool must return unchanged (modern era, asking round).
    """
    if ctx is None or not supports_form_elicitation(ctx):
        return ConfirmationDenied(UNSUPPORTED_CLIENT_ERROR)

    if is_modern_protocol(ctx):
        responses = ctx.input_responses
        if not responses or CONFIRMATION_INPUT_KEY not in responses:
            return _ask(message, binding)
        if ctx.request_state != _binding_state(binding):
            # The answer was given for a different action, or the target state
            # changed since the user saw the prompt: ask again with the current
            # details rather than act on a stale approval.
            return _ask(message, binding)
        if _confirmed_content(responses[CONFIRMATION_INPUT_KEY]):
            return ConfirmationGranted()
        return ConfirmationDenied(CANCELLED_ERROR)

    try:
        result = await ctx.elicit(
            message,
            response_type=bool,
            response_title="Confirm",
            response_description="Approve this action?",
        )
    except MCPError as exc:
        if exc.error.code != mcp_types.METHOD_NOT_FOUND:
            raise
        return ConfirmationDenied(
            "The MCP client advertised elicitation but did not implement it. "
            "The action was not performed."
        )
    if result.action == "accept" and result.data is True:
        return ConfirmationGranted()
    return ConfirmationDenied(CANCELLED_ERROR)
