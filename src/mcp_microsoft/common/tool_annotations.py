"""Automatic tool annotation inference for mcp-microsoft.

Post-registration pass that infers MCP ToolAnnotations (readOnlyHint,
destructiveHint, idempotentHint, openWorldHint) for every registered tool
based on its name prefix.

Usage::

    from mcp_microsoft.common.tool_annotations import apply_tool_annotations
    apply_tool_annotations(mcp)
"""

from __future__ import annotations

from fastmcp import FastMCP
from mcp.types import ToolAnnotations

# ---------------------------------------------------------------------------
# Prefix tables
# ---------------------------------------------------------------------------

# Tools that only read state and never modify the environment.
_READ_ONLY_PREFIXES = (
    "check_",
    "download_",
    "find_",
    "get_",
    "list_",
    "read_",
    "search_",
)

# Destructive operations that are NOT idempotent (a second call has additional
# effect — e.g. delete twice is an error or deletes more items).
_DESTRUCTIVE_NON_IDEMPOTENT_PREFIXES = (
    "bulk_delete_",
    "delete_",
    "purge_",
    "remove_",
)

# Destructive operations that ARE idempotent (a second call with the same
# arguments leaves the environment in the same state as the first).
_DESTRUCTIVE_IDEMPOTENT_PREFIXES = (
    "bulk_mark_",
    "bulk_move_",
    "mark_",
    "move_",
)

# Non-destructive (additive/replacement) idempotent operations.
_NON_DESTRUCTIVE_IDEMPOTENT_PREFIXES = (
    "create_",
    "set_",
    "update_",
    "upload_",
)

# Non-destructive, non-idempotent operations (each call has a new side-effect,
# e.g. sending a message appends to history).
_NON_DESTRUCTIVE_NON_IDEMPOTENT_PREFIXES = (
    "add_",
    "forward_",
    "reply_",
    "rsvp_",
    "send_",
)

# Operations that have no clearly safe default — apply a conservative default.
# trash_ is treated the same as delete_ in spirit (moves to trash = reversible
# but still mutating and non-idempotent a second time).
_TRASH_PREFIXES = (
    "bulk_trash_",
    "trash_",
)

# Teams tools use a compound prefix (teams_<verb>_*).  We strip the "teams_"
# namespace and re-classify by the embedded verb so that teams_get_channel,
# teams_list_chats, teams_send_message, etc. are annotated correctly.
_TEAMS_NAMESPACE = "teams_"

# authenticate_ / filter_ have no natural classification; they fall through
# to the default.

# ---------------------------------------------------------------------------
# Title humanization
# ---------------------------------------------------------------------------

# Namespaces that become a display prefix in the tool title.
_TITLE_NAMESPACES = ("teams_",)


def _humanize_title(name: str) -> str:
    """Convert a snake_case tool name to a human-readable Title Case string.

    Examples::

        read_email              → "Read Email"
        send_email              → "Send Email"
        teams_list_joined       → "Teams List Joined"
        search_sharepoint_sites → "Search SharePoint Sites"
        add_ms_profile          → "Add MS Profile"
        list_enabled_services   → "List Enabled Services"
    """
    # Word-level corrections for abbreviations and proper nouns.
    _WORD_FIX: dict[str, str] = {
        "ms": "MS",
        "sharepoint": "SharePoint",
        "onedrive": "OneDrive",
        "odata": "OData",
    }

    prefix = ""
    base = name
    for ns in _TITLE_NAMESPACES:
        if name.startswith(ns):
            prefix = ns.rstrip("_").capitalize() + " "
            base = name[len(ns):]
            break
    words = " ".join(
        _WORD_FIX.get(w, w.capitalize()) for w in base.split("_") if w
    )
    return f"{prefix}{words}".strip()


# ---------------------------------------------------------------------------
# Classification helpers
# ---------------------------------------------------------------------------


def _classify(name: str) -> ToolAnnotations:
    """Return inferred ToolAnnotations for a tool with the given *name*."""

    # Strip optional teams_ namespace so we can classify by the inner verb.
    effective = name[len(_TEAMS_NAMESPACE):] if name.startswith(_TEAMS_NAMESPACE) else name

    # Read-only
    if effective.startswith(_READ_ONLY_PREFIXES):
        return ToolAnnotations(
            readOnlyHint=True,
            destructiveHint=False,
            idempotentHint=True,   # read-only implies idempotent
            openWorldHint=True,
        )

    # Destructive & non-idempotent
    if effective.startswith(_DESTRUCTIVE_NON_IDEMPOTENT_PREFIXES) or effective.startswith(_TRASH_PREFIXES):
        return ToolAnnotations(
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=True,
        )

    # Destructive & idempotent
    if effective.startswith(_DESTRUCTIVE_IDEMPOTENT_PREFIXES):
        return ToolAnnotations(
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=True,
            openWorldHint=True,
        )

    # Non-destructive & idempotent
    if effective.startswith(_NON_DESTRUCTIVE_IDEMPOTENT_PREFIXES):
        return ToolAnnotations(
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=True,
        )

    # Non-destructive & non-idempotent
    if effective.startswith(_NON_DESTRUCTIVE_NON_IDEMPOTENT_PREFIXES):
        return ToolAnnotations(
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=False,
            openWorldHint=True,
        )

    # Default — unknown prefix; conservative baseline
    return ToolAnnotations(
        readOnlyHint=False,
        destructiveHint=False,
        idempotentHint=False,
        openWorldHint=True,
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def apply_tool_annotations(server: FastMCP) -> None:
    """Post-registration pass: infer and apply ToolAnnotations to every tool.

    Iterates the tools registered in *server* via ``_local_provider._components``
    and sets ``annotations`` on each ``FunctionTool`` component whose annotations
    have not already been set explicitly.

    Call this once after all tool-module imports have run, e.g. at the bottom of
    ``server.py``.
    """
    components = server._local_provider._components
    annotated = 0
    for key, component in components.items():
        if not key.startswith("tool:"):
            continue
        # Set a human-readable title if not already provided.
        if not component.title:
            component.title = _humanize_title(component.name)

        # Only infer when no explicit annotations have been provided.
        if component.annotations is not None:
            continue
        component.annotations = _classify(component.name)
        annotated += 1

    import logging
    logging.getLogger(__name__).debug(
        "apply_tool_annotations: annotated %d tool(s) registered on server %r",
        annotated,
        server.name,
    )
