from __future__ import annotations

from collections.abc import Callable
from typing import Any

from mcp.types import ToolAnnotations

READ_ONLY_TOOL = ToolAnnotations(readOnlyHint=True, openWorldHint=True)
WRITE_TOOL = ToolAnnotations(destructiveHint=False, openWorldHint=True)
IDEMPOTENT_WRITE_TOOL = ToolAnnotations(
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=True,
)
DESTRUCTIVE_TOOL = ToolAnnotations(destructiveHint=True, openWorldHint=True)

LOCAL_READ_TOOL = ToolAnnotations(
    readOnlyHint=True,
    idempotentHint=True,
    openWorldHint=False,
)
LOCAL_WRITE_TOOL = ToolAnnotations(destructiveHint=False, openWorldHint=False)
LOCAL_IDEMPOTENT_TOOL = ToolAnnotations(
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=False,
)
LOCAL_DESTRUCTIVE_TOOL = ToolAnnotations(destructiveHint=True, openWorldHint=False)

_TITLE_OVERRIDES = {
    "id": "ID",
    "onedrive": "OneDrive",
    "mcp": "MCP",
    "ms": "MS",
    "rsvp": "RSVP",
    "sharepoint": "SharePoint",
    "url": "URL",
}


def humanize_tool_title(name: str) -> str:
    """Convert a snake_case tool name into a display title for MCP clients."""
    words = []
    for part in name.split("_"):
        if not part:
            continue
        words.append(_TITLE_OVERRIDES.get(part.lower(), part.capitalize()))
    return " ".join(words)


def register_tool(
    server: Any,
    fn: Callable[..., Any],
    *,
    annotations: ToolAnnotations | dict[str, Any] | None = None,
    name: str | None = None,
    title: str | None = None,
    **kwargs: Any,
) -> Any:
    """Register a tool with a stable human-readable title."""
    tool_name = name or fn.__name__
    tool_title = title or humanize_tool_title(tool_name)
    return server.tool(
        name=tool_name,
        title=tool_title,
        annotations=annotations,
        **kwargs,
    )(fn)
