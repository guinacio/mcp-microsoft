"""
Attachment tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.

Implemented:
  - list_attachments
  - download_attachment
"""

from __future__ import annotations

import base64
import os
from typing import Optional

from mcp_microsoft.graph import get_graph
from mcp_microsoft.server import mcp

# ---------------------------------------------------------------------------
# list_attachments
# ---------------------------------------------------------------------------


@mcp.tool()
async def list_attachments(message_id: str, profile: str | None = None) -> str:
    """
    List all attachments on an email message.

    Args:
        message_id: The Graph message ID to list attachments for.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted table of attachments with name, size,
        content type, and attachment ID.
    """
    g = get_graph(profile)
    params = {
        "$select": "id,name,size,contentType,isInline",
    }

    result = await g.get(f"/me/messages/{message_id}/attachments", params=params)
    attachments = result.get("value", [])

    if not attachments:
        return "No attachments found on this message."

    lines = [f"## Attachments ({len(attachments)} found)\n"]
    lines.append("| Name | Size | Type | Inline | ID |")
    lines.append("|---|---|---|---|---|")

    for att in attachments:
        name = att.get("name", "(unnamed)")
        size_bytes = att.get("size", 0)
        size_str = _fmt_size(size_bytes)
        content_type = att.get("contentType", "unknown")
        is_inline = "Yes" if att.get("isInline") else "No"
        att_id = att.get("id", "")
        lines.append(f"| {name} | {size_str} | {content_type} | {is_inline} | `{att_id}` |")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# download_attachment
# ---------------------------------------------------------------------------


@mcp.tool()
async def download_attachment(
    message_id: str,
    attachment_id: str,
    save_path: Optional[str] = None,
    profile: str | None = None,
) -> str:
    """
    Download an email attachment, saving it to disk or returning raw base64.

    Args:
        message_id: The Graph message ID that contains the attachment.
        attachment_id: The attachment ID from list_attachments.
        save_path: Optional path to save the file. If this is a directory,
            the attachment's original filename is used inside that directory.
            If omitted, the base64-encoded content is returned directly.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        If save_path is provided: confirmation with the saved file path and size.
        If save_path is omitted: the base64-encoded content string with filename.
    """
    g = get_graph(profile)
    result = await g.get(f"/me/messages/{message_id}/attachments/{attachment_id}")

    att_name = result.get("name", "attachment")
    content_type = result.get("contentType", "application/octet-stream")
    content_bytes_b64: Optional[str] = result.get("contentBytes")

    if content_bytes_b64 is None:
        return (
            f"Attachment `{att_name}` has no downloadable content "
            "(it may be a reference/link attachment rather than a file attachment)."
        )

    if save_path is None:
        # Return base64 directly
        return (
            f"**Filename:** {att_name}\n"
            f"**Content-Type:** {content_type}\n"
            f"**Encoding:** base64\n\n"
            f"{content_bytes_b64}"
        )

    # Resolve final output path — sanitize remote filename to prevent traversal
    resolved_path = save_path
    if os.path.isdir(save_path):
        from pathlib import Path
        safe_name = Path(att_name).name  # strip directory components
        if not safe_name or safe_name.startswith("."):
            safe_name = "attachment"
        resolved_path = os.path.join(save_path, safe_name)

    raw_bytes = base64.b64decode(content_bytes_b64)

    with open(resolved_path, "wb") as fh:
        fh.write(raw_bytes)

    size_str = _fmt_size(len(raw_bytes))
    return (
        f"Attachment saved successfully.\n"
        f"**File:** `{resolved_path}`\n"
        f"**Filename:** {att_name}\n"
        f"**Size:** {size_str}\n"
        f"**Content-Type:** {content_type}"
    )


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------


def _fmt_size(size_bytes: int) -> str:
    """Format a byte count as a human-readable string."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
