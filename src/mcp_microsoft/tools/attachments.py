"""
Attachment tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.

Implemented:
  - list_attachments
  - download_attachment
"""

from __future__ import annotations

import base64
from pathlib import Path
from typing import Optional

from fastmcp.utilities.types import File
from mcp.types import ToolAnnotations

from mcp_microsoft.models import AttachmentInfo, DownloadAttachmentResponse, ListAttachmentsResponse
from mcp_microsoft.graph import get_graph

# ---------------------------------------------------------------------------
# list_attachments
# ---------------------------------------------------------------------------

_READ_ONLY = ToolAnnotations(readOnlyHint=True, openWorldHint=True)
_WRITE = ToolAnnotations(destructiveHint=False, openWorldHint=True)

async def list_attachments(message_id: str, profile: str | None = None) -> ListAttachmentsResponse:
    """
    List all attachments on an email message.

    Args:
        message_id: The Graph message ID to list attachments for.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured attachment metadata.
    """
    g = get_graph(profile)
    params = {
        "$select": "id,name,size,contentType,isInline",
    }

    result = await g.get(f"/me/messages/{message_id}/attachments", params=params)
    attachments = result.get("value", [])

    items: list[AttachmentInfo] = []
    for att in attachments:
        size_bytes = att.get("size", 0)
        items.append(
            AttachmentInfo(
                id=att.get("id", ""),
                name=att.get("name", "(unnamed)"),
                size_bytes=size_bytes,
                size_display=_fmt_size(size_bytes),
                content_type=att.get("contentType", "unknown"),
                is_inline=att.get("isInline", False),
            )
        )

    return ListAttachmentsResponse(message_id=message_id, count=len(items), attachments=items)


# ---------------------------------------------------------------------------
# download_attachment
# ---------------------------------------------------------------------------


async def download_attachment(
    message_id: str,
    attachment_id: str,
    save_path: Optional[Path] = None,
    profile: str | None = None,
) -> DownloadAttachmentResponse | File:
    """
    Download an email attachment, saving it to disk or returning a file payload.

    When running in Claude Desktop, always provide save_path (e.g. the user's
    Downloads folder) so the file is written to disk. Omitting save_path returns
    a FastMCP file object, which is only useful for programmatic embedding.

    Args:
        message_id: The Graph message ID that contains the attachment.
        attachment_id: The attachment ID from list_attachments.
        save_path: Path to save the file on the host machine. If this is a
            directory, the attachment's original filename is used inside it.
            Recommended: always provide a path for Claude Desktop usage.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        A FastMCP file when no save path is given, or structured file-save metadata.
    """
    g = get_graph(profile)
    result = await g.get(f"/me/messages/{message_id}/attachments/{attachment_id}")

    att_name = result.get("name", "attachment")
    content_type = result.get("contentType", "application/octet-stream")
    content_bytes_b64: Optional[str] = result.get("contentBytes")

    if content_bytes_b64 is None:
        return DownloadAttachmentResponse(
            success=False,
            action="download_attachment",
            message_id=message_id,
            attachment_id=attachment_id,
            filename=att_name,
            error="Attachment has no downloadable content.",
        )

    raw_bytes = base64.b64decode(content_bytes_b64)

    if save_path is None:
        file_format = Path(att_name).suffix.lstrip(".")
        if not file_format and "/" in content_type:
            file_format = content_type.split("/", 1)[1]
        return File(data=raw_bytes, format=file_format or None, name=att_name)

    # Resolve final output path — sanitize remote filename to prevent traversal
    resolved_path = save_path
    if save_path.is_dir():
        safe_name = Path(att_name).name  # strip directory components
        if not safe_name or safe_name.startswith("."):
            safe_name = "attachment"
        resolved_path = save_path / safe_name

    resolved_path.parent.mkdir(parents=True, exist_ok=True)
    with resolved_path.open("wb") as fh:
        fh.write(raw_bytes)

    size_str = _fmt_size(len(raw_bytes))
    return DownloadAttachmentResponse(
        success=True,
        action="download_attachment",
        message_id=message_id,
        attachment_id=attachment_id,
        path=str(resolved_path),
        filename=att_name,
        size_bytes=len(raw_bytes),
        size_display=size_str,
        content_type=content_type,
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


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------


def register(server) -> None:
    """Register all attachment tools with the given FastMCP server instance."""
    server.tool(annotations=_READ_ONLY)(list_attachments)
    server.tool(annotations=_WRITE)(download_attachment)
