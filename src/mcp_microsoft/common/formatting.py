from __future__ import annotations

from datetime import datetime
from typing import Any


def format_datetime_display(iso: str | None) -> str:
    """Format an ISO 8601 datetime string into a compact local-looking display."""
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return iso


def format_size_display(size_bytes: int) -> str:
    """Format a byte count into a human-readable string."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    return f"{size_bytes / (1024 * 1024):.1f} MB"


def drive_item_payload(item: dict[str, Any]) -> "DriveItemInfo":
    """Normalize a Graph DriveItem dict into a structured DriveItemInfo."""
    from mcp_microsoft.models import DriveItemInfo

    return DriveItemInfo(
        id=item.get("id", ""),
        name=item.get("name", "(unnamed)"),
        size_bytes=item.get("size", 0),
        size_display=format_size_display(item.get("size", 0)),
        last_modified_at=item.get("lastModifiedDateTime"),
        last_modified_at_display=format_datetime_display(item.get("lastModifiedDateTime")),
        web_url=item.get("webUrl", ""),
        is_folder="folder" in item,
        child_count=item.get("folder", {}).get("childCount", 0),
        mime_type=item.get("file", {}).get("mimeType", ""),
        parent_path=(item.get("parentReference") or {}).get("path", ""),
    )
