from __future__ import annotations

from datetime import datetime

from mcp_microsoft.graph_types import GraphDriveItem


def format_datetime_display(iso: str | None) -> str:
    """Format an ISO 8601 datetime string into a compact local-looking display."""
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return iso


def format_size_display(size_bytes: int | None) -> str:
    """Format a byte count into a human-readable string."""
    if size_bytes is None:
        return "—"
    if size_bytes < 1024:
        return f"{size_bytes} B"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    return f"{size_bytes / (1024 * 1024):.1f} MB"


def drive_item_payload(item: GraphDriveItem) -> "DriveItemInfo":
    """Normalize a typed Graph DriveItem into a structured DriveItemInfo."""
    from mcp_microsoft.models import DriveItemInfo

    return DriveItemInfo(
        id=item.id,
        name=item.name or "(unnamed)",
        size_bytes=item.size or 0,
        size_display=format_size_display(item.size),
        last_modified_at=item.last_modified_date_time,
        last_modified_at_display=format_datetime_display(item.last_modified_date_time),
        web_url=item.web_url,
        is_folder=item.folder is not None,
        child_count=item.folder.child_count if item.folder else 0,
        mime_type=item.file.mime_type if item.file else "",
        parent_path=item.parent_reference.path if item.parent_reference else "",
    )
