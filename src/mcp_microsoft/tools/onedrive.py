"""
OneDrive tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.
Endpoints live under /me/drive in Graph v1.0.

Implemented:
  - list_drive_items
  - get_drive_item
  - search_drive
  - create_folder
  - upload_file
  - download_file
  - delete_drive_item
  - move_or_copy_item
"""

from __future__ import annotations

import asyncio
import os
from pathlib import Path
from typing import Optional
from urllib.parse import quote

from mcp_microsoft.graph import get_graph
from mcp_microsoft.server import mcp

# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

_4MB = 4 * 1024 * 1024  # Simple upload threshold


def _fmt_size(size_bytes: int) -> str:
    """Format a byte count as a human-readable string."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"


def _fmt_dt(iso: Optional[str]) -> str:
    """Format an ISO 8601 datetime to a human-readable form."""
    if not iso:
        return ""
    try:
        from datetime import datetime
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return iso


def _fmt_item(item: dict) -> str:
    """Format a DriveItem as a single markdown line."""
    name = item.get("name", "(unnamed)")
    item_id = item.get("id", "")
    size = item.get("size", 0)
    modified = _fmt_dt(item.get("lastModifiedDateTime"))
    is_folder = "folder" in item

    if is_folder:
        child_count = item.get("folder", {}).get("childCount", 0)
        return (
            f"- **{name}/** ({child_count} items) | Modified: {modified}\n"
            f"  ID: `{item_id}`"
        )
    else:
        mime = item.get("file", {}).get("mimeType", "")
        return (
            f"- {name} ({_fmt_size(size)}, {mime}) | Modified: {modified}\n"
            f"  ID: `{item_id}`"
        )


# ---------------------------------------------------------------------------
# list_drive_items
# ---------------------------------------------------------------------------


@mcp.tool()
async def list_drive_items(
    folder_id: Optional[str] = None,
    max_results: int = 25,
    profile: str | None = None,
) -> str:
    """
    List files and folders in OneDrive.

    Args:
        folder_id: Optional folder ID to list contents of. When omitted,
                   lists the root of the user's OneDrive.
        max_results: Maximum number of items to return (1-200). Defaults to 25.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted list of files and folders.
    """
    g = get_graph(profile)
    params: dict = {
        "$top": max_results,
        "$select": "id,name,size,file,folder,lastModifiedDateTime,webUrl",
        "$orderby": "name",
    }

    if folder_id:
        path = f"/me/drive/items/{folder_id}/children"
    else:
        path = "/me/drive/root/children"

    result = await g.get(path, params=params)
    items = result.get("value", [])

    if not items:
        loc = f"folder `{folder_id}`" if folder_id else "OneDrive root"
        return f"No items found in {loc}."

    loc = f"folder `{folder_id}`" if folder_id else "OneDrive root"
    lines = [f"## OneDrive: {loc} ({len(items)} items)\n"]
    for item in items:
        lines.append(_fmt_item(item))
    lines.append("")

    next_link = result.get("@odata.nextLink")
    if next_link:
        lines.append("*More items available — use a higher `max_results` or paginate.*")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# get_drive_item
# ---------------------------------------------------------------------------


@mcp.tool()
async def get_drive_item(item_id: str, profile: str | None = None) -> str:
    """
    Get metadata for a specific OneDrive file or folder.

    Args:
        item_id: The DriveItem ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted item details.
    """
    g = get_graph(profile)
    params = {
        "$select": (
            "id,name,size,file,folder,lastModifiedDateTime,createdDateTime,"
            "webUrl,parentReference,createdBy,lastModifiedBy"
        ),
    }

    item = await g.get(f"/me/drive/items/{item_id}", params=params)

    name = item.get("name", "(unnamed)")
    item_type = "Folder" if "folder" in item else "File"
    size = _fmt_size(item.get("size", 0))
    created = _fmt_dt(item.get("createdDateTime"))
    modified = _fmt_dt(item.get("lastModifiedDateTime"))
    web_url = item.get("webUrl", "")

    # Parent path
    parent_ref = item.get("parentReference", {})
    parent_path = parent_ref.get("path", "")

    # Creator / modifier
    created_by = item.get("createdBy", {}).get("user", {}).get("displayName", "")
    modified_by = item.get("lastModifiedBy", {}).get("user", {}).get("displayName", "")

    lines = [
        f"## {name}",
        f"**Type:** {item_type}",
        f"**Size:** {size}",
        f"**Created:** {created}" + (f" by {created_by}" if created_by else ""),
        f"**Modified:** {modified}" + (f" by {modified_by}" if modified_by else ""),
        f"**Path:** {parent_path}",
        f"**ID:** `{item_id}`",
    ]

    if "folder" in item:
        child_count = item.get("folder", {}).get("childCount", 0)
        lines.append(f"**Children:** {child_count}")

    if "file" in item:
        mime = item.get("file", {}).get("mimeType", "")
        lines.append(f"**MIME Type:** {mime}")

    if web_url:
        lines.append(f"**Web URL:** {web_url}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# search_drive
# ---------------------------------------------------------------------------


@mcp.tool()
async def search_drive(query: str, max_results: int = 10, profile: str | None = None) -> str:
    """
    Search for files and folders in OneDrive by name or content.

    Args:
        query: Search query string.
        max_results: Maximum number of results. Defaults to 10.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted list of matching items.
    """
    g = get_graph(profile)
    params: dict = {
        "$top": max_results,
        "$select": "id,name,size,file,folder,lastModifiedDateTime,webUrl,parentReference",
    }

    safe_query = quote(query.replace("'", "''"), safe="")  # escape OData + URL-encode
    path = f"/me/drive/root/search(q='{safe_query}')"
    result = await g.get(path, params=params)
    items = result.get("value", [])

    if not items:
        return f"No items found matching **{query}**."

    lines = [f"## Search results for '{query}' ({len(items)} items)\n"]
    for item in items:
        parent_path = item.get("parentReference", {}).get("path", "")
        lines.append(_fmt_item(item))
        if parent_path:
            lines.append(f"  Path: {parent_path}")
    lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# create_drive_folder
# ---------------------------------------------------------------------------


@mcp.tool()
async def create_drive_folder(
    name: str,
    parent_folder_id: Optional[str] = None,
    profile: str | None = None,
) -> str:
    """
    Create a new folder in OneDrive.

    Args:
        name: Name for the new folder.
        parent_folder_id: Optional parent folder ID. When omitted,
                          creates the folder at the OneDrive root.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Confirmation string with the new folder ID.
    """
    g = get_graph(profile)
    payload = {
        "name": name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename",
    }

    if parent_folder_id:
        path = f"/me/drive/items/{parent_folder_id}/children"
    else:
        path = "/me/drive/root/children"

    result = await g.post(path, json=payload)

    folder_id = (result or {}).get("id", "unknown")
    folder_name = (result or {}).get("name", name)
    web_url = (result or {}).get("webUrl", "")

    lines = [
        f"Folder created successfully.",
        f"**Name:** {folder_name}",
        f"**Folder ID:** `{folder_id}`",
    ]
    if web_url:
        lines.append(f"**Web URL:** {web_url}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# upload_file
# ---------------------------------------------------------------------------


@mcp.tool()
async def upload_file(
    local_path: str,
    parent_folder_id: Optional[str] = None,
    filename: Optional[str] = None,
    profile: str | None = None,
) -> str:
    """
    Upload a local file to OneDrive.

    Files under 4 MB use the simple PUT upload. Larger files use a
    resumable upload session automatically.

    Args:
        local_path: Path to the local file to upload.
        parent_folder_id: Optional destination folder ID. Defaults to OneDrive root.
        filename: Optional filename in OneDrive. Defaults to the local file's name.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Confirmation string with the new file's ID and web URL.
    """
    g = get_graph(profile)
    local = Path(local_path)
    if not local.is_file():
        return f"File not found: `{local_path}`"

    upload_name = filename or local.name
    encoded_name = quote(upload_name, safe="")  # percent-encode for URL path segment
    file_size = local.stat().st_size
    file_bytes = local.read_bytes()

    if file_size <= _4MB:
        # Simple PUT upload
        if parent_folder_id:
            path = f"/me/drive/items/{parent_folder_id}:/{encoded_name}:/content"
        else:
            path = f"/me/drive/root:/{encoded_name}:/content"

        result = await g.put(path, content=file_bytes)
    else:
        # Resumable upload session for large files
        if parent_folder_id:
            session_path = f"/me/drive/items/{parent_folder_id}:/{encoded_name}:/createUploadSession"
        else:
            session_path = f"/me/drive/root:/{encoded_name}:/createUploadSession"

        session_payload = {
            "item": {
                "@microsoft.graph.conflictBehavior": "rename",
                "name": upload_name,
            }
        }
        session = await g.post(session_path, json=session_payload)
        upload_url = (session or {}).get("uploadUrl", "")

        if not upload_url:
            return "Failed to create upload session — no upload URL returned."

        result = await _upload_large_file(upload_url, file_bytes, file_size)

    item_id = (result or {}).get("id", "unknown")
    web_url = (result or {}).get("webUrl", "")
    size_str = _fmt_size(file_size)

    lines = [
        f"File uploaded successfully.",
        f"**Filename:** {upload_name}",
        f"**Size:** {size_str}",
        f"**File ID:** `{item_id}`",
    ]
    if web_url:
        lines.append(f"**Web URL:** {web_url}")

    return "\n".join(lines)


async def _upload_large_file(upload_url: str, data: bytes, total_size: int) -> dict:
    """Upload a large file in chunks using a resumable upload session."""
    import httpx

    chunk_size = 10 * 1024 * 1024  # 10 MB chunks (must be multiple of 320 KB)
    result = {}

    async with httpx.AsyncClient(timeout=120.0) as client:
        offset = 0
        while offset < total_size:
            end = min(offset + chunk_size, total_size)
            chunk = data[offset:end]
            content_range = f"bytes {offset}-{end - 1}/{total_size}"

            response = await client.put(
                upload_url,
                content=chunk,
                headers={
                    "Content-Range": content_range,
                    "Content-Length": str(len(chunk)),
                },
            )
            response.raise_for_status()

            if response.status_code in (200, 201):
                result = response.json()
            offset = end

    return result


# ---------------------------------------------------------------------------
# download_file
# ---------------------------------------------------------------------------


@mcp.tool()
async def download_file(
    item_id: str,
    destination_path: str,
    profile: str | None = None,
) -> str:
    """
    Download a file from OneDrive to a local path.

    Args:
        item_id: The DriveItem ID of the file to download.
        destination_path: Local path to save the file. If a directory is
                          given, the original filename from OneDrive is used.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Confirmation string with the saved file path and size.
    """
    g = get_graph(profile)
    # Get item metadata first to know the filename
    item = await g.get(
        f"/me/drive/items/{item_id}",
        params={"$select": "id,name,size"},
    )
    filename = item.get("name", "download")
    expected_size = item.get("size", 0)

    # Resolve output path — sanitize remote filename to prevent traversal
    dest = Path(destination_path)
    if dest.is_dir():
        safe_name = Path(filename).name  # strip directory components
        if not safe_name or safe_name.startswith("."):
            safe_name = "download"
        dest = dest / safe_name

    # Download content
    content = await g.get_raw(f"/me/drive/items/{item_id}/content")

    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(content)

    size_str = _fmt_size(len(content))
    return (
        f"File downloaded successfully.\n"
        f"**Saved to:** `{dest}`\n"
        f"**Filename:** {filename}\n"
        f"**Size:** {size_str}"
    )


# ---------------------------------------------------------------------------
# delete_drive_item
# ---------------------------------------------------------------------------


@mcp.tool()
async def delete_drive_item(item_id: str, profile: str | None = None) -> str:
    """
    Delete a file or folder from OneDrive.

    The item is moved to the OneDrive recycle bin and can be restored
    for a limited period.

    Args:
        item_id: The DriveItem ID to delete.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Confirmation string.
    """
    g = get_graph(profile)
    await g.delete(f"/me/drive/items/{item_id}")
    return f"Item `{item_id}` deleted (moved to OneDrive recycle bin)."


# ---------------------------------------------------------------------------
# move_or_copy_item
# ---------------------------------------------------------------------------


@mcp.tool()
async def move_or_copy_item(
    item_id: str,
    destination_folder_id: str,
    new_name: Optional[str] = None,
    copy: bool = False,
    profile: str | None = None,
) -> str:
    """
    Move or copy a OneDrive item to a different folder.

    Args:
        item_id: The DriveItem ID to move or copy.
        destination_folder_id: The target folder's DriveItem ID.
        new_name: Optional new name for the item at the destination.
        copy: When True, copy the item instead of moving. Defaults to False (move).
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Confirmation string.
    """
    g = get_graph(profile)
    # Get the destination folder's driveId for correct parentReference
    dest_meta = await g.get(
        f"/me/drive/items/{destination_folder_id}",
        params={"$select": "id,parentReference"},
    )
    drive_id = (dest_meta or {}).get("parentReference", {}).get("driveId", "")

    parent_ref: dict = {"id": destination_folder_id}
    if drive_id:
        parent_ref["driveId"] = drive_id

    if copy:
        # Copy is a POST that returns 202 Accepted with a monitor URL
        payload: dict = {"parentReference": parent_ref}
        if new_name:
            payload["name"] = new_name

        result = await g.post(f"/me/drive/items/{item_id}/copy", json=payload)

        # Poll the monitor URL if available (injected by GraphClient as _monitor_url)
        monitor_url = ""
        if isinstance(result, dict):
            monitor_url = result.get("_monitor_url", "")

        if monitor_url:
            import httpx
            async with httpx.AsyncClient(timeout=60.0) as client:
                for _ in range(30):  # max ~60s wait
                    resp = await client.get(monitor_url)
                    if resp.status_code == 200:
                        copy_result = resp.json()
                        status = copy_result.get("status", "")
                        if status == "completed":
                            resource_id = copy_result.get("resourceId", "unknown")
                            return (
                                f"Item copied successfully.\n"
                                f"**New Item ID:** `{resource_id}`"
                            )
                        elif status == "failed":
                            error_msg = copy_result.get("error", {}).get("message", "Unknown error")
                            return f"Copy failed: {error_msg}"
                    await asyncio.sleep(2)

        return f"Item `{item_id}` copy initiated to folder `{destination_folder_id}`."

    else:
        # Move is a PATCH with parentReference
        payload = {"parentReference": parent_ref}
        if new_name:
            payload["name"] = new_name

        result = await g.patch(f"/me/drive/items/{item_id}", json=payload)

        new_id = (result or {}).get("id", item_id)
        item_name = (result or {}).get("name", "")
        name_str = f" (**{item_name}**)" if item_name else ""
        return (
            f"Item moved successfully{name_str}.\n"
            f"**Item ID:** `{new_id}`\n"
            f"**Destination:** folder `{destination_folder_id}`"
        )
