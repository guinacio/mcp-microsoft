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
from pathlib import Path
from typing import Any, Optional
from urllib.parse import quote

import httpx
from fastmcp.server.context import Context
from mcp.types import ToolAnnotations

from mcp_microsoft.models import (
    CreateDriveFolderResponse,
    DeleteDriveItemResponse,
    DownloadFileResponse,
    DriveItemDetailResponse,
    DriveItemInfo,
    ListDriveItemsResponse,
    MoveOrCopyItemResponse,
    SearchDriveResponse,
    UploadFileResponse,
)
from mcp_microsoft.graph import get_graph, get_transfer_http_client

# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

_4MB = 4 * 1024 * 1024  # Simple upload threshold
_READ_ONLY = ToolAnnotations(readOnlyHint=True, openWorldHint=True)
_WRITE = ToolAnnotations(destructiveHint=False, openWorldHint=True)
_DESTRUCTIVE = ToolAnnotations(destructiveHint=True, openWorldHint=True)


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


def _drive_item_payload(item: dict[str, Any]) -> DriveItemInfo:
    """Normalize a DriveItem into a structured payload."""
    return DriveItemInfo(
        id=item.get("id", ""),
        name=item.get("name", "(unnamed)"),
        size_bytes=item.get("size", 0),
        size_display=_fmt_size(item.get("size", 0)),
        last_modified_at=item.get("lastModifiedDateTime"),
        last_modified_at_display=_fmt_dt(item.get("lastModifiedDateTime")),
        web_url=item.get("webUrl", ""),
        is_folder="folder" in item,
        child_count=item.get("folder", {}).get("childCount", 0),
        mime_type=item.get("file", {}).get("mimeType", ""),
        parent_path=(item.get("parentReference") or {}).get("path", ""),
    )


# ---------------------------------------------------------------------------
# list_drive_items
# ---------------------------------------------------------------------------


async def list_drive_items(
    folder_id: Optional[str] = None,
    max_results: int = 25,
    profile: str | None = None,
) -> ListDriveItemsResponse:
    """
    List files and folders in OneDrive.

    Args:
        folder_id: Optional folder ID to list contents of. When omitted,
                   lists the root of the user's OneDrive.
        max_results: Maximum number of items to return (1-200). Defaults to 25.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured OneDrive item data.
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

    return ListDriveItemsResponse(
        folder_id=folder_id,
        count=len(items),
        items=[_drive_item_payload(item) for item in items],
        has_more=result.get("@odata.nextLink") is not None,
    )


# ---------------------------------------------------------------------------
# get_drive_item
# ---------------------------------------------------------------------------


async def get_drive_item(item_id: str, profile: str | None = None) -> DriveItemDetailResponse:
    """
    Get metadata for a specific OneDrive file or folder.

    Args:
        item_id: The DriveItem ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured item details.
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
    parent_ref = item.get("parentReference") or {}
    parent_path = parent_ref.get("path", "")

    # Creator / modifier
    created_by = ((item.get("createdBy") or {}).get("user") or {}).get("displayName", "")
    modified_by = ((item.get("lastModifiedBy") or {}).get("user") or {}).get("displayName", "")

    return DriveItemDetailResponse(
        id=item_id,
        name=name,
        type=item_type,
        size_bytes=item.get("size", 0),
        size_display=size,
        created_at=item.get("createdDateTime"),
        created_at_display=created,
        created_by=created_by,
        modified_at=item.get("lastModifiedDateTime"),
        modified_at_display=modified,
        modified_by=modified_by,
        path=parent_path,
        web_url=web_url,
        child_count=item.get("folder", {}).get("childCount", 0),
        mime_type=item.get("file", {}).get("mimeType", ""),
    )


# ---------------------------------------------------------------------------
# search_drive
# ---------------------------------------------------------------------------


async def search_drive(query: str, max_results: int = 10, profile: str | None = None) -> SearchDriveResponse:
    """
    Search for files and folders in OneDrive by name or content.

    Args:
        query: Search query string.
        max_results: Maximum number of results. Defaults to 10.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured search results.
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

    return SearchDriveResponse(query=query, count=len(items), items=[_drive_item_payload(item) for item in items])


# ---------------------------------------------------------------------------
# create_drive_folder
# ---------------------------------------------------------------------------


async def create_drive_folder(
    name: str,
    parent_folder_id: Optional[str] = None,
    profile: str | None = None,
) -> CreateDriveFolderResponse:
    """
    Create a new folder in OneDrive.

    Args:
        name: Name for the new folder.
        parent_folder_id: Optional parent folder ID. When omitted,
                          creates the folder at the OneDrive root.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured folder creation confirmation.
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

    return CreateDriveFolderResponse(
        success=True,
        action="create_drive_folder",
        folder_id=folder_id,
        name=folder_name,
        parent_folder_id=parent_folder_id,
        web_url=web_url,
    )


# ---------------------------------------------------------------------------
# upload_file
# ---------------------------------------------------------------------------


async def upload_file(
    local_path: Optional[Path] = None,
    parent_folder_id: Optional[str] = None,
    filename: Optional[str] = None,
    content_base64: Optional[str] = None,
    profile: str | None = None,
    ctx: Context | None = None,
) -> UploadFileResponse:
    """
    Upload a file to OneDrive.

    Two modes of operation:
    1. **Local file**: pass local_path pointing to a file on the MCP server host.
       Files under 4 MB use simple PUT; larger files use resumable upload.
    2. **Base64 content**: pass content_base64 with the file bytes encoded as
       base64, plus filename. Use this when local filesystem access is not
       available (e.g. container/sandbox environments).

    Args:
        local_path: Absolute path to the file on the host machine. Optional
                    when content_base64 is provided instead.
        parent_folder_id: Optional destination folder ID. Defaults to OneDrive root.
        filename: Filename in OneDrive. Required when using content_base64.
                  Defaults to the local file's name when using local_path.
        content_base64: Base64-encoded file content. Use as an alternative to
                        local_path when the file isn't on disk. Requires filename.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured upload confirmation.
    """
    import base64
    import tempfile

    g = get_graph(profile)
    temp_local_path: Path | None = None

    try:
        # Base64 fallback: decode into a generated temp file instead of trusting the caller's path.
        if (local_path is None or not local_path.is_file()) and content_base64:
            if not filename:
                return UploadFileResponse(success=False, action="upload_file", path=str(local_path), error="filename is required when using content_base64.")
            try:
                raw = base64.b64decode(content_base64, validate=True)
            except Exception as e:
                return UploadFileResponse(success=False, action="upload_file", path=str(local_path), error=f"Invalid base64: {e}")
            suffix = Path(filename).suffix
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(raw)
                temp_local_path = Path(tmp.name)
            local_path = temp_local_path

        if local_path is None:
            return UploadFileResponse(success=False, action="upload_file", error="Provide local_path or content_base64 with filename.")

        if not local_path.is_file():
            return UploadFileResponse(success=False, action="upload_file", path=str(local_path), error="File not found.")

        upload_name = filename or local_path.name
        encoded_name = quote(upload_name, safe="")  # percent-encode for URL path segment
        file_size = local_path.stat().st_size

        if file_size <= _4MB:
            # Simple PUT upload
            file_bytes = local_path.read_bytes()
            if parent_folder_id:
                path = f"/me/drive/items/{parent_folder_id}:/{encoded_name}:/content"
            else:
                path = f"/me/drive/root:/{encoded_name}:/content"

            if ctx:
                await ctx.info(f"Uploading {upload_name} ({_fmt_size(file_size)})...")
            result = await g.put(path, content=file_bytes)
            if ctx:
                await ctx.info("Upload complete.")
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
                return UploadFileResponse(success=False, action="upload_file", path=str(local_path), error="No upload URL returned.")

            result = await _upload_large_file(upload_url, local_path, file_size, ctx)

        item_id = (result or {}).get("id", "unknown")
        web_url = (result or {}).get("webUrl", "")
        size_str = _fmt_size(file_size)

        return UploadFileResponse(
            success=True,
            action="upload_file",
            filename=upload_name,
            size_bytes=file_size,
            size_display=size_str,
            file_id=item_id,
            web_url=web_url,
            parent_folder_id=parent_folder_id,
        )
    finally:
        if temp_local_path is not None:
            try:
                temp_local_path.unlink(missing_ok=True)
            except OSError:
                pass


async def _upload_large_file(
    upload_url: str,
    file_path: Path,
    total_size: int,
    ctx: Context | None = None,
) -> dict:
    """Upload a large file in chunks using a resumable upload session."""

    chunk_size = 10 * 1024 * 1024  # 10 MB chunks (must be multiple of 320 KB)
    result = {}
    shared_client = get_transfer_http_client()

    if ctx is not None:
        await ctx.report_progress(progress=0, total=total_size)

    async def _send_chunks(client: httpx.AsyncClient) -> dict:
        offset = 0
        with file_path.open("rb") as stream:
            while offset < total_size:
                chunk = stream.read(chunk_size)
                if not chunk:
                    break

                end = offset + len(chunk)
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
                    local_result = response.json()
                else:
                    local_result = {}

                if ctx is not None:
                    await ctx.report_progress(progress=end, total=total_size)

                offset = end

        return local_result

    if shared_client is None:
        async with httpx.AsyncClient(
            timeout=120.0,
            follow_redirects=True,
        ) as ephemeral_client:
            result = await _send_chunks(ephemeral_client)
    else:
        result = await _send_chunks(shared_client)

    return result


# ---------------------------------------------------------------------------
# download_file
# ---------------------------------------------------------------------------


async def download_file(
    item_id: str,
    destination_path: Path,
    profile: str | None = None,
) -> DownloadFileResponse:
    """
    Download a file from OneDrive to a local path.

    IMPORTANT: destination_path must be on the machine running this MCP
    server (the user's computer). Use an absolute path on the user's
    filesystem (e.g. their home directory or Downloads folder).

    Args:
        item_id: The DriveItem ID of the file to download.
        destination_path: Local path to save the file. If a directory is
                          given, the original filename from OneDrive is used.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured download confirmation.
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
    dest = destination_path
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
    return DownloadFileResponse(
        success=True,
        action="download_file",
        item_id=item_id,
        path=str(dest),
        filename=filename,
        size_bytes=len(content),
        size_display=size_str,
        expected_size_bytes=expected_size,
    )


# ---------------------------------------------------------------------------
# delete_drive_item
# ---------------------------------------------------------------------------


async def delete_drive_item(item_id: str, profile: str | None = None) -> DeleteDriveItemResponse:
    """
    Delete a file or folder from OneDrive.

    The item is moved to the OneDrive recycle bin and can be restored
    for a limited period.

    Args:
        item_id: The DriveItem ID to delete.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured delete confirmation.
    """
    g = get_graph(profile)
    await g.delete(f"/me/drive/items/{item_id}")
    return DeleteDriveItemResponse(success=True, action="delete_drive_item", item_id=item_id, soft_delete=True)


# ---------------------------------------------------------------------------
# move_or_copy_item
# ---------------------------------------------------------------------------


async def move_or_copy_item(
    item_id: str,
    destination_folder_id: str,
    new_name: Optional[str] = None,
    copy: bool = False,
    profile: str | None = None,
    ctx: Context = None,
) -> MoveOrCopyItemResponse:
    """
    Move or copy a OneDrive item to a different folder.

    Args:
        item_id: The DriveItem ID to move or copy.
        destination_folder_id: The target folder's DriveItem ID.
        new_name: Optional new name for the item at the destination.
        copy: When True, copy the item instead of moving. Defaults to False (move).
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured move/copy confirmation.
    """
    g = get_graph(profile)
    # Get the destination folder's driveId for correct parentReference
    dest_meta = await g.get(
        f"/me/drive/items/{destination_folder_id}",
        params={"$select": "id,parentReference"},
    )
    drive_id = ((dest_meta or {}).get("parentReference") or {}).get("driveId", "")

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
            shared_client = get_transfer_http_client()
            if ctx is not None:
                await ctx.report_progress(progress=0, total=30)

            async def _poll(client: httpx.AsyncClient) -> MoveOrCopyItemResponse | None:
                for _ in range(30):  # max ~60s wait
                    attempt = _ + 1
                    resp = await client.get(monitor_url)
                    if resp.status_code == 200:
                        copy_result = resp.json()
                        status = copy_result.get("status", "")
                        if status == "completed":
                            if ctx is not None:
                                await ctx.report_progress(progress=30, total=30)
                            resource_id = copy_result.get("resourceId", "unknown")
                            return MoveOrCopyItemResponse(
                                success=True,
                                action="copy_item",
                                item_id=item_id,
                                new_item_id=resource_id,
                                destination_folder_id=destination_folder_id,
                            )
                        elif status == "failed":
                            if ctx is not None:
                                await ctx.report_progress(progress=30, total=30)
                            error_msg = copy_result.get("error", {}).get("message", "Unknown error")
                            return MoveOrCopyItemResponse(
                                success=False,
                                action="copy_item",
                                item_id=item_id,
                                destination_folder_id=destination_folder_id,
                                error=error_msg,
                            )
                    if ctx is not None:
                        await ctx.report_progress(progress=attempt, total=30)
                    await asyncio.sleep(2)
                return None

            if shared_client is None:
                async with httpx.AsyncClient(
                    timeout=60.0,
                    follow_redirects=True,
                ) as ephemeral_client:
                    completed = await _poll(ephemeral_client)
            else:
                completed = await _poll(shared_client)

            if completed is not None:
                return completed

        return MoveOrCopyItemResponse(
            success=True,
            action="copy_item",
            item_id=item_id,
            destination_folder_id=destination_folder_id,
            status="in_progress",
        )

    else:
        # Move is a PATCH with parentReference
        payload = {"parentReference": parent_ref}
        if new_name:
            payload["name"] = new_name

        result = await g.patch(f"/me/drive/items/{item_id}", json=payload)

        new_id = (result or {}).get("id", item_id)
        item_name = (result or {}).get("name", "")
        return MoveOrCopyItemResponse(
            success=True,
            action="move_item",
            item_id=item_id,
            new_item_id=new_id,
            name=item_name or None,
            destination_folder_id=destination_folder_id,
        )


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------


def register(server) -> None:
    """Register all OneDrive tools with the given FastMCP server instance."""
    server.tool(annotations=_READ_ONLY)(list_drive_items)
    server.tool(annotations=_READ_ONLY)(get_drive_item)
    server.tool(annotations=_READ_ONLY)(search_drive)
    server.tool(annotations=_WRITE)(create_drive_folder)
    server.tool(annotations=_WRITE)(upload_file)
    server.tool(annotations=_WRITE)(download_file)
    server.tool(annotations=_DESTRUCTIVE)(delete_drive_item)
    server.tool(annotations=_WRITE)(move_or_copy_item)
