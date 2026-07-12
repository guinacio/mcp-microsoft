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
import logging
import httpx
from pathlib import Path
from typing import Any
from urllib.parse import quote

from fastmcp.server.context import Context

from pydantic import Field

from mcp_microsoft.common.formatting import drive_item_payload, format_datetime_display, format_size_display
from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.transfer import upload_large_file_via_session
from mcp_microsoft.common.tooling import DESTRUCTIVE_TOOL, READ_ONLY_TOOL, WRITE_TOOL, register_tool
from mcp_microsoft.config import get_app_config
from mcp_microsoft.feature_flags import is_deletion_disabled
from mcp_microsoft.graph_types import GraphDriveItem, graph_identity_display, parse_graph_collection
from mcp_microsoft.models import (
    CreateDriveFolderResponse,
    DeleteDriveItemResponse,
    DownloadFileResponse,
    DriveItemDetailResponse,
    ListDriveItemsResponse,
    MoveOrCopyItemResponse,
    SearchDriveResponse,
    UploadFileResponse,
)
from mcp_microsoft.graph import get_graph, get_transfer_http_client

_log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

_4MB = 4 * 1024 * 1024  # Simple upload threshold


class UploadFileInput(ToolRequestModel):
    local_path: Path | None = None
    parent_folder_id: str | None = None
    filename: str | None = None
    content_base64: str | None = None
    uploaded_file: str | None = None
    profile: str | None = None


class ListDriveItemsInput(ToolRequestModel):
    folder_id: str | None = None
    max_results: int = 25
    profile: str | None = None


class GetDriveItemInput(ToolRequestModel):
    item_id: str
    profile: str | None = None


class SearchDriveInput(ToolRequestModel):
    query: str
    max_results: int = 10
    profile: str | None = None


class CreateDriveFolderInput(ToolRequestModel):
    name: str
    parent_folder_id: str | None = None
    profile: str | None = None


class DownloadFileInput(ToolRequestModel):
    item_id: str
    destination_path: Path
    profile: str | None = None


class DeleteDriveItemInput(ToolRequestModel):
    item_id: str
    profile: str | None = None


class MoveOrCopyItemInput(ToolRequestModel):
    item_id: str
    destination_folder_id: str
    new_name: str | None = None
    copy_value: bool = Field(False, alias="copy")
    profile: str | None = None




# ---------------------------------------------------------------------------
# list_drive_items
# ---------------------------------------------------------------------------


async def list_drive_items(
    params: ListDriveItemsInput,
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
    g = get_graph(params.profile)
    query: dict[str, Any] = {
        "$top": params.max_results,
        "$select": "id,name,size,file,folder,lastModifiedDateTime,webUrl",
        "$orderby": "name",
    }

    if params.folder_id:
        path = f"/me/drive/items/{params.folder_id}/children"
    else:
        path = "/me/drive/root/children"

    result = await g.get(path, params=query)
    items = parse_graph_collection(result, GraphDriveItem)

    return ListDriveItemsResponse(
        folder_id=params.folder_id,
        count=len(items),
        items=[drive_item_payload(item) for item in items],
        has_more=result.get("@odata.nextLink") is not None,
    )


# ---------------------------------------------------------------------------
# get_drive_item
# ---------------------------------------------------------------------------


async def get_drive_item(params: GetDriveItemInput) -> DriveItemDetailResponse:
    """
    Get metadata for a specific OneDrive file or folder.

    Args:
        item_id: The DriveItem ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured item details.
    """
    g = get_graph(params.profile)
    query = {
        "$select": (
            "id,name,size,file,folder,lastModifiedDateTime,createdDateTime,"
            "webUrl,parentReference,createdBy,lastModifiedBy"
        ),
    }

    item = GraphDriveItem.model_validate(await g.get(f"/me/drive/items/{params.item_id}", params=query))

    return DriveItemDetailResponse(
        id=item.id or params.item_id,
        name=item.name or "(unnamed)",
        type="Folder" if item.folder else "File",
        size_bytes=item.size or 0,
        size_display=format_size_display(item.size),
        created_at=item.created_date_time,
        created_at_display=format_datetime_display(item.created_date_time),
        created_by=graph_identity_display(item.created_by),
        modified_at=item.last_modified_date_time,
        modified_at_display=format_datetime_display(item.last_modified_date_time),
        modified_by=graph_identity_display(item.last_modified_by),
        path=item.parent_reference.path if item.parent_reference else "",
        web_url=item.web_url,
        child_count=item.folder.child_count if item.folder else 0,
        mime_type=item.file.mime_type if item.file else "",
    )


# ---------------------------------------------------------------------------
# search_drive
# ---------------------------------------------------------------------------


async def search_drive(params: SearchDriveInput) -> SearchDriveResponse:
    """
    Search for files and folders in OneDrive by name or content.

    Args:
        query: Search string used by Graph search.
        max_results: Maximum number of results to return. Defaults to 10.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured search results.
    """
    g = get_graph(params.profile)
    query_params: dict[str, Any] = {
        "$top": params.max_results,
        "$select": "id,name,size,file,folder,lastModifiedDateTime,webUrl,parentReference",
    }

    safe_query = quote(params.query.replace("'", "''"), safe="")  # escape OData + URL-encode
    path = f"/me/drive/root/search(q='{safe_query}')"
    result = await g.get(path, params=query_params)
    items = parse_graph_collection(result, GraphDriveItem)

    return SearchDriveResponse(
        query=params.query,
        count=len(items),
        items=[drive_item_payload(item) for item in items],
    )


# ---------------------------------------------------------------------------
# create_drive_folder
# ---------------------------------------------------------------------------


async def create_drive_folder(
    params: CreateDriveFolderInput,
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
    g = get_graph(params.profile)
    payload = {
        "name": params.name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename",
    }

    if params.parent_folder_id:
        path = f"/me/drive/items/{params.parent_folder_id}/children"
    else:
        path = "/me/drive/root/children"

    created = GraphDriveItem.model_validate(await g.post(path, json=payload) or {})

    return CreateDriveFolderResponse(
        success=True,
        action="create_drive_folder",
        folder_id=created.id or "unknown",
        name=created.name or params.name,
        parent_folder_id=params.parent_folder_id,
        web_url=created.web_url,
    )


# ---------------------------------------------------------------------------
# upload_file
# ---------------------------------------------------------------------------


async def upload_file(
    params: UploadFileInput,
    ctx: Context | None = None,
) -> UploadFileResponse:
    """
    Upload a file to OneDrive.

    Three modes of operation:
    1. **Local file**: pass local_path pointing to a file on the MCP server host.
       Files under 4 MB use simple PUT; larger files use resumable upload.
    2. **Base64 content**: pass content_base64 with the file bytes encoded as
       base64, plus filename. Use this when local filesystem access is not
       available (e.g. container/sandbox environments).
    3. **Uploaded file**: pass uploaded_file with the name of a file previously
       uploaded via the file-upload UI (see list_files). The file's bytes never
       passed through the model's context window.

    Args:
        local_path: Optional local file path on the MCP host machine.
        parent_folder_id: Optional destination folder ID. Omit to upload to the OneDrive root.
        filename: Optional override for the uploaded filename. Required when using `content_base64`.
        content_base64: Optional base64-encoded file content when no local path is available.
        uploaded_file: Optional name of a file previously uploaded via the
            file-upload UI (see list_files). Mutually exclusive with local_path
            and content_base64.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured upload confirmation.
    """
    if params.local_path is not None and get_app_config().transport == "http":
        raise ValueError(
            "local_path is not available in multi-user http mode (the server's "
            "disk is not the caller's disk); use content_base64 instead."
        )

    import base64
    import tempfile

    # File-upload UI source: resolve the stored bytes by name. Mutually
    # exclusive with the local_path / content_base64 sources.
    uploaded_bytes: bytes | None = None
    default_upload_name: str | None = None
    if params.uploaded_file is not None:
        if params.local_path is not None or params.content_base64 is not None:
            raise ValueError(
                "uploaded_file cannot be combined with local_path or "
                "content_base64; provide exactly one file source."
            )
        from mcp_microsoft.uploads import resolve_uploaded_file

        uploaded_bytes, _content_type = resolve_uploaded_file(params.uploaded_file)
        default_upload_name = params.uploaded_file

    g = get_graph(params.profile)
    temp_local_path: Path | None = None
    local_path = params.local_path

    try:
        # File-upload UI: stream the resolved bytes through a temp file so the
        # existing small-PUT / chunked-session paths are reused unchanged.
        if uploaded_bytes is not None:
            name_for_suffix = params.filename or default_upload_name or "upload"
            suffix = Path(name_for_suffix).suffix
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(uploaded_bytes)
                temp_local_path = Path(tmp.name)
            local_path = temp_local_path

        # Base64 fallback: decode into a generated temp file instead of trusting the caller's path.
        elif (local_path is None or not local_path.is_file()) and params.content_base64:
            if not params.filename:
                return UploadFileResponse(success=False, action="upload_file", path=str(local_path), error="filename is required when using content_base64.")
            try:
                raw = base64.b64decode(params.content_base64, validate=True)
            except Exception as e:
                return UploadFileResponse(success=False, action="upload_file", path=str(local_path), error=f"Invalid base64: {e}")
            suffix = Path(params.filename).suffix
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(raw)
                temp_local_path = Path(tmp.name)
            local_path = temp_local_path

        if local_path is None:
            return UploadFileResponse(success=False, action="upload_file", error="Provide local_path or content_base64 with filename.")

        if not local_path.is_file():
            return UploadFileResponse(success=False, action="upload_file", path=str(local_path), error="File not found.")

        upload_name = params.filename or default_upload_name or local_path.name
        encoded_name = quote(upload_name, safe="")  # percent-encode for URL path segment
        file_size = local_path.stat().st_size

        if file_size <= _4MB:
            # Simple PUT upload
            file_bytes = local_path.read_bytes()
            if params.parent_folder_id:
                path = f"/me/drive/items/{params.parent_folder_id}:/{encoded_name}:/content"
            else:
                path = f"/me/drive/root:/{encoded_name}:/content"

            if ctx:
                await ctx.info(f"Uploading {upload_name} ({format_size_display(file_size)})...")
            result = GraphDriveItem.model_validate(await g.put(path, content=file_bytes) or {})
            if ctx:
                await ctx.info("Upload complete.")
        else:
            # Resumable upload session for large files
            if params.parent_folder_id:
                session_path = f"/me/drive/items/{params.parent_folder_id}:/{encoded_name}:/createUploadSession"
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

            result = GraphDriveItem.model_validate(
                await upload_large_file_via_session(upload_url, local_path, file_size, ctx) or {}
            )

        size_str = format_size_display(file_size)

        return UploadFileResponse(
            success=True,
            action="upload_file",
            filename=upload_name,
            size_bytes=file_size,
            size_display=size_str,
            file_id=result.id or "unknown",
            web_url=result.web_url,
            parent_folder_id=params.parent_folder_id,
        )
    finally:
        if temp_local_path is not None:
            try:
                temp_local_path.unlink(missing_ok=True)
            except OSError:
                pass


# ---------------------------------------------------------------------------
# download_file
# ---------------------------------------------------------------------------


async def download_file(
    params: DownloadFileInput,
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
    g = get_graph(params.profile)
    # Get item metadata first to know the filename
    item = GraphDriveItem.model_validate(await g.get(
        f"/me/drive/items/{params.item_id}",
        params={"$select": "id,name,size"},
    ) or {})
    filename = item.name or "download"
    expected_size = item.size or 0

    # Resolve output path — sanitize remote filename to prevent traversal
    dest = params.destination_path
    if dest.is_dir():
        safe_name = Path(filename).name  # strip directory components
        if not safe_name or safe_name.startswith("."):
            safe_name = "download"
        dest = dest / safe_name

    # Download content
    content = await g.get_raw(f"/me/drive/items/{params.item_id}/content")

    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(content)

    size_str = format_size_display(len(content))
    return DownloadFileResponse(
        success=True,
        action="download_file",
        item_id=params.item_id,
        path=str(dest),
        filename=filename,
        size_bytes=len(content),
        size_display=size_str,
        expected_size_bytes=expected_size,
    )


# ---------------------------------------------------------------------------
# delete_drive_item
# ---------------------------------------------------------------------------


async def delete_drive_item(params: DeleteDriveItemInput) -> DeleteDriveItemResponse:
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
    g = get_graph(params.profile)
    await g.delete(f"/me/drive/items/{params.item_id}")
    return DeleteDriveItemResponse(
        success=True,
        action="delete_drive_item",
        item_id=params.item_id,
        soft_delete=True,
    )


# ---------------------------------------------------------------------------
# move_or_copy_item
# ---------------------------------------------------------------------------


async def move_or_copy_item(
    params: MoveOrCopyItemInput,
    ctx: Context | None = None,
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
    g = get_graph(params.profile)
    # Get the destination folder's driveId for correct parentReference
    dest_meta = GraphDriveItem.model_validate(await g.get(
        f"/me/drive/items/{params.destination_folder_id}",
        params={"$select": "id,parentReference"},
    ) or {})
    drive_id = dest_meta.parent_reference.drive_id if dest_meta.parent_reference else ""

    parent_ref: dict[str, Any] = {"id": params.destination_folder_id}
    if drive_id:
        parent_ref["driveId"] = drive_id

    if params.copy_value:
        # Copy is a POST that returns 202 Accepted with a monitor URL
        payload: dict = {"parentReference": parent_ref}
        if params.new_name:
            payload["name"] = params.new_name

        result = await g.post(f"/me/drive/items/{params.item_id}/copy", json=payload)

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
                                item_id=params.item_id,
                                new_item_id=resource_id,
                                destination_folder_id=params.destination_folder_id,
                            )
                        elif status == "failed":
                            if ctx is not None:
                                await ctx.report_progress(progress=30, total=30)
                            error_msg = copy_result.get("error", {}).get("message", "Unknown error")
                            return MoveOrCopyItemResponse(
                                success=False,
                                action="copy_item",
                                item_id=params.item_id,
                                destination_folder_id=params.destination_folder_id,
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
            item_id=params.item_id,
            destination_folder_id=params.destination_folder_id,
            status="in_progress",
        )

    else:
        # Move is a PATCH with parentReference
        payload = {"parentReference": parent_ref}
        if params.new_name:
            payload["name"] = params.new_name

        result = GraphDriveItem.model_validate(
            await g.patch(f"/me/drive/items/{params.item_id}", json=payload) or {"id": params.item_id}
        )
        return MoveOrCopyItemResponse(
            success=True,
            action="move_item",
            item_id=params.item_id,
            new_item_id=result.id or params.item_id,
            name=result.name or None,
            destination_folder_id=params.destination_folder_id,
        )


def register(server) -> None:
    """Register all OneDrive tools with the given FastMCP server instance."""
    register_tool(server, list_drive_items, annotations=READ_ONLY_TOOL)
    register_tool(server, get_drive_item, annotations=READ_ONLY_TOOL)
    register_tool(server, search_drive, annotations=READ_ONLY_TOOL)
    register_tool(server, create_drive_folder, annotations=WRITE_TOOL)
    register_tool(server, upload_file, annotations=WRITE_TOOL)
    if get_app_config().transport == "http":
        _log.info(
            "download_file not registered (http transport; server disk is "
            "not the caller's disk)"
        )
    else:
        register_tool(server, download_file, annotations=WRITE_TOOL)
    register_tool(server, move_or_copy_item, annotations=WRITE_TOOL)
    if not is_deletion_disabled():
        register_tool(server, delete_drive_item, annotations=DESTRUCTIVE_TOOL)
