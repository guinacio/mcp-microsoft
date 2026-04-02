"""
SharePoint tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.
Endpoints live under /sites in Graph v1.0.

Note: SharePoint is only available with work/organizational Microsoft 365
accounts. Personal Outlook.com/Live accounts do not support SharePoint.

Implemented:
  - search_sharepoint_sites
  - get_sharepoint_site
  - list_site_libraries
  - list_site_files
  - get_site_file
  - upload_to_site
  - download_from_site
  - list_site_lists
  - get_list_items
  - create_list_item
  - update_list_item
  - delete_list_item
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Optional
from urllib.parse import quote

import httpx
from fastmcp.server.context import Context
from mcp.types import ToolAnnotations

from mcp_microsoft.models import (
    CreateListItemResponse,
    DeleteListItemResponse,
    DownloadSiteFileResponse,
    DriveItemInfo,
    GetListItemsResponse,
    ListSiteFilesResponse,
    ListSiteLibrariesResponse,
    ListSiteListsResponse,
    SearchSharePointSitesResponse,
    SharePointFields,
    SharePointLibraryInfo,
    SharePointListInfo,
    SharePointListItemInfo,
    SharePointSiteDetailResponse,
    SharePointSiteInfo,
    SiteFileDetailResponse,
    UpdateListItemResponse,
    UploadSiteFileResponse,
)
from mcp_microsoft.graph import get_graph, get_transfer_http_client
from mcp_microsoft.profiles import ProfileManager
from mcp_microsoft.server import mcp

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


def _fmt_drive_item(item: dict) -> str:
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


def _fmt_site(site: dict) -> str:
    """Format a SharePoint site as a markdown block."""
    name = site.get("displayName", "(unnamed)")
    site_id = site.get("id", "")
    web_url = site.get("webUrl", "")
    desc = site.get("description", "")
    lines = [f"- **{name}**"]
    if desc:
        lines.append(f"  {desc}")
    if web_url:
        lines.append(f"  URL: {web_url}")
    lines.append(f"  ID: `{site_id}`")
    return "\n".join(lines)


def _site_payload(site: dict[str, Any]) -> SharePointSiteInfo:
    """Normalize a SharePoint site into a structured payload."""
    return SharePointSiteInfo(
        id=site.get("id", ""),
        display_name=site.get("displayName", "(unnamed)"),
        description=site.get("description", ""),
        web_url=site.get("webUrl", ""),
        created_at=site.get("createdDateTime"),
        created_at_display=_fmt_dt(site.get("createdDateTime")),
        last_modified_at=site.get("lastModifiedDateTime"),
        last_modified_at_display=_fmt_dt(site.get("lastModifiedDateTime")),
    )


def _drive_item_payload(item: dict[str, Any]) -> DriveItemInfo:
    """Normalize a SharePoint DriveItem into a structured payload."""
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
        parent_path=item.get("parentReference", {}).get("path", ""),
    )


def _get_sharepoint_graph(profile: str | None):
    """Resolve a profile and return a Graph client with a clearer consumer-tenant error."""
    cfg = ProfileManager.get().resolve_profile(profile)
    if cfg.tenant_id == "consumers":
        raise ValueError(
            "SharePoint tools require a work or school Microsoft 365 account. "
            "Use a profile configured for an organization tenant."
        )
    return get_graph(profile)


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
# search_sharepoint_sites
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def search_sharepoint_sites(
    query: str = "",
    max_results: int = 25,
    profile: str | None = None,
) -> SearchSharePointSitesResponse:
    """
    Search SharePoint sites the user has access to.

    When query is empty, uses a wildcard search to discover accessible sites.
    Requires a work/organizational Microsoft 365 account.

    Args:
        query: Search query string. Leave empty to discover accessible sites via wildcard.
        max_results: Maximum number of sites to return (1-200). Defaults to 25.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured SharePoint site results.
    """
    g = _get_sharepoint_graph(profile)
    search_term = query if query else "*"
    params: dict = {
        "search": search_term,
        "$top": max_results,
        "$select": "id,displayName,description,webUrl",
    }

    result = await g.get("/sites", params=params)
    sites = result.get("value", [])

    return SearchSharePointSitesResponse(
        query=query,
        count=len(sites),
        sites=[_site_payload(site) for site in sites],
        has_more=result.get("@odata.nextLink") is not None,
    )


# ---------------------------------------------------------------------------
# get_sharepoint_site
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def get_sharepoint_site(
    site_id: str,
    profile: str | None = None,
) -> SharePointSiteDetailResponse:
    """
    Get details of a specific SharePoint site.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID (e.g. 'contoso.sharepoint.com,site-guid,web-guid').
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured site details.
    """
    g = _get_sharepoint_graph(profile)
    params = {
        "$select": "id,displayName,description,webUrl,createdDateTime,lastModifiedDateTime",
    }

    site = await g.get(f"/sites/{site_id}", params=params)

    name = site.get("displayName", "(unnamed)")
    desc = site.get("description", "")
    web_url = site.get("webUrl", "")
    created = _fmt_dt(site.get("createdDateTime"))
    modified = _fmt_dt(site.get("lastModifiedDateTime"))

    return SharePointSiteDetailResponse(
        id=site_id,
        display_name=name,
        description=desc,
        created_at=site.get("createdDateTime"),
        created_at_display=created,
        last_modified_at=site.get("lastModifiedDateTime"),
        last_modified_at_display=modified,
        web_url=web_url,
    )


# ---------------------------------------------------------------------------
# list_site_libraries
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def list_site_libraries(
    site_id: str,
    profile: str | None = None,
) -> ListSiteLibrariesResponse:
    """
    List document libraries (drives) in a SharePoint site.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured document library metadata.
    """
    g = _get_sharepoint_graph(profile)
    params = {
        "$select": "id,name,description,driveType,webUrl",
    }

    result = await g.get(f"/sites/{site_id}/drives", params=params)
    drives = result.get("value", [])

    return ListSiteLibrariesResponse(
        site_id=site_id,
        count=len(drives),
        libraries=[
            SharePointLibraryInfo(
                id=drive.get("id", ""),
                name=drive.get("name", "(unnamed)"),
                description=drive.get("description", ""),
                drive_type=drive.get("driveType", ""),
                web_url=drive.get("webUrl", ""),
            )
            for drive in drives
        ],
    )


# ---------------------------------------------------------------------------
# list_site_files
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def list_site_files(
    site_id: str,
    drive_id: str,
    folder_id: Optional[str] = None,
    max_results: int = 25,
    profile: str | None = None,
) -> ListSiteFilesResponse:
    """
    List files and folders in a SharePoint document library.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        drive_id: The document library (drive) ID.
        folder_id: Optional folder ID to list contents of. When omitted,
                   lists the root of the document library.
        max_results: Maximum number of items to return (1-200). Defaults to 25.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured document-library item data.
    """
    g = _get_sharepoint_graph(profile)
    params: dict = {
        "$top": max_results,
        "$select": "id,name,size,file,folder,lastModifiedDateTime,webUrl",
        "$orderby": "name",
    }

    if folder_id:
        path = f"/drives/{drive_id}/items/{folder_id}/children"
    else:
        path = f"/drives/{drive_id}/root/children"

    result = await g.get(path, params=params)
    items = result.get("value", [])

    return ListSiteFilesResponse(
        site_id=site_id,
        drive_id=drive_id,
        folder_id=folder_id,
        count=len(items),
        items=[_drive_item_payload(item) for item in items],
        has_more=result.get("@odata.nextLink") is not None,
    )


# ---------------------------------------------------------------------------
# get_site_file
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def get_site_file(
    site_id: str,
    drive_id: str,
    item_id: str,
    profile: str | None = None,
) -> SiteFileDetailResponse:
    """
    Get metadata for a file or folder in a SharePoint document library.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        drive_id: The document library (drive) ID.
        item_id: The DriveItem ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured file/folder details.
    """
    g = _get_sharepoint_graph(profile)
    params = {
        "$select": (
            "id,name,size,file,folder,lastModifiedDateTime,createdDateTime,"
            "webUrl,parentReference,createdBy,lastModifiedBy"
        ),
    }

    item = await g.get(
        f"/drives/{drive_id}/items/{item_id}",
        params=params,
    )

    name = item.get("name", "(unnamed)")
    item_type = "Folder" if "folder" in item else "File"
    size = _fmt_size(item.get("size", 0))
    created = _fmt_dt(item.get("createdDateTime"))
    modified = _fmt_dt(item.get("lastModifiedDateTime"))
    web_url = item.get("webUrl", "")

    parent_ref = item.get("parentReference", {})
    parent_path = parent_ref.get("path", "")

    created_by = item.get("createdBy", {}).get("user", {}).get("displayName", "")
    modified_by = item.get("lastModifiedBy", {}).get("user", {}).get("displayName", "")

    return SiteFileDetailResponse(
        site_id=site_id,
        drive_id=drive_id,
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
        child_count=item.get("folder", {}).get("childCount", 0),
        mime_type=item.get("file", {}).get("mimeType", ""),
        web_url=web_url,
    )


# ---------------------------------------------------------------------------
# upload_to_site
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_WRITE)
async def upload_to_site(
    site_id: str,
    drive_id: str,
    local_path: Path,
    folder_id: Optional[str] = None,
    filename: Optional[str] = None,
    profile: str | None = None,
    ctx: Context = None,
) -> UploadSiteFileResponse:
    """
    Upload a local file to a SharePoint document library.

    Files under 4 MB use the simple PUT upload. Larger files use a
    resumable upload session automatically.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        drive_id: The document library (drive) ID.
        local_path: Path to the local file to upload.
        folder_id: Optional destination folder ID within the library.
                   Defaults to the library root.
        filename: Optional filename in SharePoint. Defaults to the local file's name.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured upload confirmation.
    """
    g = _get_sharepoint_graph(profile)
    if not local_path.is_file():
        return UploadSiteFileResponse(success=False, action="upload_to_site", path=str(local_path), error="File not found.")

    upload_name = filename or local_path.name
    encoded_name = quote(upload_name, safe="")
    file_size = local_path.stat().st_size

    base = f"/drives/{drive_id}"

    if file_size <= _4MB:
        file_bytes = local_path.read_bytes()
        if folder_id:
            path = f"{base}/items/{folder_id}:/{encoded_name}:/content"
        else:
            path = f"{base}/root:/{encoded_name}:/content"

        result = await g.put(path, content=file_bytes)
    else:
        if folder_id:
            session_path = f"{base}/items/{folder_id}:/{encoded_name}:/createUploadSession"
        else:
            session_path = f"{base}/root:/{encoded_name}:/createUploadSession"

        session_payload = {
            "item": {
                "@microsoft.graph.conflictBehavior": "rename",
                "name": upload_name,
            }
        }
        session = await g.post(session_path, json=session_payload)
        upload_url = (session or {}).get("uploadUrl", "")

        if not upload_url:
            return UploadSiteFileResponse(success=False, action="upload_to_site", path=str(local_path), error="No upload URL returned.")

        result = await _upload_large_file(upload_url, local_path, file_size, ctx)

    item_id = (result or {}).get("id", "unknown")
    web_url = (result or {}).get("webUrl", "")
    size_str = _fmt_size(file_size)

    return UploadSiteFileResponse(
        success=True,
        action="upload_to_site",
        site_id=site_id,
        drive_id=drive_id,
        folder_id=folder_id,
        filename=upload_name,
        size_bytes=file_size,
        size_display=size_str,
        file_id=item_id,
        web_url=web_url,
    )


# ---------------------------------------------------------------------------
# download_from_site
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_WRITE)
async def download_from_site(
    site_id: str,
    drive_id: str,
    item_id: str,
    destination_path: Path,
    profile: str | None = None,
) -> DownloadSiteFileResponse:
    """
    Download a file from a SharePoint document library to a local path.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        drive_id: The document library (drive) ID.
        item_id: The DriveItem ID of the file to download.
        destination_path: Local path to save the file. If a directory is
                          given, the original filename from SharePoint is used.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured download confirmation.
    """
    g = _get_sharepoint_graph(profile)
    base = f"/drives/{drive_id}/items/{item_id}"

    # Get item metadata for filename
    item = await g.get(base, params={"$select": "id,name,size"})
    filename = item.get("name", "download")

    # Resolve output path
    dest = destination_path
    if dest.is_dir():
        safe_name = Path(filename).name
        if not safe_name or safe_name.startswith("."):
            safe_name = "download"
        dest = dest / safe_name

    # Download content
    content = await g.get_raw(f"{base}/content")

    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(content)

    size_str = _fmt_size(len(content))
    return DownloadSiteFileResponse(
        success=True,
        action="download_from_site",
        site_id=site_id,
        drive_id=drive_id,
        item_id=item_id,
        path=str(dest),
        filename=filename,
        size_bytes=len(content),
        size_display=size_str,
    )


# ---------------------------------------------------------------------------
# list_site_lists
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def list_site_lists(
    site_id: str,
    max_results: int = 25,
    profile: str | None = None,
) -> ListSiteListsResponse:
    """
    List all lists in a SharePoint site.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        max_results: Maximum number of lists to return. Defaults to 25.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured SharePoint list metadata.
    """
    g = _get_sharepoint_graph(profile)
    params: dict = {
        "$top": max_results,
        "$select": "id,displayName,description,webUrl,list",
    }

    result = await g.get(f"/sites/{site_id}/lists", params=params)
    lists = result.get("value", [])

    return ListSiteListsResponse(
        site_id=site_id,
        count=len(lists),
        lists=[
            SharePointListInfo(
                id=lst.get("id", ""),
                display_name=lst.get("displayName", "(unnamed)"),
                description=lst.get("description", ""),
                web_url=lst.get("webUrl", ""),
                template=lst.get("list", {}).get("template", ""),
            )
            for lst in lists
        ],
    )


# ---------------------------------------------------------------------------
# get_list_items
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def get_list_items(
    site_id: str,
    list_id: str,
    max_results: int = 25,
    profile: str | None = None,
) -> GetListItemsResponse:
    """
    Get items from a SharePoint list.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        list_id: The SharePoint list ID.
        max_results: Maximum number of items to return. Defaults to 25.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured SharePoint list-item data.
    """
    g = _get_sharepoint_graph(profile)
    params: dict = {
        "$top": max_results,
        "$expand": "fields",
        "$select": "id,createdDateTime,lastModifiedDateTime",
    }

    result = await g.get(
        f"/sites/{site_id}/lists/{list_id}/items",
        params=params,
    )
    items = result.get("value", [])

    normalized: list[SharePointListItemInfo] = []
    for item in items:
        item_id = item.get("id", "")
        created = _fmt_dt(item.get("createdDateTime"))
        modified = _fmt_dt(item.get("lastModifiedDateTime"))
        fields = item.get("fields", {})

        # Filter out internal/system fields
        user_fields = {
            k: v for k, v in fields.items()
            if not k.startswith("@") and not k.startswith("_")
        }

        title = user_fields.pop("Title", user_fields.pop("title", f"Item {item_id}"))
        filtered_fields = {
            key: value
            for key, value in user_fields.items()
            if key not in ("id", "ContentType", "Attachments", "Edit", "LinkTitleNoMenu", "LinkTitle")
        }
        normalized.append(
            SharePointListItemInfo(
                id=item_id,
                title=title,
                created_at=item.get("createdDateTime"),
                created_at_display=created,
                modified_at=item.get("lastModifiedDateTime"),
                modified_at_display=modified,
                fields=filtered_fields,
            )
        )

    return GetListItemsResponse(
        site_id=site_id,
        list_id=list_id,
        count=len(normalized),
        items=normalized,
        has_more=result.get("@odata.nextLink") is not None,
    )


# ---------------------------------------------------------------------------
# create_list_item
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_WRITE)
async def create_list_item(
    site_id: str,
    list_id: str,
    fields: SharePointFields,
    profile: str | None = None,
) -> CreateListItemResponse:
    """
    Add an item to a SharePoint list.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        list_id: The SharePoint list ID.
        fields: Field values for the new list item. Example:
                {"Title": "New item", "Status": "Active", "Priority": "High"}
                The available fields depend on the list's column definitions.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured list-item creation confirmation.
    """
    g = _get_sharepoint_graph(profile)

    payload = {"fields": fields.root}
    result = await g.post(
        f"/sites/{site_id}/lists/{list_id}/items",
        json=payload,
    )

    item_id = (result or {}).get("id", "unknown")
    return CreateListItemResponse(
        success=True,
        action="create_list_item",
        site_id=site_id,
        list_id=list_id,
        item_id=item_id,
        fields=fields.root,
    )


# ---------------------------------------------------------------------------
# update_list_item
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_WRITE)
async def update_list_item(
    site_id: str,
    list_id: str,
    item_id: str,
    fields: SharePointFields,
    profile: str | None = None,
) -> UpdateListItemResponse:
    """
    Update a SharePoint list item.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        list_id: The SharePoint list ID.
        item_id: The list item ID to update.
        fields: Field values to update. Example:
                {"Status": "Completed", "Notes": "Done on time"}
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured list-item update confirmation.
    """
    g = _get_sharepoint_graph(profile)

    await g.patch(
        f"/sites/{site_id}/lists/{list_id}/items/{item_id}/fields",
        json=fields.root,
    )

    return UpdateListItemResponse(
        success=True,
        action="update_list_item",
        site_id=site_id,
        list_id=list_id,
        item_id=item_id,
        updated_fields=list(fields.root.keys()),
        fields=fields.root,
    )


# ---------------------------------------------------------------------------
# delete_list_item
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_DESTRUCTIVE)
async def delete_list_item(
    site_id: str,
    list_id: str,
    item_id: str,
    profile: str | None = None,
) -> DeleteListItemResponse:
    """
    Delete an item from a SharePoint list.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        list_id: The SharePoint list ID.
        item_id: The list item ID to delete.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured list-item deletion confirmation.
    """
    g = _get_sharepoint_graph(profile)
    await g.delete(f"/sites/{site_id}/lists/{list_id}/items/{item_id}")
    return DeleteListItemResponse(
        success=True,
        action="delete_list_item",
        site_id=site_id,
        list_id=list_id,
        item_id=item_id,
    )
