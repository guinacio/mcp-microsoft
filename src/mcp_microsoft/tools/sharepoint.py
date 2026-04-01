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

import json
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
# search_sharepoint_sites
# ---------------------------------------------------------------------------


@mcp.tool()
async def search_sharepoint_sites(
    query: str = "",
    max_results: int = 25,
    profile: str | None = None,
) -> str:
    """
    Search SharePoint sites the user has access to.

    When query is empty, uses a wildcard search to discover accessible sites.
    Requires a work/organizational Microsoft 365 account.

    Args:
        query: Search query string. Leave empty to discover accessible sites via wildcard.
        max_results: Maximum number of sites to return (1-200). Defaults to 25.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted list of SharePoint sites.
    """
    g = get_graph(profile)
    search_term = query if query else "*"
    params: dict = {
        "search": search_term,
        "$top": max_results,
        "$select": "id,displayName,description,webUrl",
    }

    result = await g.get("/sites", params=params)
    sites = result.get("value", [])

    if not sites:
        return f"No SharePoint sites found matching **{search_term}**."

    lines = [f"## SharePoint Sites ({len(sites)} results)\n"]
    for site in sites:
        lines.append(_fmt_site(site))
    lines.append("")

    next_link = result.get("@odata.nextLink")
    if next_link:
        lines.append("*More sites available — use a higher `max_results`.*")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# get_sharepoint_site
# ---------------------------------------------------------------------------


@mcp.tool()
async def get_sharepoint_site(
    site_id: str,
    profile: str | None = None,
) -> str:
    """
    Get details of a specific SharePoint site.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID (e.g. 'contoso.sharepoint.com,site-guid,web-guid').
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted site details.
    """
    g = get_graph(profile)
    params = {
        "$select": "id,displayName,description,webUrl,createdDateTime,lastModifiedDateTime",
    }

    site = await g.get(f"/sites/{site_id}", params=params)

    name = site.get("displayName", "(unnamed)")
    desc = site.get("description", "")
    web_url = site.get("webUrl", "")
    created = _fmt_dt(site.get("createdDateTime"))
    modified = _fmt_dt(site.get("lastModifiedDateTime"))

    lines = [
        f"## {name}",
    ]
    if desc:
        lines.append(f"**Description:** {desc}")
    lines.extend([
        f"**Created:** {created}",
        f"**Modified:** {modified}",
        f"**URL:** {web_url}",
        f"**Site ID:** `{site_id}`",
    ])

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# list_site_libraries
# ---------------------------------------------------------------------------


@mcp.tool()
async def list_site_libraries(
    site_id: str,
    profile: str | None = None,
) -> str:
    """
    List document libraries (drives) in a SharePoint site.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted list of document libraries.
    """
    g = get_graph(profile)
    params = {
        "$select": "id,name,description,driveType,webUrl",
    }

    result = await g.get(f"/sites/{site_id}/drives", params=params)
    drives = result.get("value", [])

    if not drives:
        return f"No document libraries found in site `{site_id}`."

    lines = [f"## Document Libraries ({len(drives)})\n"]
    for drive in drives:
        name = drive.get("name", "(unnamed)")
        drive_id = drive.get("id", "")
        desc = drive.get("description", "")
        drive_type = drive.get("driveType", "")
        web_url = drive.get("webUrl", "")

        lines.append(f"- **{name}** ({drive_type})")
        if desc:
            lines.append(f"  {desc}")
        if web_url:
            lines.append(f"  URL: {web_url}")
        lines.append(f"  Drive ID: `{drive_id}`")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# list_site_files
# ---------------------------------------------------------------------------


@mcp.tool()
async def list_site_files(
    site_id: str,
    drive_id: str,
    folder_id: Optional[str] = None,
    max_results: int = 25,
    profile: str | None = None,
) -> str:
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
        Markdown-formatted list of files and folders.
    """
    g = get_graph(profile)
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

    if not items:
        loc = f"folder `{folder_id}`" if folder_id else "library root"
        return f"No items found in {loc}."

    loc = f"folder `{folder_id}`" if folder_id else "library root"
    lines = [f"## SharePoint Library: {loc} ({len(items)} items)\n"]
    for item in items:
        lines.append(_fmt_drive_item(item))
    lines.append("")

    next_link = result.get("@odata.nextLink")
    if next_link:
        lines.append("*More items available — use a higher `max_results` or paginate.*")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# get_site_file
# ---------------------------------------------------------------------------


@mcp.tool()
async def get_site_file(
    site_id: str,
    drive_id: str,
    item_id: str,
    profile: str | None = None,
) -> str:
    """
    Get metadata for a file or folder in a SharePoint document library.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        drive_id: The document library (drive) ID.
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
# upload_to_site
# ---------------------------------------------------------------------------


@mcp.tool()
async def upload_to_site(
    site_id: str,
    drive_id: str,
    local_path: str,
    folder_id: Optional[str] = None,
    filename: Optional[str] = None,
    profile: str | None = None,
) -> str:
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
        Confirmation string with the new file's ID and web URL.
    """
    g = get_graph(profile)
    local = Path(local_path)
    if not local.is_file():
        return f"File not found: `{local_path}`"

    upload_name = filename or local.name
    encoded_name = quote(upload_name, safe="")
    file_size = local.stat().st_size
    file_bytes = local.read_bytes()

    base = f"/drives/{drive_id}"

    if file_size <= _4MB:
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


# ---------------------------------------------------------------------------
# download_from_site
# ---------------------------------------------------------------------------


@mcp.tool()
async def download_from_site(
    site_id: str,
    drive_id: str,
    item_id: str,
    destination_path: str,
    profile: str | None = None,
) -> str:
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
        Confirmation string with the saved file path and size.
    """
    g = get_graph(profile)
    base = f"/drives/{drive_id}/items/{item_id}"

    # Get item metadata for filename
    item = await g.get(base, params={"$select": "id,name,size"})
    filename = item.get("name", "download")

    # Resolve output path
    dest = Path(destination_path)
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
    return (
        f"File downloaded successfully.\n"
        f"**Saved to:** `{dest}`\n"
        f"**Filename:** {filename}\n"
        f"**Size:** {size_str}"
    )


# ---------------------------------------------------------------------------
# list_site_lists
# ---------------------------------------------------------------------------


@mcp.tool()
async def list_site_lists(
    site_id: str,
    max_results: int = 25,
    profile: str | None = None,
) -> str:
    """
    List all lists in a SharePoint site.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        max_results: Maximum number of lists to return. Defaults to 25.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted list of SharePoint lists.
    """
    g = get_graph(profile)
    params: dict = {
        "$top": max_results,
        "$select": "id,displayName,description,webUrl,list",
    }

    result = await g.get(f"/sites/{site_id}/lists", params=params)
    lists = result.get("value", [])

    if not lists:
        return f"No lists found in site `{site_id}`."

    lines = [f"## SharePoint Lists ({len(lists)})\n"]
    for lst in lists:
        name = lst.get("displayName", "(unnamed)")
        list_id = lst.get("id", "")
        desc = lst.get("description", "")
        web_url = lst.get("webUrl", "")
        template = lst.get("list", {}).get("template", "")

        lines.append(f"- **{name}**" + (f" ({template})" if template else ""))
        if desc:
            lines.append(f"  {desc}")
        if web_url:
            lines.append(f"  URL: {web_url}")
        lines.append(f"  List ID: `{list_id}`")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# get_list_items
# ---------------------------------------------------------------------------


@mcp.tool()
async def get_list_items(
    site_id: str,
    list_id: str,
    max_results: int = 25,
    profile: str | None = None,
) -> str:
    """
    Get items from a SharePoint list.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        list_id: The SharePoint list ID.
        max_results: Maximum number of items to return. Defaults to 25.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted list items with their field values.
    """
    g = get_graph(profile)
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

    if not items:
        return f"No items found in list `{list_id}`."

    lines = [f"## List Items ({len(items)})\n"]
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
        lines.append(f"### {title}")
        lines.append(f"**Item ID:** `{item_id}` | Created: {created} | Modified: {modified}")

        for key, value in user_fields.items():
            # Skip common metadata fields that add noise
            if key in ("id", "ContentType", "Attachments", "Edit", "LinkTitleNoMenu", "LinkTitle"):
                continue
            lines.append(f"- **{key}:** {value}")
        lines.append("")

    next_link = result.get("@odata.nextLink")
    if next_link:
        lines.append("*More items available — use a higher `max_results`.*")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# create_list_item
# ---------------------------------------------------------------------------


@mcp.tool()
async def create_list_item(
    site_id: str,
    list_id: str,
    fields: str,
    profile: str | None = None,
) -> str:
    """
    Add an item to a SharePoint list.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        list_id: The SharePoint list ID.
        fields: A JSON string of field values. Example:
                '{"Title": "New item", "Status": "Active", "Priority": "High"}'
                The available fields depend on the list's column definitions.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Confirmation string with the new item ID.
    """
    g = get_graph(profile)

    try:
        field_data = json.loads(fields)
    except json.JSONDecodeError as e:
        return f"Invalid JSON in `fields`: {e}"

    if not isinstance(field_data, dict):
        return "`fields` must be a JSON object, e.g. '{\"Title\": \"value\"}'."

    payload = {"fields": field_data}
    result = await g.post(
        f"/sites/{site_id}/lists/{list_id}/items",
        json=payload,
    )

    item_id = (result or {}).get("id", "unknown")
    return (
        f"List item created successfully.\n"
        f"**Item ID:** `{item_id}`"
    )


# ---------------------------------------------------------------------------
# update_list_item
# ---------------------------------------------------------------------------


@mcp.tool()
async def update_list_item(
    site_id: str,
    list_id: str,
    item_id: str,
    fields: str,
    profile: str | None = None,
) -> str:
    """
    Update a SharePoint list item.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        list_id: The SharePoint list ID.
        item_id: The list item ID to update.
        fields: A JSON string of field values to update. Example:
                '{"Status": "Completed", "Notes": "Done on time"}'
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Confirmation string with updated field names.
    """
    g = get_graph(profile)

    try:
        field_data = json.loads(fields)
    except json.JSONDecodeError as e:
        return f"Invalid JSON in `fields`: {e}"

    if not isinstance(field_data, dict):
        return "`fields` must be a JSON object, e.g. '{\"Status\": \"Completed\"}'."

    await g.patch(
        f"/sites/{site_id}/lists/{list_id}/items/{item_id}/fields",
        json=field_data,
    )

    updated_keys = ", ".join(field_data.keys())
    return (
        f"List item updated successfully.\n"
        f"**Item ID:** `{item_id}`\n"
        f"**Updated fields:** {updated_keys}"
    )


# ---------------------------------------------------------------------------
# delete_list_item
# ---------------------------------------------------------------------------


@mcp.tool()
async def delete_list_item(
    site_id: str,
    list_id: str,
    item_id: str,
    profile: str | None = None,
) -> str:
    """
    Delete an item from a SharePoint list.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        list_id: The SharePoint list ID.
        item_id: The list item ID to delete.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Confirmation string.
    """
    g = get_graph(profile)
    await g.delete(f"/sites/{site_id}/lists/{list_id}/items/{item_id}")
    return f"List item `{item_id}` deleted successfully."
