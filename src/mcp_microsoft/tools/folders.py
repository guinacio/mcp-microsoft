"""
Mail folder tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.

Implemented:
  - list_folders
  - create_folder
  - delete_folder
"""

from __future__ import annotations

from typing import Optional

from mcp_microsoft.graph import get_graph
from mcp_microsoft.server import mcp

# ---------------------------------------------------------------------------
# list_folders
# ---------------------------------------------------------------------------


@mcp.tool()
async def list_folders(include_child_folders: bool = False, profile: str | None = None) -> str:
    """
    List mail folders in the mailbox.

    Returns all well-known folders (Inbox, SentItems, Drafts, DeletedItems,
    JunkEmail, Archive) and any custom folders created by the user.

    Well-known folder names accepted by other tools:
        inbox, sentitems, drafts, deleteditems, junkemail, archive

    Args:
        include_child_folders: When True, recursively include child folders
            for each top-level folder. Defaults to False.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted table of folders with id, displayName,
        totalItemCount, unreadItemCount, and childFolderCount.
    """
    g = get_graph(profile)
    params: dict = {
        "$top": 50,
        "$select": "id,displayName,totalItemCount,unreadItemCount,childFolderCount",
    }

    result = await g.get("/me/mailFolders", params=params)
    folders = result.get("value", [])

    if not folders:
        return "No folders found."

    lines = [f"## Mail Folders ({len(folders)} folders)\n"]
    lines.append("| Folder | Unread | Total | Children | ID |")
    lines.append("|---|---|---|---|---|")

    all_folders = list(folders)

    for folder in folders:
        child_count = folder.get("childFolderCount", 0)
        if include_child_folders and child_count > 0:
            folder_id = folder["id"]
            child_result = await g.get(
                f"/me/mailFolders/{folder_id}/childFolders",
                params={
                    "$top": 50,
                    "$select": "id,displayName,totalItemCount,unreadItemCount,childFolderCount",
                },
            )
            child_folders = child_result.get("value", [])
            # Annotate child folders with indentation marker
            for cf in child_folders:
                cf["_display_name"] = f"  \u2514 {cf.get('displayName', '')}"
            all_folders.extend(child_folders)

    seen_ids: set = set()
    rows = []
    for folder in all_folders:
        fid = folder.get("id", "")
        if fid in seen_ids:
            continue
        seen_ids.add(fid)
        name = folder.get("_display_name") or folder.get("displayName", "")
        unread = folder.get("unreadItemCount", 0)
        total = folder.get("totalItemCount", 0)
        children = folder.get("childFolderCount", 0)
        rows.append(f"| {name} | {unread} | {total} | {children} | `{fid}` |")

    lines.extend(rows)
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# create_folder
# ---------------------------------------------------------------------------


@mcp.tool()
async def create_folder(
    display_name: str,
    parent_folder_id: Optional[str] = None,
    profile: str | None = None,
) -> str:
    """
    Create a new mail folder.

    Args:
        display_name: The name for the new folder.
        parent_folder_id: Optional parent folder ID or well-known name
            (e.g. 'inbox'). When omitted, the folder is created at the
            top level of the mailbox (under the root).
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Confirmation string with the new folder's ID and display name.
    """
    g = get_graph(profile)
    payload = {"displayName": display_name}

    if parent_folder_id:
        result = await g.post(
            f"/me/mailFolders/{parent_folder_id}/childFolders",
            json=payload,
        )
    else:
        result = await g.post("/me/mailFolders", json=payload)

    folder_id = (result or {}).get("id", "unknown")
    folder_name = (result or {}).get("displayName", display_name)

    parent_info = f" under `{parent_folder_id}`" if parent_folder_id else " at top level"
    return (
        f"Folder created successfully{parent_info}.\n"
        f"**Name:** {folder_name}\n"
        f"**Folder ID:** `{folder_id}`"
    )


# ---------------------------------------------------------------------------
# delete_folder
# ---------------------------------------------------------------------------


@mcp.tool()
async def delete_folder(folder_id: str, profile: str | None = None) -> str:
    """
    Delete a mail folder and all its contents.

    WARNING: This action is permanent. All messages and sub-folders inside
    the deleted folder are also permanently removed. Unlike deleting individual
    messages, the folder and its contents are NOT moved to Deleted Items first —
    they are immediately and irreversibly destroyed. Well-known system folders
    (inbox, drafts, sentitems, etc.) cannot be deleted via this API.

    Args:
        folder_id: The folder ID to delete (opaque Graph ID, not a well-known name).
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Confirmation string on success.
    """
    g = get_graph(profile)
    await g.delete(f"/me/mailFolders/{folder_id}")
    return (
        f"Folder `{folder_id}` deleted permanently.\n"
        "All messages and sub-folders within it have been removed."
    )
