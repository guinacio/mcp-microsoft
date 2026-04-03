"""
Mail folder tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.

Implemented:
  - list_folders
  - create_folder
  - delete_folder
"""

from __future__ import annotations

from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.tooling import DESTRUCTIVE_TOOL, READ_ONLY_TOOL, WRITE_TOOL, register_tool
from mcp_microsoft.graph import get_graph
from mcp_microsoft.models import (
    CreateFolderResponse,
    DeleteFolderResponse,
    ListFoldersResponse,
    MailFolderInfo,
)

# ---------------------------------------------------------------------------
# list_folders
# ---------------------------------------------------------------------------

class ListFoldersInput(ToolRequestModel):
    include_child_folders: bool = False
    profile: str | None = None


class CreateFolderInput(ToolRequestModel):
    display_name: str
    parent_folder_id: str | None = None
    profile: str | None = None


class DeleteFolderInput(ToolRequestModel):
    folder_id: str
    profile: str | None = None


async def list_folders(params: ListFoldersInput) -> ListFoldersResponse:
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
        Structured mail folder data.
    """
    g = get_graph(params.profile)
    query: dict = {
        "$top": 50,
        "$select": "id,displayName,totalItemCount,unreadItemCount,childFolderCount",
    }

    result = await g.get("/me/mailFolders", params=query)
    folders = result.get("value", [])

    all_folders = list(folders)

    for folder in folders:
        child_count = folder.get("childFolderCount", 0)
        if params.include_child_folders and child_count > 0:
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
    items: list[MailFolderInfo] = []
    for folder in all_folders:
        fid = folder.get("id", "")
        if fid in seen_ids:
            continue
        seen_ids.add(fid)
        name = folder.get("_display_name") or folder.get("displayName", "")
        unread = folder.get("unreadItemCount", 0)
        total = folder.get("totalItemCount", 0)
        children = folder.get("childFolderCount", 0)
        items.append(
            MailFolderInfo(
                id=fid,
                display_name=folder.get("displayName", ""),
                display_label=name,
                unread_count=unread,
                total_count=total,
                child_folder_count=children,
                is_child="_display_name" in folder,
            )
        )

    return ListFoldersResponse(
        count=len(items),
        include_child_folders=params.include_child_folders,
        folders=items,
    )


# ---------------------------------------------------------------------------
# create_folder
# ---------------------------------------------------------------------------


async def create_folder(
    params: CreateFolderInput,
) -> CreateFolderResponse:
    """
    Create a new mail folder in the mailbox (not OneDrive — use create_drive_folder for files).

    Args:
        display_name: The name for the new folder.
        parent_folder_id: Optional parent folder ID or well-known name
            (e.g. 'inbox'). When omitted, the folder is created at the
            top level of the mailbox (under the root).
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured folder creation confirmation.
    """
    g = get_graph(params.profile)
    payload = {"displayName": params.display_name}

    if params.parent_folder_id:
        result = await g.post(
            f"/me/mailFolders/{params.parent_folder_id}/childFolders",
            json=payload,
        )
    else:
        result = await g.post("/me/mailFolders", json=payload)

    folder_id = (result or {}).get("id", "unknown")
    folder_name = (result or {}).get("displayName", params.display_name)

    return CreateFolderResponse(
        success=True,
        action="create_folder",
        folder_id=folder_id,
        display_name=folder_name,
        parent_folder_id=params.parent_folder_id,
    )


# ---------------------------------------------------------------------------
# delete_folder
# ---------------------------------------------------------------------------


async def delete_folder(params: DeleteFolderInput) -> DeleteFolderResponse:
    """
    Delete a mail folder and all its contents.

    WARNING: This action is permanent. All messages and sub-folders inside
    the deleted folder are also permanently removed. Unlike deleting individual
    messages, the folder and its contents are NOT moved to Deleted Items first —
    they are immediately and irreversibly destroyed. Well-known system folders
    (inbox, drafts, sentitems, etc.) cannot be deleted via this API.

    Args:
        folder_id: The opaque Graph ID of the folder to delete. This is NOT a
            display name or well-known name — use list_folders first to obtain
            the folder ID by its display name.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured delete confirmation.
    """
    g = get_graph(params.profile)
    await g.delete(f"/me/mailFolders/{params.folder_id}")
    return DeleteFolderResponse(
        success=True,
        action="delete_folder",
        folder_id=params.folder_id,
        irreversible=True,
    )


def register(server) -> None:
    """Register all folder tools with the given FastMCP server instance."""
    register_tool(server, list_folders, annotations=READ_ONLY_TOOL)
    register_tool(server, create_folder, annotations=WRITE_TOOL)
    register_tool(server, delete_folder, annotations=DESTRUCTIVE_TOOL)
