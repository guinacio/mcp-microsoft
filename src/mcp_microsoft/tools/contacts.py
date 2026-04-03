"""
Contacts tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.
Endpoints live under /me/contacts and /me/contactFolders in Graph v1.0.

Implemented:
  - list_contacts
  - get_contact
  - create_contact
  - update_contact
  - delete_contact
  - list_contact_folders
  - search_contacts
  - get_contact_photo
"""

from __future__ import annotations

import base64
from pathlib import Path
from typing import Any, Optional
from urllib.parse import parse_qs, urlparse

from mcp_microsoft.common.tooling import DESTRUCTIVE_TOOL, READ_ONLY_TOOL, WRITE_TOOL, register_tool
from mcp_microsoft.models import (
    ContactEmailAddress,
    ContactFolderInfo,
    ContactInfo,
    CreateContactResponse,
    DeleteContactResponse,
    GetContactPhotoResponse,
    GetContactResponse,
    ListContactFoldersResponse,
    ListContactsResponse,
    SearchContactsResponse,
    UpdateContactResponse,
)
from mcp_microsoft.graph import get_graph
from mcp_microsoft.common.request_model import ToolRequestModel

# Fields to $select for contact list/search operations (lighter payload)
_CONTACT_SELECT = (
    "id,displayName,givenName,surname,emailAddresses,"
    "mobilePhone,businessPhones,jobTitle,companyName,department"
)

# Fields to $select for full contact detail
_CONTACT_DETAIL_SELECT = (
    "id,displayName,givenName,surname,emailAddresses,"
    "mobilePhone,businessPhones,jobTitle,companyName,department,personalNotes"
)


class ListContactsInput(ToolRequestModel):
    folder_id: str | None = None
    search: str | None = None
    top: int = 50
    skip_token: str | None = None
    profile: str | None = None


class GetContactInput(ToolRequestModel):
    contact_id: str
    profile: str | None = None


class CreateContactInput(ToolRequestModel):
    display_name: str
    given_name: Optional[str] = None
    surname: Optional[str] = None
    email_addresses: Optional[list[dict]] = None
    mobile_phone: Optional[str] = None
    business_phones: Optional[list[str]] = None
    job_title: Optional[str] = None
    company_name: Optional[str] = None
    department: Optional[str] = None
    notes: Optional[str] = None
    folder_id: Optional[str] = None
    profile: str | None = None


class UpdateContactInput(ToolRequestModel):
    contact_id: str
    display_name: Optional[str] = None
    given_name: Optional[str] = None
    surname: Optional[str] = None
    email_addresses: Optional[list[dict]] = None
    mobile_phone: Optional[str] = None
    business_phones: Optional[list[str]] = None
    job_title: Optional[str] = None
    company_name: Optional[str] = None
    department: Optional[str] = None
    notes: Optional[str] = None
    profile: str | None = None


class DeleteContactInput(ToolRequestModel):
    contact_id: str
    profile: str | None = None


class ListContactFoldersInput(ToolRequestModel):
    profile: str | None = None


class SearchContactsInput(ToolRequestModel):
    query: str
    top: int = 25
    profile: str | None = None


class GetContactPhotoInput(ToolRequestModel):
    contact_id: str
    save_path: Optional[str] = None
    profile: str | None = None


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------


def _normalize_contact(c: dict[str, Any]) -> ContactInfo:
    """Normalize a Graph contact object into a ContactInfo summary."""
    return ContactInfo(
        id=c.get("id", ""),
        display_name=c.get("displayName") or "",
        given_name=c.get("givenName") or "",
        surname=c.get("surname") or "",
        email_addresses=[
            ContactEmailAddress(
                address=ea.get("address", ""),
                name=ea.get("name", ""),
            )
            for ea in (c.get("emailAddresses") or [])
        ],
        mobile_phone=c.get("mobilePhone", "") or "",
        business_phones=c.get("businessPhones") or [],
        job_title=c.get("jobTitle", "") or "",
        company_name=c.get("companyName", "") or "",
        department=c.get("department", "") or "",
    )


# ---------------------------------------------------------------------------
# list_contacts
# ---------------------------------------------------------------------------


async def list_contacts(params: ListContactsInput) -> ListContactsResponse:
    """
    List contacts from the user's default contacts folder or a specific folder.

    Args:
        params: Structured contact-list request.

    Returns:
        Structured list of contacts with id, name, email, phone, and job info.
        When has_more is True, pass next_page_token as skip_token to retrieve
        the next page.
    """
    g = get_graph(params.profile)
    top = max(1, min(params.top, 100))

    query: dict[str, Any] = {
        "$select": _CONTACT_SELECT,
        "$top": top,
        "$orderby": "displayName",
    }
    if params.skip_token is not None:
        query["$skiptoken"] = params.skip_token

    extra_headers: dict[str, str] | None = None
    if params.search:
        # $search requires ConsistencyLevel: eventual
        query["$search"] = params.search
        extra_headers = {"ConsistencyLevel": "eventual"}

    if params.folder_id:
        path = f"/me/contactFolders/{params.folder_id}/contacts"
    else:
        path = "/me/contacts"

    result = await g.get(path, params=query, headers=extra_headers)
    contacts = (result or {}).get("value", [])

    next_link = (result or {}).get("@odata.nextLink", "")
    next_page_token: str | None = None
    if next_link:
        qs = parse_qs(urlparse(next_link).query)
        next_page_token = qs.get("$skiptoken", [None])[0]

    return ListContactsResponse(
        count=len(contacts),
        folder_id=params.folder_id,
        contacts=[_normalize_contact(c) for c in contacts],
        next_page_token=next_page_token,
        has_more=(next_page_token is not None),
    )


# ---------------------------------------------------------------------------
# get_contact
# ---------------------------------------------------------------------------


async def get_contact(params: GetContactInput) -> GetContactResponse:
    """
    Fetch a single contact by ID with full details including notes.

    Args:
        params: Structured contact-detail request.

    Returns:
        Full contact details including personal notes.
    """
    g = get_graph(params.profile)
    query = {"$select": _CONTACT_DETAIL_SELECT}

    c = await g.get(f"/me/contacts/{params.contact_id}", params=query)

    return GetContactResponse(
        id=params.contact_id,
        display_name=c.get("displayName") or "",
        given_name=c.get("givenName") or "",
        surname=c.get("surname") or "",
        email_addresses=[
            ContactEmailAddress(
                address=ea.get("address", ""),
                name=ea.get("name", ""),
            )
            for ea in (c.get("emailAddresses") or [])
        ],
        mobile_phone=c.get("mobilePhone", "") or "",
        business_phones=c.get("businessPhones") or [],
        job_title=c.get("jobTitle", "") or "",
        company_name=c.get("companyName", "") or "",
        department=c.get("department", "") or "",
        notes=c.get("personalNotes", "") or "",
    )


# ---------------------------------------------------------------------------
# create_contact
# ---------------------------------------------------------------------------


async def create_contact(params: CreateContactInput) -> CreateContactResponse:
    """
    Create a new contact in the user's contacts folder.

    Args:
        params: Structured contact-create request.

    Returns:
        Created contact ID and display name.
    """
    g = get_graph(params.profile)

    body: dict[str, Any] = {"displayName": params.display_name}

    if params.given_name is not None:
        body["givenName"] = params.given_name
    if params.surname is not None:
        body["surname"] = params.surname
    if params.email_addresses is not None:
        body["emailAddresses"] = [
            {"address": ea.get("address", ""), "name": ea.get("name", "")}
            for ea in params.email_addresses
        ]
    if params.mobile_phone is not None:
        body["mobilePhone"] = params.mobile_phone
    if params.business_phones is not None:
        body["businessPhones"] = params.business_phones
    if params.job_title is not None:
        body["jobTitle"] = params.job_title
    if params.company_name is not None:
        body["companyName"] = params.company_name
    if params.department is not None:
        body["department"] = params.department
    if params.notes is not None:
        body["personalNotes"] = params.notes

    if params.folder_id:
        path = f"/me/contactFolders/{params.folder_id}/contacts"
    else:
        path = "/me/contacts"

    result = await g.post(path, json=body)

    contact_id = (result or {}).get("id", "unknown")
    created_name = (result or {}).get("displayName", params.display_name)

    return CreateContactResponse(
        success=True,
        action="create_contact",
        contact_id=contact_id,
        display_name=created_name,
        folder_id=params.folder_id,
    )


# ---------------------------------------------------------------------------
# update_contact
# ---------------------------------------------------------------------------


async def update_contact(params: UpdateContactInput) -> UpdateContactResponse:
    """
    Update an existing contact. Only provided fields are changed (PATCH semantics).

    Args:
        params: Structured contact-update request.

    Returns:
        Updated contact ID and list of changed fields.
    """
    g = get_graph(params.profile)
    patch: dict[str, Any] = {}

    if params.display_name is not None:
        patch["displayName"] = params.display_name
    if params.given_name is not None:
        patch["givenName"] = params.given_name
    if params.surname is not None:
        patch["surname"] = params.surname
    if params.email_addresses is not None:
        patch["emailAddresses"] = [
            {"address": ea.get("address", ""), "name": ea.get("name", "")}
            for ea in params.email_addresses
        ]
    if params.mobile_phone is not None:
        patch["mobilePhone"] = params.mobile_phone
    if params.business_phones is not None:
        patch["businessPhones"] = params.business_phones
    if params.job_title is not None:
        patch["jobTitle"] = params.job_title
    if params.company_name is not None:
        patch["companyName"] = params.company_name
    if params.department is not None:
        patch["department"] = params.department
    if params.notes is not None:
        patch["personalNotes"] = params.notes

    if not patch:
        return UpdateContactResponse(
            success=False,
            action="update_contact",
            contact_id=params.contact_id,
            updated_fields=[],
            error="No fields to update.",
        )

    result = await g.patch(f"/me/contacts/{params.contact_id}", json=patch)
    updated_id = (result or {}).get("id", params.contact_id)
    updated_fields = list(patch.keys())

    return UpdateContactResponse(
        success=True,
        action="update_contact",
        contact_id=updated_id,
        updated_fields=updated_fields,
        updated_fields_display=", ".join(updated_fields),
    )


# ---------------------------------------------------------------------------
# delete_contact
# ---------------------------------------------------------------------------


async def delete_contact(params: DeleteContactInput) -> DeleteContactResponse:
    """
    Permanently delete a contact by ID.

    Args:
        params: Structured contact-delete request.

    Returns:
        Deletion confirmation.
    """
    g = get_graph(params.profile)
    await g.delete(f"/me/contacts/{params.contact_id}")
    return DeleteContactResponse(
        success=True,
        action="delete_contact",
        contact_id=params.contact_id,
    )


# ---------------------------------------------------------------------------
# list_contact_folders
# ---------------------------------------------------------------------------


async def list_contact_folders(
    params: ListContactFoldersInput,
) -> ListContactFoldersResponse:
    """
    List all contact folders in the user's mailbox.

    Returns folder IDs that can be passed to list_contacts or create_contact
    to scope operations to a specific folder.

    Args:
        params: Structured folder-list request.

    Returns:
        Structured list of contact folders with id, displayName, and item count.
    """
    g = get_graph(params.profile)
    query = {
        "$select": "id,displayName,parentFolderId,totalItemCount",
        "$top": 100,
    }

    result = await g.get("/me/contactFolders", params=query)
    folders = (result or {}).get("value", [])

    return ListContactFoldersResponse(
        count=len(folders),
        folders=[
            ContactFolderInfo(
                id=f.get("id", ""),
                display_name=f.get("displayName", ""),
                parent_folder_id=f.get("parentFolderId", "") or "",
                total_item_count=f.get("totalItemCount", 0) or 0,
            )
            for f in folders
        ],
    )


# ---------------------------------------------------------------------------
# search_contacts
# ---------------------------------------------------------------------------


async def search_contacts(params: SearchContactsInput) -> SearchContactsResponse:
    """
    Search contacts by display name prefix using $filter startswith.

    Uses OData $filter with startswith() for broad tenant compatibility.
    For full-text search, use list_contacts with the search parameter instead.

    Args:
        params: Structured contact-search request.

    Returns:
        Contacts whose displayName starts with the given query.
    """
    g = get_graph(params.profile)
    top = max(1, min(params.top, 100))

    # Sanitize: remove single quotes to prevent OData injection
    safe_query = params.query.replace("'", "")

    query_params: dict[str, Any] = {
        "$select": _CONTACT_SELECT,
        "$filter": f"startswith(displayName,'{safe_query}')",
        "$top": top,
        "$orderby": "displayName",
    }

    result = await g.get("/me/contacts", params=query_params)
    contacts = (result or {}).get("value", [])

    return SearchContactsResponse(
        query=params.query,
        count=len(contacts),
        contacts=[_normalize_contact(c) for c in contacts],
    )


# ---------------------------------------------------------------------------
# get_contact_photo
# ---------------------------------------------------------------------------


async def get_contact_photo(params: GetContactPhotoInput) -> GetContactPhotoResponse:
    """
    Retrieve the profile photo for a contact.

    Args:
        params: Structured contact-photo request.

    Returns:
        Base64-encoded photo bytes and/or saved file path confirmation.
    """
    g = get_graph(params.profile)
    raw: bytes = await g.get_raw(f"/me/contacts/{params.contact_id}/photo/$value")

    photo_b64 = base64.b64encode(raw).decode("ascii")
    saved_path = ""

    if params.save_path:
        dest = Path(params.save_path)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(raw)
        saved_path = str(dest)

    return GetContactPhotoResponse(
        success=True,
        action="get_contact_photo",
        contact_id=params.contact_id,
        photo_base64=photo_b64,
        saved_path=saved_path,
        size_bytes=len(raw),
    )


def register(server) -> None:
    """Register all contact tools with the given FastMCP server instance."""
    register_tool(server, list_contacts, annotations=READ_ONLY_TOOL)
    register_tool(server, get_contact, annotations=READ_ONLY_TOOL)
    register_tool(server, create_contact, annotations=WRITE_TOOL)
    register_tool(server, update_contact, annotations=WRITE_TOOL)
    register_tool(server, delete_contact, annotations=DESTRUCTIVE_TOOL)
    register_tool(server, list_contact_folders, annotations=READ_ONLY_TOOL)
    register_tool(server, search_contacts, annotations=READ_ONLY_TOOL)
    register_tool(server, get_contact_photo, annotations=READ_ONLY_TOOL)
