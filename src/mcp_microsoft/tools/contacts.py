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

from mcp.types import ToolAnnotations

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
from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.graph import get_graph

# ---------------------------------------------------------------------------
# Annotation constants
# ---------------------------------------------------------------------------

_READ_ONLY = ToolAnnotations(readOnlyHint=True, openWorldHint=True)
_WRITE = ToolAnnotations(destructiveHint=False, openWorldHint=True)
_DESTRUCTIVE = ToolAnnotations(destructiveHint=True, openWorldHint=True)

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


async def list_contacts(
    folder_id: Optional[str] = None,
    search: Optional[str] = None,
    top: int = 50,
    skip_token: Optional[str] = None,
    profile: str | None = None,
) -> ListContactsResponse:
    """
    List contacts from the user's default contacts folder or a specific folder.

    Args:
        folder_id: Optional contact folder ID. If provided, lists contacts from
                   that folder instead of the default contacts folder.
        search: Optional OData $search query string (e.g. 'displayName:John').
                Note: $search requires the ConsistencyLevel header and may not be
                supported in all tenants. Use search_contacts for broader compatibility.
        top: Maximum number of contacts to return (1-100). Defaults to 50.
        skip_token: Opaque pagination cursor returned as next_page_token from a
                    previous call. Omit for the first page.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured list of contacts with id, name, email, phone, and job info.
        When has_more is True, pass next_page_token as skip_token to retrieve
        the next page.
    """
    g = get_graph(profile)
    top = max(1, min(top, 100))

    params: dict[str, Any] = {
        "$select": _CONTACT_SELECT,
        "$top": top,
        "$orderby": "displayName",
    }
    if skip_token is not None:
        params["$skiptoken"] = skip_token

    extra_headers: dict[str, str] | None = None
    if search:
        # $search requires ConsistencyLevel: eventual
        params["$search"] = search
        extra_headers = {"ConsistencyLevel": "eventual"}

    if folder_id:
        path = f"/me/contactFolders/{folder_id}/contacts"
    else:
        path = "/me/contacts"

    result = await g.get(path, params=params, headers=extra_headers)
    contacts = (result or {}).get("value", [])

    next_link = (result or {}).get("@odata.nextLink", "")
    next_page_token: str | None = None
    if next_link:
        qs = parse_qs(urlparse(next_link).query)
        next_page_token = qs.get("$skiptoken", [None])[0]

    return ListContactsResponse(
        count=len(contacts),
        folder_id=folder_id,
        contacts=[_normalize_contact(c) for c in contacts],
        next_page_token=next_page_token,
        has_more=(next_page_token is not None),
    )


# ---------------------------------------------------------------------------
# get_contact
# ---------------------------------------------------------------------------


async def get_contact(
    contact_id: str,
    profile: str | None = None,
) -> GetContactResponse:
    """
    Fetch a single contact by ID with full details including notes.

    Args:
        contact_id: The Graph contact ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Full contact details including personal notes.
    """
    g = get_graph(profile)
    params = {"$select": _CONTACT_DETAIL_SELECT}

    c = await g.get(f"/me/contacts/{contact_id}", params=params)

    return GetContactResponse(
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
        notes=c.get("personalNotes", "") or "",
    )


# ---------------------------------------------------------------------------
# create_contact
# ---------------------------------------------------------------------------


class CreateContactInput(ToolRequestModel):
    """Validated input for the create_contact tool."""

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


async def create_contact(
    display_name: str,
    given_name: Optional[str] = None,
    surname: Optional[str] = None,
    email_addresses: Optional[list[dict]] = None,
    mobile_phone: Optional[str] = None,
    business_phones: Optional[list[str]] = None,
    job_title: Optional[str] = None,
    company_name: Optional[str] = None,
    department: Optional[str] = None,
    notes: Optional[str] = None,
    folder_id: Optional[str] = None,
    profile: str | None = None,
) -> CreateContactResponse:
    """
    Create a new contact in the user's contacts folder.

    Args:
        display_name: Full display name for the contact (required).
        given_name: First name.
        surname: Last name / family name.
        email_addresses: List of email address dicts, each with 'address' and
                         optional 'name' keys. Example:
                         [{"address": "jane@example.com", "name": "Jane"}]
        mobile_phone: Mobile phone number.
        business_phones: List of business phone number strings.
        job_title: Job title / position.
        company_name: Employer / company name.
        department: Department within the company.
        notes: Free-text personal notes.
        folder_id: Optional contact folder ID. Defaults to the root contacts folder.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Created contact ID and display name.
    """
    p = CreateContactInput.model_validate({
        "display_name": display_name, "given_name": given_name, "surname": surname,
        "email_addresses": email_addresses, "mobile_phone": mobile_phone,
        "business_phones": business_phones, "job_title": job_title,
        "company_name": company_name, "department": department,
        "notes": notes, "folder_id": folder_id, "profile": profile,
    })
    g = get_graph(p.profile)

    body: dict[str, Any] = {"displayName": p.display_name}

    if p.given_name is not None:
        body["givenName"] = p.given_name
    if p.surname is not None:
        body["surname"] = p.surname
    if p.email_addresses is not None:
        body["emailAddresses"] = [
            {"address": ea.get("address", ""), "name": ea.get("name", "")}
            for ea in p.email_addresses
        ]
    if p.mobile_phone is not None:
        body["mobilePhone"] = p.mobile_phone
    if p.business_phones is not None:
        body["businessPhones"] = p.business_phones
    if p.job_title is not None:
        body["jobTitle"] = p.job_title
    if p.company_name is not None:
        body["companyName"] = p.company_name
    if p.department is not None:
        body["department"] = p.department
    if p.notes is not None:
        body["personalNotes"] = p.notes

    if p.folder_id:
        path = f"/me/contactFolders/{p.folder_id}/contacts"
    else:
        path = "/me/contacts"

    result = await g.post(path, json=body)

    contact_id = (result or {}).get("id", "unknown")
    created_name = (result or {}).get("displayName", p.display_name)

    return CreateContactResponse(
        success=True,
        action="create_contact",
        contact_id=contact_id,
        display_name=created_name,
        folder_id=p.folder_id,
    )


# ---------------------------------------------------------------------------
# update_contact
# ---------------------------------------------------------------------------


async def update_contact(
    contact_id: str,
    display_name: Optional[str] = None,
    given_name: Optional[str] = None,
    surname: Optional[str] = None,
    email_addresses: Optional[list[dict]] = None,
    mobile_phone: Optional[str] = None,
    business_phones: Optional[list[str]] = None,
    job_title: Optional[str] = None,
    company_name: Optional[str] = None,
    department: Optional[str] = None,
    notes: Optional[str] = None,
    profile: str | None = None,
) -> UpdateContactResponse:
    """
    Update an existing contact. Only provided fields are changed (PATCH semantics).

    Args:
        contact_id: The Graph contact ID to update (required).
        display_name: Replace display name (omit to leave unchanged).
        given_name: Replace first name (omit to leave unchanged).
        surname: Replace last name (omit to leave unchanged).
        email_addresses: Replace email list entirely. Each dict must have 'address'
                         and optionally 'name'. Omit to leave unchanged.
        mobile_phone: Replace mobile phone (omit to leave unchanged).
        business_phones: Replace business phone list (omit to leave unchanged).
        job_title: Replace job title (omit to leave unchanged).
        company_name: Replace company name (omit to leave unchanged).
        department: Replace department (omit to leave unchanged).
        notes: Replace personal notes (omit to leave unchanged).
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Updated contact ID and list of changed fields.
    """
    g = get_graph(profile)
    patch: dict[str, Any] = {}

    if display_name is not None:
        patch["displayName"] = display_name
    if given_name is not None:
        patch["givenName"] = given_name
    if surname is not None:
        patch["surname"] = surname
    if email_addresses is not None:
        patch["emailAddresses"] = [
            {"address": ea.get("address", ""), "name": ea.get("name", "")}
            for ea in email_addresses
        ]
    if mobile_phone is not None:
        patch["mobilePhone"] = mobile_phone
    if business_phones is not None:
        patch["businessPhones"] = business_phones
    if job_title is not None:
        patch["jobTitle"] = job_title
    if company_name is not None:
        patch["companyName"] = company_name
    if department is not None:
        patch["department"] = department
    if notes is not None:
        patch["personalNotes"] = notes

    if not patch:
        return UpdateContactResponse(
            success=False,
            action="update_contact",
            contact_id=contact_id,
            updated_fields=[],
            error="No fields to update.",
        )

    result = await g.patch(f"/me/contacts/{contact_id}", json=patch)
    updated_id = (result or {}).get("id", contact_id)
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


async def delete_contact(
    contact_id: str,
    profile: str | None = None,
) -> DeleteContactResponse:
    """
    Permanently delete a contact by ID.

    Args:
        contact_id: The Graph contact ID to delete.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Deletion confirmation.
    """
    g = get_graph(profile)
    await g.delete(f"/me/contacts/{contact_id}")
    return DeleteContactResponse(
        success=True,
        action="delete_contact",
        contact_id=contact_id,
    )


# ---------------------------------------------------------------------------
# list_contact_folders
# ---------------------------------------------------------------------------


async def list_contact_folders(
    profile: str | None = None,
) -> ListContactFoldersResponse:
    """
    List all contact folders in the user's mailbox.

    Returns folder IDs that can be passed to list_contacts or create_contact
    to scope operations to a specific folder.

    Args:
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured list of contact folders with id, displayName, and item count.
    """
    g = get_graph(profile)
    params = {
        "$select": "id,displayName,parentFolderId,totalItemCount",
        "$top": 100,
    }

    result = await g.get("/me/contactFolders", params=params)
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


async def search_contacts(
    query: str,
    top: int = 25,
    profile: str | None = None,
) -> SearchContactsResponse:
    """
    Search contacts by display name prefix using $filter startswith.

    Uses OData $filter with startswith() for broad tenant compatibility.
    For full-text search, use list_contacts with the search parameter instead.

    Args:
        query: Display name prefix to search for (case-insensitive).
        top: Maximum number of results to return (1-100). Defaults to 25.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Contacts whose displayName starts with the given query.
    """
    g = get_graph(profile)
    top = max(1, min(top, 100))

    # Sanitize: remove single quotes to prevent OData injection
    safe_query = query.replace("'", "")

    params: dict[str, Any] = {
        "$select": _CONTACT_SELECT,
        "$filter": f"startswith(displayName,'{safe_query}')",
        "$top": top,
        "$orderby": "displayName",
    }

    result = await g.get("/me/contacts", params=params)
    contacts = (result or {}).get("value", [])

    return SearchContactsResponse(
        query=query,
        count=len(contacts),
        contacts=[_normalize_contact(c) for c in contacts],
    )


# ---------------------------------------------------------------------------
# get_contact_photo
# ---------------------------------------------------------------------------


async def get_contact_photo(
    contact_id: str,
    save_path: Optional[str] = None,
    profile: str | None = None,
) -> GetContactPhotoResponse:
    """
    Retrieve the profile photo for a contact.

    Args:
        contact_id: The Graph contact ID.
        save_path: Optional absolute file path to save the photo to disk
                   (e.g. 'C:/Users/you/Desktop/photo.jpg'). If omitted, the
                   photo is returned as base64-encoded bytes only.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Base64-encoded photo bytes and/or saved file path confirmation.
    """
    g = get_graph(profile)
    raw: bytes = await g.get_raw(f"/me/contacts/{contact_id}/photo/$value")

    photo_b64 = base64.b64encode(raw).decode("ascii")
    saved_path = ""

    if save_path:
        dest = Path(save_path)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(raw)
        saved_path = str(dest)

    return GetContactPhotoResponse(
        success=True,
        action="get_contact_photo",
        contact_id=contact_id,
        photo_base64=photo_b64,
        saved_path=saved_path,
        size_bytes=len(raw),
    )


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------


def register(server) -> None:
    """Register all contact tools with the given FastMCP server instance."""
    server.tool(annotations=_READ_ONLY)(list_contacts)
    server.tool(annotations=_READ_ONLY)(get_contact)
    server.tool(annotations=_WRITE)(create_contact)
    server.tool(annotations=_WRITE)(update_contact)
    server.tool(annotations=_DESTRUCTIVE)(delete_contact)
    server.tool(annotations=_READ_ONLY)(list_contact_folders)
    server.tool(annotations=_READ_ONLY)(search_contacts)
    server.tool(annotations=_READ_ONLY)(get_contact_photo)
