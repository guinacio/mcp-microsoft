"""
Draft management tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.

Implemented:
  - create_draft
  - list_drafts
  - get_draft
  - update_draft
  - send_draft
"""

from __future__ import annotations

from typing import Literal, Optional, Union

from mcp.types import ToolAnnotations

from mcp_microsoft.models import (
    CreateDraftResponse,
    DraftDetailResponse,
    DraftSummary,
    ListDraftsResponse,
    SendDraftResponse,
    UpdateDraftResponse,
)
from mcp_microsoft.graph import get_graph
from mcp_microsoft.server import mcp
from mcp_microsoft.tools.mail import _fmt_date, _parse_recipients, _recipient_values

# ---------------------------------------------------------------------------
# create_draft
# ---------------------------------------------------------------------------

BodyType = Literal["text", "html"]
_READ_ONLY = ToolAnnotations(readOnlyHint=True, openWorldHint=True)
_WRITE = ToolAnnotations(destructiveHint=False, openWorldHint=True)

@mcp.tool(annotations=_WRITE)
async def create_draft(
    to: Union[str, list[str]],
    subject: str,
    body: str,
    cc: Optional[Union[str, list[str]]] = None,
    body_type: BodyType = "text",
    profile: str | None = None,
) -> CreateDraftResponse:
    """
    Create a new draft message (without sending).

    Args:
        to: Recipient address(es). Comma-separated string or list.
        subject: Draft subject line.
        body: Draft body content.
        cc: Optional CC address(es). Comma-separated string or list.
        body_type: 'text' or 'html'. Defaults to 'text'.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured draft creation confirmation.
    """
    g = get_graph(profile)
    message: dict = {
        "subject": subject,
        "body": {
            "contentType": "HTML" if body_type.lower() == "html" else "Text",
            "content": body,
        },
        "toRecipients": _parse_recipients(to),
    }
    if cc:
        message["ccRecipients"] = _parse_recipients(cc)

    result = await g.post("/me/messages", json=message)

    draft_id = (result or {}).get("id", "unknown")
    return CreateDraftResponse(
        success=True,
        action="create_draft",
        draft_id=draft_id,
        to=[addr.get("emailAddress", {}).get("address", "") for addr in message["toRecipients"]],
        cc=[addr.get("emailAddress", {}).get("address", "") for addr in message.get("ccRecipients", [])],
        subject=subject,
        body_type=body_type,
    )


# ---------------------------------------------------------------------------
# list_drafts
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def list_drafts(max_results: int = 10, profile: str | None = None) -> ListDraftsResponse:
    """
    List draft messages from the Drafts folder.

    Args:
        max_results: Maximum number of drafts to return (1-100). Defaults to 10.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured draft summaries.
    """
    g = get_graph(profile)
    params: dict = {
        "$top": max_results,
        "$select": "id,subject,toRecipients,lastModifiedDateTime,bodyPreview",
        "$orderby": "lastModifiedDateTime desc",
    }

    result = await g.get("/me/mailFolders/drafts/messages", params=params)
    drafts = result.get("value", [])

    items: list[DraftSummary] = []
    for draft in drafts:
        items.append(
            DraftSummary(
                id=draft.get("id", ""),
                subject=draft.get("subject") or "(no subject)",
                to=_recipient_values(draft.get("toRecipients", [])),
                last_modified_at=draft.get("lastModifiedDateTime"),
                last_modified_at_display=_fmt_date(draft.get("lastModifiedDateTime")),
                preview=(draft.get("bodyPreview") or "").replace("\n", " ")[:100],
            )
        )

    return ListDraftsResponse(count=len(items), drafts=items)


# ---------------------------------------------------------------------------
# get_draft
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def get_draft(draft_id: str, profile: str | None = None) -> DraftDetailResponse:
    """
    Fetch a draft message by ID.

    Args:
        draft_id: The Graph message ID of the draft.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured draft details.
    """
    g = get_graph(profile)
    params = {
        "$select": (
            "id,subject,from,toRecipients,ccRecipients,bccRecipients,"
            "lastModifiedDateTime,body,bodyPreview,isDraft"
        ),
    }

    draft = await g.get(f"/me/messages/{draft_id}", params=params)

    subject = draft.get("subject") or "(no subject)"
    modified = _fmt_date(draft.get("lastModifiedDateTime"))

    body_obj = draft.get("body") or {}
    content_type = (body_obj.get("contentType") or "text").lower()
    raw_body = body_obj.get("content", "")

    # Import _strip_html locally to avoid circular import issues at module level
    from mcp_microsoft.tools.mail import _strip_html

    if content_type == "html":
        body_text = _strip_html(raw_body)
    else:
        body_text = raw_body

    return DraftDetailResponse(
        id=draft_id,
        subject=subject,
        to=_recipient_values(draft.get("toRecipients", [])),
        cc=_recipient_values(draft.get("ccRecipients", [])),
        bcc=_recipient_values(draft.get("bccRecipients", [])),
        last_modified_at=draft.get("lastModifiedDateTime"),
        last_modified_at_display=modified,
        body=body_text,
        body_content_type=content_type,
        is_draft=draft.get("isDraft", True),
    )


# ---------------------------------------------------------------------------
# update_draft
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_WRITE)
async def update_draft(
    draft_id: str,
    subject: Optional[str] = None,
    body: Optional[str] = None,
    to: Optional[Union[str, list[str]]] = None,
    cc: Optional[Union[str, list[str]]] = None,
    body_type: BodyType = "text",
    profile: str | None = None,
) -> UpdateDraftResponse:
    """
    Update an existing draft message. Only provided fields are changed.

    Args:
        draft_id: The Graph message ID of the draft to update.
        subject: Replace subject line (omit to leave unchanged).
        body: Replace body content (omit to leave unchanged).
        to: Replace recipient list (omit to leave unchanged).
        cc: Replace CC list (omit to leave unchanged).
        body_type: 'text' or 'html'. Only used when body is provided. Defaults to 'text'.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured update confirmation.
    """
    g = get_graph(profile)
    patch: dict = {}

    if subject is not None:
        patch["subject"] = subject
    if body is not None:
        patch["body"] = {
            "contentType": "HTML" if body_type.lower() == "html" else "Text",
            "content": body,
        }
    if to is not None:
        patch["toRecipients"] = _parse_recipients(to)
    if cc is not None:
        patch["ccRecipients"] = _parse_recipients(cc)

    if not patch:
        return UpdateDraftResponse(
            success=False,
            action="update_draft",
            draft_id=draft_id,
            updated_fields=[],
            error="No fields to update.",
        )

    result = await g.patch(f"/me/messages/{draft_id}", json=patch)

    updated_id = (result or {}).get("id", draft_id)
    updated_fields = ", ".join(patch.keys())
    return UpdateDraftResponse(
        success=True,
        action="update_draft",
        draft_id=updated_id,
        updated_fields=list(patch.keys()),
        updated_fields_display=updated_fields,
    )


# ---------------------------------------------------------------------------
# send_draft
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_WRITE)
async def send_draft(draft_id: str, profile: str | None = None) -> SendDraftResponse:
    """
    Send an existing draft message.

    Args:
        draft_id: The Graph message ID of the draft to send.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured send confirmation.
    """
    g = get_graph(profile)
    await g.post(f"/me/messages/{draft_id}/send", json={})
    return SendDraftResponse(success=True, action="send_draft", draft_id=draft_id)
