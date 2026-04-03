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

from typing import Literal

from mcp_microsoft.common.mail_utils import format_mail_datetime, parse_recipients, recipient_values
from mcp_microsoft.common.text import strip_html
from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.tooling import READ_ONLY_TOOL, WRITE_TOOL, register_tool
from mcp_microsoft.models import (
    CreateDraftResponse,
    DraftDetailResponse,
    DraftSummary,
    ListDraftsResponse,
    SendDraftResponse,
    UpdateDraftResponse,
)
from mcp_microsoft.graph import get_graph

# ---------------------------------------------------------------------------
# create_draft
# ---------------------------------------------------------------------------

BodyType = Literal["text", "html"]


class CreateDraftInput(ToolRequestModel):
    to: str | list[str]
    subject: str
    body: str
    cc: str | list[str] | None = None
    body_type: BodyType = "text"
    profile: str | None = None


class ListDraftsInput(ToolRequestModel):
    max_results: int = 10
    profile: str | None = None


class GetDraftInput(ToolRequestModel):
    draft_id: str
    profile: str | None = None


class UpdateDraftInput(ToolRequestModel):
    draft_id: str
    subject: str | None = None
    body: str | None = None
    to: str | list[str] | None = None
    cc: str | list[str] | None = None
    body_type: BodyType = "text"
    profile: str | None = None


class SendDraftInput(ToolRequestModel):
    draft_id: str
    profile: str | None = None


async def create_draft(
    params: CreateDraftInput,
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
    g = get_graph(params.profile)
    message: dict = {
        "subject": params.subject,
        "body": {
            "contentType": "HTML" if params.body_type.lower() == "html" else "Text",
            "content": params.body,
        },
        "toRecipients": parse_recipients(params.to),
    }
    if params.cc:
        message["ccRecipients"] = parse_recipients(params.cc)

    result = await g.post("/me/messages", json=message)

    draft_id = (result or {}).get("id", "unknown")
    return CreateDraftResponse(
        success=True,
        action="create_draft",
        draft_id=draft_id,
        to=[addr.get("emailAddress", {}).get("address", "") for addr in message["toRecipients"]],
        cc=[addr.get("emailAddress", {}).get("address", "") for addr in message.get("ccRecipients", [])],
        subject=params.subject,
        body_type=params.body_type,
    )


# ---------------------------------------------------------------------------
# list_drafts
# ---------------------------------------------------------------------------


async def list_drafts(params: ListDraftsInput) -> ListDraftsResponse:
    """
    List draft messages from the Drafts folder.

    Args:
        max_results: Maximum number of drafts to return (1-100). Defaults to 10.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured draft summaries.
    """
    g = get_graph(params.profile)
    query: dict = {
        "$top": params.max_results,
        "$select": "id,subject,toRecipients,lastModifiedDateTime,bodyPreview",
        "$orderby": "lastModifiedDateTime desc",
    }

    result = await g.get("/me/mailFolders/drafts/messages", params=query)
    drafts = result.get("value", [])

    items: list[DraftSummary] = []
    for draft in drafts:
        items.append(
            DraftSummary(
                id=draft.get("id", ""),
                subject=draft.get("subject") or "(no subject)",
                to=recipient_values(draft.get("toRecipients", [])),
                last_modified_at=draft.get("lastModifiedDateTime"),
                last_modified_at_display=format_mail_datetime(draft.get("lastModifiedDateTime")),
                preview=(draft.get("bodyPreview") or "").replace("\n", " ")[:100],
            )
        )

    return ListDraftsResponse(count=len(items), drafts=items)


# ---------------------------------------------------------------------------
# get_draft
# ---------------------------------------------------------------------------


async def get_draft(params: GetDraftInput) -> DraftDetailResponse:
    """
    Fetch a draft message by ID.

    Args:
        draft_id: The Graph message ID of the draft.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured draft details.
    """
    g = get_graph(params.profile)
    query = {
        "$select": (
            "id,subject,from,toRecipients,ccRecipients,bccRecipients,"
            "lastModifiedDateTime,body,bodyPreview,isDraft"
        ),
    }

    draft = await g.get(f"/me/messages/{params.draft_id}", params=query)

    subject = draft.get("subject") or "(no subject)"
    modified = format_mail_datetime(draft.get("lastModifiedDateTime"))

    body_obj = draft.get("body") or {}
    content_type = (body_obj.get("contentType") or "text").lower()
    raw_body = body_obj.get("content", "")

    if content_type == "html":
        body_text = strip_html(raw_body)
    else:
        body_text = raw_body

    return DraftDetailResponse(
        id=params.draft_id,
        subject=subject,
        to=recipient_values(draft.get("toRecipients", [])),
        cc=recipient_values(draft.get("ccRecipients", [])),
        bcc=recipient_values(draft.get("bccRecipients", [])),
        last_modified_at=draft.get("lastModifiedDateTime"),
        last_modified_at_display=modified,
        body=body_text,
        body_content_type=content_type,
        is_draft=draft.get("isDraft", True),
    )


# ---------------------------------------------------------------------------
# update_draft
# ---------------------------------------------------------------------------


async def update_draft(
    params: UpdateDraftInput,
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
    g = get_graph(params.profile)
    patch: dict = {}

    if params.subject is not None:
        patch["subject"] = params.subject
    if params.body is not None:
        patch["body"] = {
            "contentType": "HTML" if params.body_type.lower() == "html" else "Text",
            "content": params.body,
        }
    if params.to is not None:
        patch["toRecipients"] = parse_recipients(params.to)
    if params.cc is not None:
        patch["ccRecipients"] = parse_recipients(params.cc)

    if not patch:
        return UpdateDraftResponse(
            success=False,
            action="update_draft",
            draft_id=params.draft_id,
            updated_fields=[],
            error="No fields to update.",
        )

    result = await g.patch(f"/me/messages/{params.draft_id}", json=patch)

    updated_id = (result or {}).get("id", params.draft_id)
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


async def send_draft(params: SendDraftInput) -> SendDraftResponse:
    """
    Send an existing draft message.

    Args:
        draft_id: The Graph message ID of the draft to send.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured send confirmation.
    """
    g = get_graph(params.profile)
    await g.post(f"/me/messages/{params.draft_id}/send", json={})
    return SendDraftResponse(success=True, action="send_draft", draft_id=params.draft_id)


def register(server) -> None:
    """Register all draft tools with the given FastMCP server instance."""
    register_tool(server, create_draft, annotations=WRITE_TOOL)
    register_tool(server, list_drafts, annotations=READ_ONLY_TOOL)
    register_tool(server, get_draft, annotations=READ_ONLY_TOOL)
    register_tool(server, update_draft, annotations=WRITE_TOOL)
    register_tool(server, send_draft, annotations=WRITE_TOOL)
