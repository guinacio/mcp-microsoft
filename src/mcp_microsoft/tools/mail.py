"""
Core mail tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.

Implemented:
  - list_emails
  - read_email
  - search_emails
  - filter_emails
  - send_email
  - reply_email
  - forward_email
  - mark_as_read
  - mark_as_unread
  - move_email
  - trash_email
  - delete_email
  - bulk_move_emails
  - bulk_trash_emails
  - bulk_delete_emails
"""

from __future__ import annotations

import html as html_module
import re
from datetime import datetime
from typing import Any, Literal, Optional, Union

from mcp.server.fastmcp import Context
from mcp.types import ToolAnnotations
from pydantic import BaseModel

from mcp_microsoft.models import (
    Address,
    AttachmentInfo,
    BulkDeleteEmailsResponse,
    BulkEmailFailure,
    BulkMovedEmail,
    BulkMoveEmailsResponse,
    BulkTrashEmailsResponse,
    DeleteEmailResponse,
    DisplayAddress,
    ForwardEmailResponse,
    ListEmailsResponse,
    MarkEmailReadResponse,
    MessageSummary,
    MoveEmailResponse,
    ReadEmailResponse,
    ReadEmailSummaryResponse,
    ReplyEmailResponse,
    SearchEmailsResponse,
    SendEmailResponse,
    TrashEmailResponse,
)
from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.graph import get_graph

# ---------------------------------------------------------------------------
# Elicitation helpers
# ---------------------------------------------------------------------------


class _Confirmation(BaseModel):
    confirmed: bool


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

BodyType = Literal["text", "html"]
_READ_ONLY = ToolAnnotations(readOnlyHint=True, openWorldHint=True)
_WRITE = ToolAnnotations(destructiveHint=False, openWorldHint=True)
_IDEMPOTENT_WRITE = ToolAnnotations(
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=True,
)
_DESTRUCTIVE = ToolAnnotations(destructiveHint=True, openWorldHint=True)


def _parse_recipients(
    value: Optional[Union[str, list[str]]],
) -> list[dict]:
    """
    Convert a comma-separated string or list of addresses to Graph recipient format.

    Example:
        "alice@example.com, bob@example.com"
        -> [{"emailAddress": {"address": "alice@example.com"}}, ...]
    """
    if not value:
        return []
    if isinstance(value, str):
        addresses = [addr.strip() for addr in value.split(",") if addr.strip()]
    else:
        addresses = [addr.strip() for addr in value if addr.strip()]
    return [{"emailAddress": {"address": addr}} for addr in addresses]


def _strip_html(html: str) -> str:
    """
    Convert an HTML string to plain text by stripping tags and unescaping entities.
    """
    # Remove <style> and <script> blocks entirely
    text = re.sub(r"<(style|script)[^>]*>.*?</\1>", "", html, flags=re.DOTALL | re.IGNORECASE)
    # Replace block-level line breaks
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</p>", "\n\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</?(div|tr|li|h[1-6]|blockquote)[^>]*>", "\n", text, flags=re.IGNORECASE)
    # Strip remaining tags
    text = re.sub(r"<[^>]+>", "", text)
    # Unescape HTML entities
    text = html_module.unescape(text)
    # Collapse excessive whitespace / blank lines
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _fmt_date(iso: Optional[str]) -> str:
    """Format an ISO 8601 date string to a human-readable form."""
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return iso


def _fmt_sender(sender_obj: Optional[dict]) -> str:
    """Format a Graph sender/from object as 'Name <email>' or just 'email'."""
    if not sender_obj:
        return "unknown"
    ea = sender_obj.get("emailAddress", {})
    name = ea.get("name", "")
    addr = ea.get("address", "unknown")
    return f"{name} <{addr}>" if name else addr


def _fmt_recipients(recipients: list[dict]) -> str:
    """Format a list of Graph recipient objects as a comma-separated string."""
    parts = []
    for r in recipients or []:
        ea = r.get("emailAddress", {})
        name = ea.get("name", "")
        addr = ea.get("address", "")
        parts.append(f"{name} <{addr}>" if name else addr)
    return ", ".join(parts)


def _recipient_values(recipients: list[dict]) -> list[Address]:
    """Normalize Graph recipient objects into simple dictionaries."""
    values: list[Address] = []
    for recipient in recipients or []:
        email_address = recipient.get("emailAddress", {})
        values.append(Address(name=email_address.get("name", ""), address=email_address.get("address", "")))
    return values


def _display_address_from_sender(sender_obj: dict[str, Any] | None) -> DisplayAddress:
    """Normalize a Graph sender into a typed address model."""
    email_address = (sender_obj or {}).get("emailAddress", {})
    return DisplayAddress(
        display=_fmt_sender(sender_obj),
        name=email_address.get("name", ""),
        address=email_address.get("address", ""),
    )


def _message_summary(msg: dict[str, Any]) -> MessageSummary:
    """Normalize a Graph message into a summary payload."""
    return MessageSummary(
        id=msg.get("id", ""),
        subject=msg.get("subject") or "(no subject)",
        from_=_display_address_from_sender(msg.get("from")),
        received_at=msg.get("receivedDateTime"),
        received_at_display=_fmt_date(msg.get("receivedDateTime")),
        is_read=msg.get("isRead", False),
        has_attachments=msg.get("hasAttachments", False),
        importance=msg.get("importance", ""),
        preview=(msg.get("bodyPreview") or "").replace("\n", " ")[:120],
    )


# ---------------------------------------------------------------------------
# list_emails
# ---------------------------------------------------------------------------


async def list_emails(
    folder: str = "inbox",
    max_results: int = 10,
    unread_only: bool = False,
    sort_order: Literal["newest", "oldest"] = "newest",
    skip_token: Optional[str] = None,
    profile: str | None = None,
) -> ListEmailsResponse:
    """
    List emails from a mail folder.

    Args:
        folder: Well-known folder name or folder ID.
                Well-known names: inbox, sentitems, drafts, deleteditems,
                junkemail, archive. Defaults to 'inbox'.
        max_results: Maximum number of messages to return (1-100). Defaults to 10.
        unread_only: When True, return only unread messages. Defaults to False.
        sort_order: 'newest' (default) or 'oldest' first.
        skip_token: Opaque pagination cursor returned as next_page_token from a
                    previous call. Omit for the first page.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured message summaries with pagination metadata. When has_more is
        True, pass next_page_token as skip_token to retrieve the next page.
    """
    from urllib.parse import parse_qs, urlparse

    g = get_graph(profile)
    order = "receivedDateTime asc" if sort_order == "oldest" else "receivedDateTime desc"
    params: dict = {
        "$top": max_results,
        "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments,importance",
        "$orderby": order,
    }
    if unread_only:
        params["$filter"] = "isRead eq false"
    if skip_token is not None:
        params["$skiptoken"] = skip_token

    result = await g.get(f"/me/mailFolders/{folder}/messages", params=params)

    messages = result.get("value", [])
    next_link = result.get("@odata.nextLink", "")

    next_page_token: str | None = None
    if next_link:
        qs = parse_qs(urlparse(next_link).query)
        next_page_token = qs.get("$skiptoken", [None])[0]

    return ListEmailsResponse(
        folder=folder,
        count=len(messages),
        messages=[_message_summary(msg) for msg in messages],
        next_page_token=next_page_token,
        has_more=(next_page_token is not None),
    )


# ---------------------------------------------------------------------------
# read_email
# ---------------------------------------------------------------------------


async def read_email(
    message_id: str,
    summary_mode: bool = False,
    profile: str | None = None,
) -> ReadEmailResponse | ReadEmailSummaryResponse:
    """
    Fetch a full email message by ID.

    Args:
        message_id: The Graph message ID.
        summary_mode: When True, return only subject, from, date, and body preview
                      instead of the full body.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured message details.
    """
    g = get_graph(profile)
    params = {
        "$select": (
            "id,subject,from,toRecipients,ccRecipients,receivedDateTime,"
            "body,bodyPreview,attachments,isRead,conversationId,importance"
        ),
        "$expand": "attachments($select=id,name,size,contentType)",
    }

    msg = await g.get(f"/me/messages/{message_id}", params=params)

    subject = msg.get("subject") or "(no subject)"
    sender = _fmt_sender(msg.get("from"))
    to_str = _fmt_recipients(msg.get("toRecipients", []))
    cc_str = _fmt_recipients(msg.get("ccRecipients", []))
    date = _fmt_date(msg.get("receivedDateTime"))
    is_read = msg.get("isRead", False)
    conv_id = msg.get("conversationId", "")

    if summary_mode:
        return ReadEmailSummaryResponse(
            id=message_id,
            subject=subject,
            from_=_display_address_from_sender(msg.get("from")),
            received_at=msg.get("receivedDateTime"),
            received_at_display=date,
            is_read=is_read,
            preview=(msg.get("bodyPreview") or "").replace("\n", " "),
        )

    # Full mode — extract body text
    body_obj = msg.get("body") or {}
    content_type = (body_obj.get("contentType") or "text").lower()
    raw_body = body_obj.get("content", "")
    if content_type == "html":
        body_text = _strip_html(raw_body)
    else:
        body_text = raw_body

    # Attachments
    attachments = msg.get("attachments", []) or []
    attachment_items: list[AttachmentInfo] = []
    for att in attachments:
        size_kb = (att.get("size") or 0) // 1024
        attachment_items.append(
            AttachmentInfo(
                id=att.get("id", ""),
                name=att.get("name", "unnamed"),
                content_type=att.get("contentType", ""),
                size_bytes=att.get("size", 0),
                size_kb=size_kb,
            )
        )

    return ReadEmailResponse(
        id=message_id,
        subject=subject,
        from_=_display_address_from_sender(msg.get("from")),
        to=_recipient_values(msg.get("toRecipients", [])),
        to_display=to_str,
        cc=_recipient_values(msg.get("ccRecipients", [])),
        cc_display=cc_str,
        received_at=msg.get("receivedDateTime"),
        received_at_display=date,
        is_read=is_read,
        conversation_id=conv_id,
        importance=msg.get("importance", ""),
        body=body_text,
        body_content_type=content_type,
        attachments=attachment_items,
    )


# ---------------------------------------------------------------------------
# search_emails
# ---------------------------------------------------------------------------


async def search_emails(
    query: str,
    max_results: int = 10,
    folder: Optional[str] = None,
    profile: str | None = None,
) -> SearchEmailsResponse:
    """
    Search messages using Graph KQL $search syntax.

    Note: Graph $search and $filter cannot be combined in the same request.
    The Graph API caps $search results at 25 regardless of the value requested.

    Args:
        query: KQL search string, e.g. 'from:alice@example.com' or 'project update'.
        max_results: Maximum number of results (1-25). Values above 25 are
            silently capped by the Graph API. Defaults to 10.
        folder: Optional well-known folder name or folder ID to restrict the search.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured search results.
    """
    g = get_graph(profile)
    params: dict = {
        "$search": f'"{query}"',
        "$top": min(max_results, 25),
        "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments",
    }

    if folder:
        path = f"/me/mailFolders/{folder}/messages"
    else:
        path = "/me/messages"

    result = await g.get(path, params=params)
    messages = result.get("value", [])

    return SearchEmailsResponse(
        query=query,
        folder=folder,
        count=len(messages),
        messages=[_message_summary(msg) for msg in messages],
    )


# ---------------------------------------------------------------------------
# filter_emails
# ---------------------------------------------------------------------------


async def filter_emails(
    from_address: Optional[str] = None,
    to_address: Optional[str] = None,
    subject_contains: Optional[str] = None,
    received_after: Optional[str] = None,
    received_before: Optional[str] = None,
    has_attachments: Optional[bool] = None,
    importance: Optional[Literal["low", "normal", "high"]] = None,
    folder: str = "inbox",
    max_results: int = 50,
    sort_order: Literal["newest", "oldest"] = "newest",
    skip_token: Optional[str] = None,
    profile: str | None = None,
) -> ListEmailsResponse:
    """
    Find emails matching specific criteria using OData $filter.

    Unlike search_emails (which is limited to 25 results), this tool
    supports up to 100 results per page with full pagination — ideal for
    finding all emails from a sender, within a date range, or matching
    a subject.

    All filter parameters are combined with AND logic. Omitted parameters
    are not filtered on.

    Args:
        from_address: Filter by sender email address (exact match).
        to_address: Filter by recipient email address (exact match).
        subject_contains: Filter by subject containing this text (case-insensitive).
        received_after: Only messages received on or after this date.
            ISO 8601 format: '2026-01-01' or '2026-01-01T00:00:00Z'.
        received_before: Only messages received before this date.
            ISO 8601 format: '2026-03-31' or '2026-03-31T23:59:59Z'.
        has_attachments: When True, only messages with attachments.
            When False, only messages without.
        importance: Filter by importance level: 'low', 'normal', or 'high'.
        folder: Well-known folder name or folder ID. Defaults to 'inbox'.
        max_results: Maximum number of messages to return (1-100). Defaults to 50.
        sort_order: 'newest' (default) or 'oldest' first.
        skip_token: Opaque pagination cursor returned as next_page_token from a
                    previous call. Omit for the first page.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured message summaries with pagination metadata.
    """
    g = get_graph(profile)
    order = "receivedDateTime asc" if sort_order == "oldest" else "receivedDateTime desc"
    params: dict = {
        "$top": min(max(1, max_results), 100),
        "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments,importance",
        "$orderby": order,
    }

    # Build OData $filter clauses
    clauses: list[str] = []
    if from_address:
        safe = from_address.replace("'", "''")
        clauses.append(f"from/emailAddress/address eq '{safe}'")
    if to_address:
        safe = to_address.replace("'", "''")
        clauses.append(f"toRecipients/any(r:r/emailAddress/address eq '{safe}')")
    if subject_contains:
        safe = subject_contains.replace("'", "''")
        clauses.append(f"contains(subject, '{safe}')")
    if received_after:
        # Append time component if only a date was given
        ts = received_after if "T" in received_after else f"{received_after}T00:00:00Z"
        clauses.append(f"receivedDateTime ge {ts}")
    if received_before:
        ts = received_before if "T" in received_before else f"{received_before}T23:59:59Z"
        clauses.append(f"receivedDateTime lt {ts}")
    if has_attachments is not None:
        clauses.append(f"hasAttachments eq {str(has_attachments).lower()}")
    if importance:
        clauses.append(f"importance eq '{importance}'")

    if clauses:
        params["$filter"] = " and ".join(clauses)

    if skip_token is not None:
        params["$skiptoken"] = skip_token

    result = await g.get(f"/me/mailFolders/{folder}/messages", params=params)

    messages = result.get("value", [])
    next_link = result.get("@odata.nextLink")

    from urllib.parse import parse_qs, urlparse
    next_page_token: str | None = None
    if next_link:
        qs = parse_qs(urlparse(next_link).query)
        next_page_token = qs.get("$skiptoken", [None])[0]

    return ListEmailsResponse(
        folder=folder,
        count=len(messages),
        messages=[_message_summary(msg) for msg in messages],
        next_page_token=next_page_token,
        has_more=(next_page_token is not None),
    )


# ---------------------------------------------------------------------------
# send_email
# ---------------------------------------------------------------------------


class SendEmailInput(ToolRequestModel):
    """Validated input for the send_email tool."""

    to: Union[str, list[str]]
    subject: str
    body: str
    cc: Optional[Union[str, list[str]]] = None
    bcc: Optional[Union[str, list[str]]] = None
    body_type: BodyType = "text"
    save_to_sent: bool = True
    reply_to: Optional[Union[str, list[str]]] = None
    profile: str | None = None


async def send_email(
    to: Union[str, list[str]],
    subject: str,
    body: str,
    cc: Optional[Union[str, list[str]]] = None,
    bcc: Optional[Union[str, list[str]]] = None,
    body_type: BodyType = "text",
    save_to_sent: bool = True,
    reply_to: Optional[Union[str, list[str]]] = None,
    profile: str | None = None,
    confirm: bool = False,
    ctx: Context | None = None,
) -> SendEmailResponse:
    """
    Send a new email message.

    Args:
        to: Recipient address(es). Comma-separated string or list.
        subject: Email subject line.
        body: Message body content.
        cc: Optional CC address(es). Comma-separated string or list.
        bcc: Optional BCC address(es). Comma-separated string or list.
        body_type: 'text' or 'html'. Defaults to 'text'.
        save_to_sent: When True (default), save a copy in Sent Items.
        reply_to: Optional reply-to address(es).
        profile: Microsoft 365 profile to use. Omit to use the default profile.
        confirm: When True, prompt the user to confirm before sending. Defaults to False.

    Returns:
        Structured send confirmation.
    """
    if confirm and ctx:
        to_display = to if isinstance(to, str) else ", ".join(to)
        preview = f"To: {to_display}
Subject: {subject}

{body[:200]}{'...' if len(body) > 200 else ''}"
        result = await ctx.elicit(
            f"Send this email?

{preview}",
            response_type=_Confirmation,
        )
        if result.action != "accept" or not result.data.confirmed:
            return SendEmailResponse(success=False, action="send_email", error="Cancelled by user.")

    p = SendEmailInput.model_validate({
        "to": to, "subject": subject, "body": body,
        "cc": cc, "bcc": bcc, "body_type": body_type,
        "save_to_sent": save_to_sent, "reply_to": reply_to, "profile": profile,
    })
    g = get_graph(p.profile)
    message: dict = {
        "subject": p.subject,
        "body": {
            "contentType": "HTML" if p.body_type.lower() == "html" else "Text",
            "content": p.body,
        },
        "toRecipients": _parse_recipients(p.to),
    }

    if p.cc:
        message["ccRecipients"] = _parse_recipients(p.cc)
    if p.bcc:
        message["bccRecipients"] = _parse_recipients(p.bcc)
    if p.reply_to:
        message["replyTo"] = _parse_recipients(p.reply_to)

    payload = {
        "message": message,
        "saveToSentItems": p.save_to_sent,
    }

    await g.post("/me/sendMail", json=payload)

    return SendEmailResponse(
        success=True,
        action="send_email",
        to=[addr.get("emailAddress", {}).get("address", "") for addr in message["toRecipients"]],
        cc=[addr.get("emailAddress", {}).get("address", "") for addr in message.get("ccRecipients", [])],
        bcc=[addr.get("emailAddress", {}).get("address", "") for addr in message.get("bccRecipients", [])],
        reply_to=[addr.get("emailAddress", {}).get("address", "") for addr in message.get("replyTo", [])],
        subject=p.subject,
        body_type=p.body_type,
        saved_to_sent_items=p.save_to_sent,
    )


# ---------------------------------------------------------------------------
# reply_email
# ---------------------------------------------------------------------------


async def reply_email(
    message_id: str,
    body: str,
    reply_all: bool = False,
    body_type: BodyType = "text",
    profile: str | None = None,
) -> ReplyEmailResponse:
    """
    Reply to an existing email message.

    Args:
        message_id: The Graph message ID to reply to.
        body: Reply body text or HTML.
        reply_all: When True, reply to all recipients. Defaults to False.
        body_type: 'text' or 'html'. Defaults to 'text'.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured reply confirmation.
    """
    g = get_graph(profile)
    endpoint = "replyAll" if reply_all else "reply"
    if body_type.lower() == "html":
        payload = {
            "message": {
                "body": {
                    "contentType": "HTML",
                    "content": body,
                }
            }
        }
    else:
        payload = {
            "message": {},
            "comment": body,
        }

    await g.post(f"/me/messages/{message_id}/{endpoint}", json=payload)

    return ReplyEmailResponse(
        success=True,
        action="reply_all" if reply_all else "reply",
        message_id=message_id,
        body_type=body_type,
    )


# ---------------------------------------------------------------------------
# forward_email
# ---------------------------------------------------------------------------


async def forward_email(
    message_id: str,
    to: Union[str, list[str]],
    comment: Optional[str] = None,
    profile: str | None = None,
) -> ForwardEmailResponse:
    """
    Forward an email message to one or more recipients.

    Args:
        message_id: The Graph message ID to forward.
        to: Recipient address(es). Comma-separated string or list.
        comment: Optional comment to prepend to the forwarded message.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured forward confirmation.
    """
    g = get_graph(profile)
    payload: dict = {
        "toRecipients": _parse_recipients(to),
        "comment": comment or "",
    }

    await g.post(f"/me/messages/{message_id}/forward", json=payload)

    return ForwardEmailResponse(
        success=True,
        action="forward",
        message_id=message_id,
        to=[addr.get("emailAddress", {}).get("address", "") for addr in payload["toRecipients"]],
        comment=comment or "",
    )


# ---------------------------------------------------------------------------
# mark_as_read / mark_as_unread
# ---------------------------------------------------------------------------


async def mark_as_read(message_id: str, profile: str | None = None) -> MarkEmailReadResponse:
    """
    Mark a message as read.

    Args:
        message_id: The Graph message ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured update confirmation.
    """
    g = get_graph(profile)
    await g.patch(f"/me/messages/{message_id}", json={"isRead": True})
    return MarkEmailReadResponse(success=True, action="mark_as_read", message_id=message_id, is_read=True)


async def mark_as_unread(message_id: str, profile: str | None = None) -> MarkEmailReadResponse:
    """
    Mark a message as unread.

    Args:
        message_id: The Graph message ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured update confirmation.
    """
    g = get_graph(profile)
    await g.patch(f"/me/messages/{message_id}", json={"isRead": False})
    return MarkEmailReadResponse(success=True, action="mark_as_unread", message_id=message_id, is_read=False)


# ---------------------------------------------------------------------------
# move_email
# ---------------------------------------------------------------------------


async def move_email(message_id: str, destination_folder: str, profile: str | None = None) -> MoveEmailResponse:
    """
    Move a message to a different mail folder.

    Args:
        message_id: The Graph message ID to move.
        destination_folder: Target folder — well-known name (e.g. 'archive',
            'inbox', 'junkemail', 'deleteditems') or opaque folder ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured move confirmation.
    """
    g = get_graph(profile)
    result = await g.post(
        f"/me/messages/{message_id}/move",
        json={"destinationId": destination_folder},
    )
    new_id = (result or {}).get("id", message_id)
    return MoveEmailResponse(
        success=True,
        action="move",
        message_id=message_id,
        new_message_id=new_id,
        destination_folder=destination_folder,
    )


# ---------------------------------------------------------------------------
# trash_email
# ---------------------------------------------------------------------------


async def trash_email(message_id: str, profile: str | None = None) -> TrashEmailResponse:
    """
    Move a message to the Deleted Items folder (soft delete / recoverable).

    To permanently delete without recovery, use delete_email instead.

    Args:
        message_id: The Graph message ID to trash.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured trash confirmation.
    """
    g = get_graph(profile)
    result = await g.post(
        f"/me/messages/{message_id}/move",
        json={"destinationId": "deleteditems"},
    )
    new_id = (result or {}).get("id", message_id)
    return TrashEmailResponse(
        success=True,
        action="trash",
        message_id=message_id,
        new_message_id=new_id,
        destination_folder="deleteditems",
        soft_delete=True,
        profile=profile,
    )


# ---------------------------------------------------------------------------
# delete_email
# ---------------------------------------------------------------------------


async def delete_email(
    message_id: str,
    profile: str | None = None,
    confirm: bool = False,
    ctx: Context | None = None,
) -> DeleteEmailResponse:
    """
    Permanently delete a message from the mailbox. This action is IRREVERSIBLE.

    The message will be hard-deleted and cannot be recovered from Deleted Items.
    For a recoverable soft delete, use trash_email instead.

    Args:
        message_id: The Graph message ID to permanently delete.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
        confirm: When True, prompt the user to confirm before deleting. Defaults to False.

    Returns:
        Structured delete confirmation.
    """
    if confirm and ctx:
        result = await ctx.elicit(
            f"Permanently delete this email? This action is IRREVERSIBLE.\n\nMessage ID: {message_id}",
            response_type=_Confirmation,
        )
        if result.action != "accept" or not result.data.confirmed:
            return DeleteEmailResponse(success=False, action="permanent_delete", message_id=message_id, error="Cancelled by user.", irreversible=True)

    g = get_graph(profile)
    await g.post(f"/me/messages/{message_id}/permanentDelete")
    return DeleteEmailResponse(
        success=True,
        action="permanent_delete",
        message_id=message_id,
        irreversible=True,
    )


# ---------------------------------------------------------------------------
# Batch helper (Graph $batch endpoint, max 20 requests per batch)
# ---------------------------------------------------------------------------


def _build_batch_requests(
    message_ids: list[str],
    method: str,
    url_template: str,
    body: dict[str, Any] | None = None,
) -> tuple[list[dict[str, Any]], dict[str, str]]:
    """
    Build Graph batch request entries with synthetic IDs.

    Returns (requests_list, id_map) where id_map maps synthetic "1","2",…
    back to the original message_id.
    """
    id_map: dict[str, str] = {}
    requests: list[dict[str, Any]] = []
    for idx, mid in enumerate(message_ids):
        batch_id = str(idx + 1)
        id_map[batch_id] = mid
        entry: dict[str, Any] = {
            "id": batch_id,
            "method": method,
            "url": url_template.format(mid=mid),
        }
        if body is not None:
            entry["body"] = body
            entry["headers"] = {"Content-Type": "application/json"}
        requests.append(entry)
    return requests, id_map


def _parse_batch_error(resp: dict[str, Any]) -> tuple[str | None, str]:
    """Extract error code and message from a batch response item."""
    status = resp.get("status", 0)
    body = resp.get("body")
    if not isinstance(body, dict):
        return None, f"HTTP {status}"
    error = body.get("error")
    if not isinstance(error, dict):
        return None, f"HTTP {status}"
    code = error.get("code")
    message = error.get("message") or f"HTTP {status}"
    return code, message


async def _execute_batch(
    g: Any,
    requests: list[dict[str, Any]],
    id_map: dict[str, str],
) -> tuple[int, list[BulkEmailFailure], dict[str, dict[str, Any]]]:
    """
    Send requests via GraphClient.batch() (which handles $batch chunking).

    Returns (succeeded_count, failures_list, success_bodies) where
    success_bodies maps original message_id to the response body dict.
    """
    succeeded = 0
    failures: list[BulkEmailFailure] = []
    success_bodies: dict[str, dict[str, Any]] = {}

    # Build a set of all expected batch IDs so we can detect missing responses.
    expected_ids = {r["id"] for r in requests}

    try:
        responses = await g.batch(requests)
    except Exception as exc:
        # The entire batch call failed — record a failure for every request.
        for req in requests:
            mid = id_map.get(req["id"], req["id"])
            failures.append(BulkEmailFailure(
                message_id=mid, status=0, error=str(exc),
            ))
        return succeeded, failures, success_bodies

    seen_ids: set[str] = set()
    for resp in responses:
        batch_id = resp.get("id", "")
        seen_ids.add(batch_id)
        mid = id_map.get(batch_id, batch_id)
        status = resp.get("status", 0)

        if 200 <= status < 300:
            succeeded += 1
            body = resp.get("body")
            if isinstance(body, dict):
                success_bodies[mid] = body
        else:
            code, error_msg = _parse_batch_error(resp)
            failures.append(BulkEmailFailure(
                message_id=mid, status=status, code=code, error=error_msg,
            ))

    # Detect missing responses (Graph returned fewer items than we sent).
    missing = expected_ids - seen_ids
    for batch_id in missing:
        mid = id_map.get(batch_id, batch_id)
        failures.append(BulkEmailFailure(
            message_id=mid, status=0, error="No response from batch",
        ))

    return succeeded, failures, success_bodies


# ---------------------------------------------------------------------------
# bulk_move_emails
# ---------------------------------------------------------------------------


async def _collect_folder_message_ids(g: Any, folder: str) -> list[str]:
    """Fetch all message IDs from a mail folder, paginating as needed."""
    ids: list[str] = []
    path = f"/me/mailFolders/{folder}/messages"
    params: dict[str, Any] = {"$top": 100, "$select": "id"}
    while path:
        result = await g.get(path, params=params)
        for msg in result.get("value", []):
            mid = msg.get("id")
            if mid:
                ids.append(mid)
        next_link = result.get("@odata.nextLink")
        if next_link:
            # nextLink is a full URL; extract relative path + query
            import re as _re
            m = _re.search(r"v1\.0(/.+)", next_link)
            path = m.group(1) if m else ""
            params = {}  # params are embedded in nextLink
        else:
            path = ""
    return ids


async def bulk_move_emails(
    message_ids: Optional[list[str]] = None,
    destination_folder: str = "",
    source_folder: Optional[str] = None,
    profile: str | None = None,
) -> BulkMoveEmailsResponse:
    """
    Move multiple messages to a destination folder in one operation.

    Uses the Graph batch API for efficiency (up to 20 per round-trip).

    Two modes:
    1. Pass message_ids explicitly.
    2. Pass source_folder to move ALL messages from that folder (e.g. 'junkemail').

    Args:
        message_ids: List of Graph message IDs to move. Optional if source_folder is set.
        destination_folder: Target folder — well-known name (e.g. 'archive',
            'inbox', 'junkemail', 'deleteditems') or opaque folder ID.
        source_folder: Move all messages from this folder instead of specifying IDs.
            Well-known names: inbox, sentitems, drafts, deleteditems, junkemail, archive.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Summary with success/failure counts, new message IDs, and failure details.
    """
    if not destination_folder:
        return BulkMoveEmailsResponse(success=False, action="bulk_move", error="destination_folder is required.")

    g = get_graph(profile)

    if source_folder and not message_ids:
        message_ids = await _collect_folder_message_ids(g, source_folder)

    if not message_ids:
        return BulkMoveEmailsResponse(success=True, action="bulk_move", destination_folder=destination_folder, total=0, succeeded=0, failed=0)

    requests, id_map = _build_batch_requests(
        message_ids,
        method="POST",
        url_template="/me/messages/{mid}/move",
        body={"destinationId": destination_folder},
    )

    succeeded, failures, success_bodies = await _execute_batch(g, requests, id_map)

    moved = [
        BulkMovedEmail(
            source_message_id=mid,
            new_message_id=body.get("id", ""),
        )
        for mid, body in success_bodies.items()
    ]

    return BulkMoveEmailsResponse(
        success=len(failures) == 0,
        action="bulk_move",
        destination_folder=destination_folder,
        total=len(message_ids),
        succeeded=succeeded,
        failed=len(failures),
        moved=moved,
        failures=failures,
    )


# ---------------------------------------------------------------------------
# bulk_trash_emails
# ---------------------------------------------------------------------------


async def bulk_trash_emails(
    message_ids: Optional[list[str]] = None,
    folder: Optional[str] = None,
    profile: str | None = None,
) -> BulkTrashEmailsResponse:
    """
    Move multiple messages to Deleted Items (soft delete / recoverable).

    Uses the Graph batch API for efficiency (up to 20 per round-trip).
    For permanent deletion, use bulk_delete_emails instead.

    Two modes:
    1. Pass message_ids explicitly.
    2. Pass folder to trash ALL messages from that folder (e.g. 'junkemail').

    Args:
        message_ids: List of Graph message IDs to trash. Optional if folder is set.
        folder: Trash all messages from this folder instead of specifying IDs.
            Well-known names: inbox, sentitems, drafts, junkemail, archive.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Summary with success/failure counts, new message IDs, and failure details.
    """
    g = get_graph(profile)

    if folder and not message_ids:
        message_ids = await _collect_folder_message_ids(g, folder)

    if not message_ids:
        return BulkTrashEmailsResponse(success=True, action="bulk_trash", total=0, succeeded=0, failed=0)

    requests, id_map = _build_batch_requests(
        message_ids,
        method="POST",
        url_template="/me/messages/{mid}/move",
        body={"destinationId": "deleteditems"},
    )

    succeeded, failures, success_bodies = await _execute_batch(g, requests, id_map)

    moved = [
        BulkMovedEmail(
            source_message_id=mid,
            new_message_id=body.get("id", ""),
        )
        for mid, body in success_bodies.items()
    ]

    return BulkTrashEmailsResponse(
        success=len(failures) == 0,
        action="bulk_trash",
        total=len(message_ids),
        succeeded=succeeded,
        failed=len(failures),
        moved=moved,
        failures=failures,
    )


# ---------------------------------------------------------------------------
# bulk_delete_emails
# ---------------------------------------------------------------------------


async def bulk_delete_emails(
    message_ids: Optional[list[str]] = None,
    folder: Optional[str] = None,
    profile: str | None = None,
) -> BulkDeleteEmailsResponse:
    """
    Permanently delete multiple messages from the mailbox. This action is IRREVERSIBLE.

    Uses the Graph batch API for efficiency (up to 20 per round-trip).
    Messages will be hard-deleted and cannot be recovered from Deleted Items.
    For a recoverable soft delete, use bulk_trash_emails instead.

    Two modes:
    1. Pass message_ids explicitly.
    2. Pass folder to permanently delete ALL messages from that folder.

    Args:
        message_ids: List of Graph message IDs to permanently delete.
            Optional if folder is set.
        folder: Permanently delete all messages from this folder instead of
            specifying IDs. Well-known names: inbox, sentitems, drafts,
            deleteditems, junkemail, archive.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Summary with success/failure counts and failure details.
    """
    g = get_graph(profile)

    if folder and not message_ids:
        message_ids = await _collect_folder_message_ids(g, folder)

    if not message_ids:
        return BulkDeleteEmailsResponse(success=True, action="bulk_permanent_delete", total=0, succeeded=0, failed=0, irreversible=True)

    requests, id_map = _build_batch_requests(
        message_ids,
        method="POST",
        url_template="/me/messages/{mid}/permanentDelete",
    )

    succeeded, failures, _ = await _execute_batch(g, requests, id_map)

    return BulkDeleteEmailsResponse(
        success=len(failures) == 0,
        action="bulk_permanent_delete",
        total=len(message_ids),
        succeeded=succeeded,
        failed=len(failures),
        irreversible=True,
        failures=failures,
    )


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------


def register(server) -> None:
    """Register all mail tools with the given FastMCP server instance."""
    server.tool(annotations=_READ_ONLY)(list_emails)
    server.tool(annotations=_READ_ONLY)(read_email)
    server.tool(annotations=_READ_ONLY)(search_emails)
    server.tool(annotations=_READ_ONLY)(filter_emails)
    server.tool(annotations=_WRITE)(send_email)
    server.tool(annotations=_WRITE)(reply_email)
    server.tool(annotations=_WRITE)(forward_email)
    server.tool(annotations=_IDEMPOTENT_WRITE)(mark_as_read)
    server.tool(annotations=_IDEMPOTENT_WRITE)(mark_as_unread)
    server.tool(annotations=_WRITE)(move_email)
    server.tool(annotations=_WRITE)(trash_email)
    server.tool(annotations=_DESTRUCTIVE)(delete_email)
    server.tool(annotations=_WRITE)(bulk_move_emails)
    server.tool(annotations=_WRITE)(bulk_trash_emails)
    server.tool(annotations=_DESTRUCTIVE)(bulk_delete_emails)
