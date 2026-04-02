"""
Core mail tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.

Implemented:
  - list_emails
  - read_email
  - search_emails
  - send_email
  - reply_email
  - forward_email
  - mark_as_read
  - mark_as_unread
  - move_email
  - trash_email
  - delete_email
"""

from __future__ import annotations

import html as html_module
import re
from datetime import datetime
from typing import Any, Literal, Optional, Union

from mcp.types import ToolAnnotations

from mcp_microsoft.models import (
    Address,
    AttachmentInfo,
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
from mcp_microsoft.graph import get_graph
from mcp_microsoft.server import mcp

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


@mcp.tool(annotations=_READ_ONLY)
async def list_emails(
    folder: str = "inbox",
    max_results: int = 10,
    unread_only: bool = False,
    page_token: Optional[int] = None,
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
        page_token: Integer offset for pagination ($skip). Omit for the first page.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured message summaries with pagination metadata.
    """
    g = get_graph(profile)
    params: dict = {
        "$top": max_results,
        "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments,importance",
        "$orderby": "receivedDateTime desc",
    }
    if unread_only:
        params["$filter"] = "isRead eq false"
    if page_token is not None:
        params["$skip"] = int(page_token)

    result = await g.get(f"/me/mailFolders/{folder}/messages", params=params)

    messages = result.get("value", [])
    next_link = result.get("@odata.nextLink")

    next_page_token: int | None = None
    if next_link:
        skip_match = re.search(r"\$skip=(\d+)", next_link)
        if skip_match:
            next_page_token = int(skip_match.group(1))

    return ListEmailsResponse(
        folder=folder,
        count=len(messages),
        messages=[_message_summary(msg) for msg in messages],
        next_page_token=next_page_token,
        has_more=next_link is not None,
    )


# ---------------------------------------------------------------------------
# read_email
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
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


@mcp.tool(annotations=_READ_ONLY)
async def search_emails(
    query: str,
    max_results: int = 10,
    folder: Optional[str] = None,
    profile: str | None = None,
) -> SearchEmailsResponse:
    """
    Search messages using Graph KQL $search syntax.

    Note: Graph $search and $filter cannot be combined in the same request.

    Args:
        query: KQL search string, e.g. 'from:alice@example.com' or 'project update'.
        max_results: Maximum number of results (1-25 when using $search). Defaults to 10.
        folder: Optional well-known folder name or folder ID to restrict the search.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured search results.
    """
    g = get_graph(profile)
    params: dict = {
        "$search": f'"{query}"',
        "$top": max_results,
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
# send_email
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_WRITE)
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

    Returns:
        Structured send confirmation.
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
    if bcc:
        message["bccRecipients"] = _parse_recipients(bcc)
    if reply_to:
        message["replyTo"] = _parse_recipients(reply_to)

    payload = {
        "message": message,
        "saveToSentItems": save_to_sent,
    }

    await g.post("/me/sendMail", json=payload)

    return SendEmailResponse(
        success=True,
        action="send_email",
        to=[addr.get("emailAddress", {}).get("address", "") for addr in message["toRecipients"]],
        cc=[addr.get("emailAddress", {}).get("address", "") for addr in message.get("ccRecipients", [])],
        bcc=[addr.get("emailAddress", {}).get("address", "") for addr in message.get("bccRecipients", [])],
        reply_to=[addr.get("emailAddress", {}).get("address", "") for addr in message.get("replyTo", [])],
        subject=subject,
        body_type=body_type,
        saved_to_sent_items=save_to_sent,
    )


# ---------------------------------------------------------------------------
# reply_email
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_WRITE)
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


@mcp.tool(annotations=_WRITE)
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


@mcp.tool(annotations=_IDEMPOTENT_WRITE)
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


@mcp.tool(annotations=_IDEMPOTENT_WRITE)
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


@mcp.tool(annotations=_WRITE)
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


@mcp.tool(annotations=_WRITE)
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


@mcp.tool(annotations=_DESTRUCTIVE)
async def delete_email(message_id: str, profile: str | None = None) -> DeleteEmailResponse:
    """
    Permanently delete a message from the mailbox. This action is IRREVERSIBLE.

    The message will be hard-deleted and cannot be recovered from Deleted Items.
    For a recoverable soft delete, use trash_email instead.

    Args:
        message_id: The Graph message ID to permanently delete.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured delete confirmation.
    """
    g = get_graph(profile)
    await g.post(f"/me/messages/{message_id}/permanentDelete")
    return DeleteEmailResponse(
        success=True,
        action="permanent_delete",
        message_id=message_id,
        irreversible=True,
    )
