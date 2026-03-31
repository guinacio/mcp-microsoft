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

from typing import Optional, Union

from mcp_microsoft.graph import get_graph
from mcp_microsoft.server import mcp
from mcp_microsoft.tools.mail import _fmt_date, _fmt_recipients, _fmt_sender, _parse_recipients

# ---------------------------------------------------------------------------
# create_draft
# ---------------------------------------------------------------------------


@mcp.tool()
async def create_draft(
    to: Union[str, list[str]],
    subject: str,
    body: str,
    cc: Optional[Union[str, list[str]]] = None,
    body_type: str = "text",
    profile: str | None = None,
) -> str:
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
        Confirmation string with the new draft ID.
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
    to_str = to if isinstance(to, str) else ", ".join(to)
    return (
        f"Draft created successfully.\n"
        f"**Draft ID:** `{draft_id}`\n"
        f"**To:** {to_str}\n"
        f"**Subject:** {subject}"
    )


# ---------------------------------------------------------------------------
# list_drafts
# ---------------------------------------------------------------------------


@mcp.tool()
async def list_drafts(max_results: int = 10, profile: str | None = None) -> str:
    """
    List draft messages from the Drafts folder.

    Args:
        max_results: Maximum number of drafts to return (1-100). Defaults to 10.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted list of draft summaries.
    """
    g = get_graph(profile)
    params: dict = {
        "$top": max_results,
        "$select": "id,subject,toRecipients,lastModifiedDateTime,bodyPreview",
        "$orderby": "lastModifiedDateTime desc",
    }

    result = await g.get("/me/mailFolders/drafts/messages", params=params)
    drafts = result.get("value", [])

    if not drafts:
        return "No drafts found."

    lines = [f"## Drafts ({len(drafts)} messages)\n"]
    for draft in drafts:
        subject = draft.get("subject") or "(no subject)"
        to_str = _fmt_recipients(draft.get("toRecipients", []))
        modified = _fmt_date(draft.get("lastModifiedDateTime"))
        preview = (draft.get("bodyPreview") or "").replace("\n", " ")[:100]
        lines.append(
            f"- **{subject}**\n"
            f"  To: {to_str or '(no recipients)'} | Modified: {modified}\n"
            f"  ID: `{draft.get('id')}`\n"
            f"  > {preview}\n"
        )

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# get_draft
# ---------------------------------------------------------------------------


@mcp.tool()
async def get_draft(draft_id: str, profile: str | None = None) -> str:
    """
    Fetch a draft message by ID.

    Args:
        draft_id: The Graph message ID of the draft.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted draft details including body and recipients.
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
    to_str = _fmt_recipients(draft.get("toRecipients", []))
    cc_str = _fmt_recipients(draft.get("ccRecipients", []))
    bcc_str = _fmt_recipients(draft.get("bccRecipients", []))
    modified = _fmt_date(draft.get("lastModifiedDateTime"))

    body_obj = draft.get("body", {})
    content_type = (body_obj.get("contentType") or "text").lower()
    raw_body = body_obj.get("content", "")

    # Import _strip_html locally to avoid circular import issues at module level
    from mcp_microsoft.tools.mail import _strip_html

    if content_type == "html":
        body_text = _strip_html(raw_body)
    else:
        body_text = raw_body

    lines = [
        f"## Draft: {subject}",
        f"**To:** {to_str or '(none)'}",
    ]
    if cc_str:
        lines.append(f"**CC:** {cc_str}")
    if bcc_str:
        lines.append(f"**BCC:** {bcc_str}")
    lines += [
        f"**Last Modified:** {modified}",
        f"**Draft ID:** `{draft_id}`",
        "",
        "---",
        "",
        body_text,
    ]

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# update_draft
# ---------------------------------------------------------------------------


@mcp.tool()
async def update_draft(
    draft_id: str,
    subject: Optional[str] = None,
    body: Optional[str] = None,
    to: Optional[Union[str, list[str]]] = None,
    cc: Optional[Union[str, list[str]]] = None,
    body_type: str = "text",
    profile: str | None = None,
) -> str:
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
        Confirmation string with updated draft ID.
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
        return "No fields to update — provide at least one of: subject, body, to, cc."

    result = await g.patch(f"/me/messages/{draft_id}", json=patch)

    updated_id = (result or {}).get("id", draft_id)
    updated_fields = ", ".join(patch.keys())
    return (
        f"Draft updated successfully.\n"
        f"**Draft ID:** `{updated_id}`\n"
        f"**Updated fields:** {updated_fields}"
    )


# ---------------------------------------------------------------------------
# send_draft
# ---------------------------------------------------------------------------


@mcp.tool()
async def send_draft(draft_id: str, profile: str | None = None) -> str:
    """
    Send an existing draft message.

    Args:
        draft_id: The Graph message ID of the draft to send.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Confirmation string.
    """
    g = get_graph(profile)
    await g.post(f"/me/messages/{draft_id}/send", json={})
    return (
        f"Draft `{draft_id}` sent successfully.\n"
        "The message has been delivered and saved to Sent Items."
    )
