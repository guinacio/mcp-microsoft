"""
Microsoft Teams tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.
Endpoints primarily live under /me/joinedTeams, /teams, /me/chats,
and /me/onlineMeetings in Graph v1.0.

Required OAuth scopes (add to ProfileConfig.effective_scopes or profiles.json):
    Team.ReadBasic.All
    Channel.ReadBasic.All
    ChannelMessage.Read.All
    ChannelMessage.Send
    Chat.ReadWrite
    OnlineMeetings.ReadWrite

Implemented:
  Teams & Channels:
  - teams_list_joined        — list teams the user has joined
  - teams_get                — get a single team by ID
  - teams_list_channels      — list channels in a team
  - teams_get_channel        — get a single channel by ID
  - teams_create_channel     — create a new standard channel

  Channel Messages:
  - teams_list_channel_messages   — list recent root messages
  - teams_get_channel_message     — get a single channel message
  - teams_send_channel_message    — post a new message to a channel
  - teams_reply_to_channel_message — reply to an existing message thread
  - teams_list_message_replies    — list replies in a thread

  Chats:
  - teams_list_chats         — list 1:1 / group chats
  - teams_get_chat           — get chat details + members
  - teams_list_chat_messages — list messages in a chat
  - teams_send_chat_message  — send a message to a chat
  - teams_create_chat        — create a new 1:1 or group chat

  Online Meetings:
  - teams_create_meeting     — create a Teams online meeting
  - teams_get_meeting        — get meeting details + join URL
  - teams_list_meetings      — list meetings (date-range filtered)

NOTE on throttling: Graph Teams APIs impose stricter rate limits than Mail/Calendar
(e.g. 4 req/s for channel message POSTs). The GraphClient raises httpx.HTTPStatusError
on 429 responses with the Retry-After value included in the message. Callers should
handle this exception and retry after the indicated delay. A future improvement would
be to add automatic retry-with-backoff to GraphClient._request().
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from typing import Any, Literal, Optional

from mcp.types import ToolAnnotations

from mcp_microsoft.graph import get_graph

# ---------------------------------------------------------------------------
# Tool annotation constants
# ---------------------------------------------------------------------------

_READ_ONLY = ToolAnnotations(readOnlyHint=True, openWorldHint=True)
_WRITE = ToolAnnotations(destructiveHint=False, openWorldHint=True)
_DESTRUCTIVE = ToolAnnotations(destructiveHint=True, openWorldHint=True)

# ---------------------------------------------------------------------------
# Type aliases
# ---------------------------------------------------------------------------

ContentType = Literal["text", "html"]

# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------


def _body_payload(content: str, content_type: ContentType = "text") -> dict:
    """Build a Graph message body object."""
    return {
        "body": {
            "contentType": content_type,
            "content": content,
        }
    }


def _build_member(upn_or_id: str, roles: list[str]) -> dict:
    """
    Wrap a UPN or user-ID into the Graph conversationMember format
    required by POST /me/chats.
    """
    return {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": roles,
        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{upn_or_id}')",
    }


def _fmt_dt(iso: Optional[str]) -> str:
    """Format an ISO 8601 datetime to a human-readable form."""
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return iso


def _extract_sender(msg: dict[str, Any]) -> str:
    """Return 'Name <email>' or display name for a Graph chatMessage sender."""
    from_obj = msg.get("from") or {}
    user = from_obj.get("user") or {}
    app = from_obj.get("application") or {}
    device = from_obj.get("device") or {}
    # Prefer user, fall back to application/device
    name = user.get("displayName") or app.get("displayName") or device.get("displayName") or ""
    return name or "(unknown)"


# ---------------------------------------------------------------------------
# Teams & Channels
# ---------------------------------------------------------------------------


async def teams_list_joined(
    top: int = 50,
    profile: str | None = None,
) -> dict:
    """
    List all Microsoft Teams the signed-in user has joined.

    Args:
        top: Maximum number of teams to return (1–100). Defaults to 50.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        dict with 'value' (list of team objects) and optional 'next_link'.
        Each team has id, displayName, description, visibility, webUrl.
    """
    g = get_graph(profile)
    result = await g.get(
        "/me/joinedTeams",
        params={
            "$top": top,
            "$select": "id,displayName,description,visibility,webUrl",
        },
    )
    return {
        "count": len(result.get("value", [])),
        "value": result.get("value", []),
        "next_link": result.get("@odata.nextLink"),
    }


async def teams_get(
    team_id: str,
    profile: str | None = None,
) -> dict:
    """
    Get details for a specific Microsoft Team.

    Args:
        team_id: The team's object ID (GUID).
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Team object with id, displayName, description, visibility,
        isArchived, webUrl, memberSettings, guestSettings, funSettings.
    """
    g = get_graph(profile)
    return await g.get(f"/teams/{team_id}")


async def teams_list_channels(
    team_id: str,
    top: int = 50,
    profile: str | None = None,
) -> dict:
    """
    List all channels in a Microsoft Team.

    Args:
        team_id: The team's object ID.
        top: Maximum number of channels to return (1–100). Defaults to 50.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        dict with 'value' (list of channel objects).
        Each channel has id, displayName, description, channelType, webUrl.
    """
    g = get_graph(profile)
    result = await g.get(
        f"/teams/{team_id}/channels",
        params={
            "$top": top,
            "$select": "id,displayName,description,channelType,webUrl,isFavoriteByDefault",
        },
    )
    return {
        "team_id": team_id,
        "count": len(result.get("value", [])),
        "value": result.get("value", []),
        "next_link": result.get("@odata.nextLink"),
    }


async def teams_get_channel(
    team_id: str,
    channel_id: str,
    profile: str | None = None,
) -> dict:
    """
    Get details for a specific channel in a Microsoft Team.

    Args:
        team_id: The team's object ID.
        channel_id: The channel's object ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Channel object with id, displayName, description, channelType, webUrl.
    """
    g = get_graph(profile)
    return await g.get(f"/teams/{team_id}/channels/{channel_id}")


async def teams_create_channel(
    team_id: str,
    display_name: str,
    description: str = "",
    confirm: bool = False,
    profile: str | None = None,
) -> dict:
    """
    Create a new standard channel in a Microsoft Team.

    This is a write operation that modifies the team structure. Set confirm=True
    to execute. If confirm is False the tool returns a dry-run preview without
    making any API call.

    Args:
        team_id: The team's object ID.
        display_name: Display name for the new channel (must be unique in the team).
        description: Optional channel description.
        confirm: Must be True to actually create the channel. Defaults to False (dry run).
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Created channel object (id, displayName, description, webUrl) on success,
        or a dry-run preview dict when confirm=False.
    """
    if not confirm:
        return {
            "dry_run": True,
            "message": "Set confirm=True to create the channel.",
            "team_id": team_id,
            "display_name": display_name,
            "description": description,
        }

    g = get_graph(profile)
    payload: dict = {
        "displayName": display_name,
        "channelType": "standard",
    }
    if description:
        payload["description"] = description

    result = await g.post(f"/teams/{team_id}/channels", json=payload)
    return result or {}


# ---------------------------------------------------------------------------
# Channel Messages
# ---------------------------------------------------------------------------


async def teams_list_channel_messages(
    team_id: str,
    channel_id: str,
    top: int = 20,
    profile: str | None = None,
) -> dict:
    """
    List recent root (non-reply) messages in a Teams channel.

    Args:
        team_id: The team's object ID.
        channel_id: The channel's object ID.
        top: Maximum number of messages to return (1–50). Defaults to 20.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        dict with 'value' (list of message objects) and optional 'next_link'.
        Messages include id, createdDateTime, from, body (preview), subject, webUrl.
    """
    g = get_graph(profile)
    result = await g.get(
        f"/teams/{team_id}/channels/{channel_id}/messages",
        params={
            "$top": top,
            "$select": "id,createdDateTime,lastModifiedDateTime,from,body,subject,webUrl,replyToId,importance",
        },
    )
    messages = result.get("value", [])
    # Trim body content to avoid bloating LLM context
    for msg in messages:
        body = msg.get("body", {})
        content = body.get("content", "")
        if len(content) > 500:
            body["content"] = content[:500] + "…"
    return {
        "team_id": team_id,
        "channel_id": channel_id,
        "count": len(messages),
        "value": messages,
        "next_link": result.get("@odata.nextLink"),
    }


async def teams_get_channel_message(
    team_id: str,
    channel_id: str,
    message_id: str,
    profile: str | None = None,
) -> dict:
    """
    Get a single channel message by ID (full content).

    Args:
        team_id: The team's object ID.
        channel_id: The channel's object ID.
        message_id: The message's object ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Full message object with id, createdDateTime, from, body, subject,
        webUrl, reactions, attachments, mentions.
    """
    g = get_graph(profile)
    return await g.get(f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}")


async def teams_send_channel_message(
    team_id: str,
    channel_id: str,
    content: str,
    content_type: ContentType = "text",
    subject: Optional[str] = None,
    profile: str | None = None,
) -> dict:
    """
    Send a new message to a Teams channel.

    Args:
        team_id: The team's object ID.
        channel_id: The channel's object ID.
        content: Message body text or HTML.
        content_type: 'text' or 'html'. Defaults to 'text'.
        subject: Optional message subject/headline (shown in bold above the body).
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        dict with success flag and the created message id, webUrl, createdDateTime.
    """
    g = get_graph(profile)
    payload = _body_payload(content, content_type)
    if subject:
        payload["subject"] = subject

    result = await g.post(
        f"/teams/{team_id}/channels/{channel_id}/messages",
        json=payload,
    )
    result = result or {}
    return {
        "success": True,
        "id": result.get("id", ""),
        "web_url": result.get("webUrl", ""),
        "created_at": result.get("createdDateTime", ""),
        "created_at_display": _fmt_dt(result.get("createdDateTime")),
    }


async def teams_reply_to_channel_message(
    team_id: str,
    channel_id: str,
    message_id: str,
    content: str,
    content_type: ContentType = "text",
    profile: str | None = None,
) -> dict:
    """
    Reply to an existing root message in a Teams channel thread.

    Args:
        team_id: The team's object ID.
        channel_id: The channel's object ID.
        message_id: The root message ID to reply to.
        content: Reply body text or HTML.
        content_type: 'text' or 'html'. Defaults to 'text'.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        dict with success flag and the created reply id, webUrl, createdDateTime.
    """
    g = get_graph(profile)
    payload = _body_payload(content, content_type)
    result = await g.post(
        f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies",
        json=payload,
    )
    result = result or {}
    return {
        "success": True,
        "id": result.get("id", ""),
        "parent_message_id": message_id,
        "web_url": result.get("webUrl", ""),
        "created_at": result.get("createdDateTime", ""),
        "created_at_display": _fmt_dt(result.get("createdDateTime")),
    }


async def teams_list_message_replies(
    team_id: str,
    channel_id: str,
    message_id: str,
    top: int = 20,
    profile: str | None = None,
) -> dict:
    """
    List all replies to a root message in a Teams channel thread.

    Args:
        team_id: The team's object ID.
        channel_id: The channel's object ID.
        message_id: The root message ID whose replies to fetch.
        top: Maximum number of replies to return (1–50). Defaults to 20.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        dict with 'value' (list of reply message objects) and optional 'next_link'.
    """
    g = get_graph(profile)
    result = await g.get(
        f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies",
        params={
            "$top": top,
            "$select": "id,createdDateTime,from,body,webUrl,importance",
        },
    )
    replies = result.get("value", [])
    for msg in replies:
        body = msg.get("body", {})
        content = body.get("content", "")
        if len(content) > 500:
            body["content"] = content[:500] + "…"
    return {
        "team_id": team_id,
        "channel_id": channel_id,
        "parent_message_id": message_id,
        "count": len(replies),
        "value": replies,
        "next_link": result.get("@odata.nextLink"),
    }


# ---------------------------------------------------------------------------
# Chats
# ---------------------------------------------------------------------------


async def teams_list_chats(
    top: int = 20,
    chat_type: str = "",
    profile: str | None = None,
) -> dict:
    """
    List all chats the signed-in user is part of.

    Args:
        top: Maximum number of chats to return (1–50). Defaults to 20.
        chat_type: Optional filter. One of: 'oneOnOne', 'group', 'meeting'.
                   Leave empty to return all chat types.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        dict with 'value' (list of chat objects) and optional 'next_link'.
        Each chat has id, chatType, topic, createdDateTime, lastUpdatedDateTime, webUrl.
    """
    g = get_graph(profile)
    params: dict = {
        "$top": top,
        "$select": "id,chatType,topic,createdDateTime,lastUpdatedDateTime,webUrl",
        "$orderby": "lastUpdatedDateTime desc",
    }
    if chat_type:
        params["$filter"] = f"chatType eq '{chat_type}'"

    result = await g.get("/me/chats", params=params)
    return {
        "count": len(result.get("value", [])),
        "value": result.get("value", []),
        "next_link": result.get("@odata.nextLink"),
    }


async def teams_get_chat(
    chat_id: str,
    profile: str | None = None,
) -> dict:
    """
    Get details and members for a specific chat.

    Args:
        chat_id: The chat's object ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Chat object with id, chatType, topic, createdDateTime, webUrl,
        and expanded 'members' list (each member has displayName, email, roles).
    """
    g = get_graph(profile)
    return await g.get(f"/me/chats/{chat_id}", params={"$expand": "members"})


async def teams_list_chat_messages(
    chat_id: str,
    top: int = 20,
    profile: str | None = None,
) -> dict:
    """
    List recent messages in a Teams chat (1:1 or group).

    Args:
        chat_id: The chat's object ID.
        top: Maximum number of messages to return (1–50). Defaults to 20.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        dict with 'value' (list of message objects) and optional 'next_link'.
        Messages include id, createdDateTime, from (displayName), body (preview), webUrl.
    """
    g = get_graph(profile)
    result = await g.get(
        f"/me/chats/{chat_id}/messages",
        params={
            "$top": top,
            "$select": "id,createdDateTime,lastModifiedDateTime,from,body,webUrl,importance",
        },
    )
    messages = result.get("value", [])
    for msg in messages:
        body = msg.get("body", {})
        content = body.get("content", "")
        if len(content) > 500:
            body["content"] = content[:500] + "…"
    return {
        "chat_id": chat_id,
        "count": len(messages),
        "value": messages,
        "next_link": result.get("@odata.nextLink"),
    }


async def teams_send_chat_message(
    chat_id: str,
    content: str,
    content_type: ContentType = "text",
    profile: str | None = None,
) -> dict:
    """
    Send a message to a Teams chat (1:1 or group).

    Args:
        chat_id: The chat's object ID.
        content: Message body text or HTML.
        content_type: 'text' or 'html'. Defaults to 'text'.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        dict with success flag and the created message id, webUrl, createdDateTime.
    """
    g = get_graph(profile)
    payload = _body_payload(content, content_type)
    result = await g.post(f"/me/chats/{chat_id}/messages", json=payload)
    result = result or {}
    return {
        "success": True,
        "chat_id": chat_id,
        "id": result.get("id", ""),
        "web_url": result.get("webUrl", ""),
        "created_at": result.get("createdDateTime", ""),
        "created_at_display": _fmt_dt(result.get("createdDateTime")),
    }


async def teams_create_chat(
    members: list[str],
    topic: str = "",
    chat_type: str = "group",
    profile: str | None = None,
) -> dict:
    """
    Create a new Teams chat (1:1 or group).

    The signed-in user is automatically added as an owner. For a 1:1 chat,
    provide exactly one other member UPN; for a group chat provide two or more.

    Args:
        members: List of UPNs (email addresses) or user object IDs to include.
                 Do NOT include the signed-in user — they are added automatically.
        topic: Optional group chat topic/name. Ignored for 1:1 chats.
        chat_type: 'oneOnOne' or 'group'. Defaults to 'group'.
                   Use 'oneOnOne' only when members has exactly 1 entry.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Created chat object with id, chatType, topic, webUrl.
    """
    g = get_graph(profile)

    # The signed-in user must be included as owner; the others as members.
    # We don't know the user's ID here, so we rely on Graph to inject the
    # caller automatically when using /me/chats — only the other members
    # need to be listed. However, Graph API requires the initiator to be in
    # the members list with role "owner". We fetch the current user's info.
    me = await g.get("/me", params={"$select": "id"})
    my_id: str = me.get("id", "")

    member_objs: list[dict] = []
    if my_id:
        member_objs.append(_build_member(my_id, ["owner"]))

    for upn in members:
        member_objs.append(_build_member(upn.strip(), []))

    payload: dict = {
        "chatType": chat_type,
        "members": member_objs,
    }
    if topic and chat_type != "oneOnOne":
        payload["topic"] = topic

    result = await g.post("/chats", json=payload)
    result = result or {}
    return {
        "success": True,
        "id": result.get("id", ""),
        "chat_type": result.get("chatType", chat_type),
        "topic": result.get("topic", topic),
        "web_url": result.get("webUrl", ""),
        "created_at": result.get("createdDateTime", ""),
        "created_at_display": _fmt_dt(result.get("createdDateTime")),
    }


# ---------------------------------------------------------------------------
# Online Meetings
# ---------------------------------------------------------------------------


async def teams_create_meeting(
    subject: str,
    start_datetime: str,
    end_datetime: str,
    attendees: list[str] | None = None,
    profile: str | None = None,
) -> dict:
    """
    Create a new Teams online meeting.

    A join URL is generated regardless of whether attendees are specified.
    The meeting is created on behalf of the signed-in user.

    Args:
        subject: Meeting subject / title.
        start_datetime: ISO 8601 start time (e.g. '2026-04-10T14:00:00Z').
        end_datetime: ISO 8601 end time (e.g. '2026-04-10T15:00:00Z').
        attendees: Optional list of attendee UPNs (email addresses).
                   If omitted, the meeting has no invited attendees but is
                   still accessible via the join URL.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        dict with success flag, meeting id, joinWebUrl, subject, start, end.
    """
    g = get_graph(profile)
    payload: dict = {
        "subject": subject,
        "startDateTime": start_datetime,
        "endDateTime": end_datetime,
    }

    if attendees:
        payload["participants"] = {
            "attendees": [
                {"upn": upn.strip(), "role": "attendee"}
                for upn in attendees
                if upn.strip()
            ]
        }

    result = await g.post("/me/onlineMeetings", json=payload)
    result = result or {}
    return {
        "success": True,
        "id": result.get("id", ""),
        "subject": result.get("subject", subject),
        "join_web_url": result.get("joinWebUrl") or "",
        "join_meeting_id": (result.get("joinMeetingIdSettings") or {}).get("joinMeetingId", ""),
        "start": result.get("startDateTime", start_datetime),
        "end": result.get("endDateTime", end_datetime),
        "start_display": _fmt_dt(result.get("startDateTime", start_datetime)),
        "end_display": _fmt_dt(result.get("endDateTime", end_datetime)),
        "created_at": result.get("createdDateTime", ""),
    }


async def teams_get_meeting(
    meeting_id: str,
    profile: str | None = None,
) -> dict:
    """
    Get details and join URL for a Teams online meeting.

    Args:
        meeting_id: The online meeting object ID (from teams_create_meeting or
                    teams_list_meetings).
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Meeting object with id, subject, joinWebUrl, startDateTime, endDateTime,
        participants, and videoTeleconferenceId.
    """
    g = get_graph(profile)
    return await g.get(f"/me/onlineMeetings/{meeting_id}")


async def teams_list_meetings(
    start_after: Optional[str] = None,
    start_before: Optional[str] = None,
    top: int = 10,
    profile: str | None = None,
) -> dict:
    """
    List Teams online meetings within a date range.

    IMPORTANT: GET /me/onlineMeetings without a filter returns a Graph API error.
    This tool always applies a startDateTime range filter. The default window is
    today ± 7 days when no explicit dates are provided.

    Args:
        start_after: ISO 8601 datetime — only meetings starting at or after this
                     time are returned. Defaults to 7 days ago (UTC).
        start_before: ISO 8601 datetime — only meetings starting before this time
                      are returned. Defaults to 7 days from now (UTC).
        top: Maximum number of meetings to return (1–50). Defaults to 10.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        dict with 'value' (list of meeting objects) and optional 'next_link'.
        Each meeting has id, subject, startDateTime, endDateTime, joinWebUrl.
    """
    now = datetime.now(tz=timezone.utc)

    if start_after is None:
        start_after = (now - timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%SZ")
    if start_before is None:
        start_before = (now + timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%SZ")

    filter_str = (
        f"startDateTime ge '{start_after}' and startDateTime le '{start_before}'"
    )

    g = get_graph(profile)
    result = await g.get(
        "/me/onlineMeetings",
        params={
            "$filter": filter_str,
            "$top": top,
            "$select": "id,subject,startDateTime,endDateTime,joinWebUrl,createdDateTime",
        },
    )
    meetings = result.get("value", [])
    # Augment with display-friendly timestamps
    for m in meetings:
        m["start_display"] = _fmt_dt(m.get("startDateTime"))
        m["end_display"] = _fmt_dt(m.get("endDateTime"))

    return {
        "filter": {"start_after": start_after, "start_before": start_before},
        "count": len(meetings),
        "value": meetings,
        "next_link": result.get("@odata.nextLink"),
    }


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------


def register(server) -> None:
    """Register all Teams tools with the given FastMCP server instance."""
    # Teams & Channels
    server.tool(annotations=_READ_ONLY)(teams_list_joined)
    server.tool(annotations=_READ_ONLY)(teams_get)
    server.tool(annotations=_READ_ONLY)(teams_list_channels)
    server.tool(annotations=_READ_ONLY)(teams_get_channel)
    server.tool(annotations=_WRITE)(teams_create_channel)
    # Channel Messages
    server.tool(annotations=_READ_ONLY)(teams_list_channel_messages)
    server.tool(annotations=_READ_ONLY)(teams_get_channel_message)
    server.tool(annotations=_WRITE)(teams_send_channel_message)
    server.tool(annotations=_WRITE)(teams_reply_to_channel_message)
    server.tool(annotations=_READ_ONLY)(teams_list_message_replies)
    # Chats
    server.tool(annotations=_READ_ONLY)(teams_list_chats)
    server.tool(annotations=_READ_ONLY)(teams_get_chat)
    server.tool(annotations=_READ_ONLY)(teams_list_chat_messages)
    server.tool(annotations=_WRITE)(teams_send_chat_message)
    server.tool(annotations=_WRITE)(teams_create_chat)
    # Online Meetings
    server.tool(annotations=_WRITE)(teams_create_meeting)
    server.tool(annotations=_READ_ONLY)(teams_get_meeting)
    server.tool(annotations=_READ_ONLY)(teams_list_meetings)
