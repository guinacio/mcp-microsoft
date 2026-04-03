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

from mcp_microsoft.common.formatting import format_datetime_display
from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.tooling import READ_ONLY_TOOL, WRITE_TOOL, register_tool
from mcp_microsoft.graph import get_graph
from mcp_microsoft.graph_types import (
    GraphChannel,
    GraphChat,
    GraphChatMember,
    GraphChatMessage,
    GraphOnlineMeetingDetail,
    GraphTeam,
    parse_graph_collection,
)
from mcp_microsoft.models import (
    ChannelDetailResponse,
    ChannelInfo,
    ChannelMessageDetailResponse,
    ChatDetailResponse,
    ChatInfo,
    ChatMemberInfo,
    CreateChannelResponse,
    CreateTeamsChatResponse,
    CreateTeamsMeetingResponse,
    MeetingDetailResponse,
    OnlineMeetingInfo,
    SendTeamsMessageResponse,
    TeamDetailResponse,
    TeamFunSettingsInfo,
    TeamGuestSettingsInfo,
    TeamMemberSettingsInfo,
    TeamInfo,
    TeamsMeetingParticipantInfo,
    TeamsMeetingParticipantsInfo,
    TeamsMentionInfo,
    TeamsMessageAttachmentInfo,
    TeamsReactionInfo,
    TeamsListChannelMessagesResponse,
    TeamsListChannelsResponse,
    TeamsListChatsResponse,
    TeamsListChatMessagesResponse,
    TeamsListJoinedResponse,
    TeamsListMeetingsResponse,
    TeamsListRepliesResponse,
    TeamsMessageInfo,
)

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


def _normalize_filter_datetime(value: str | None, fallback: datetime) -> str:
    """Normalize a caller-supplied datetime into a safe ISO 8601 UTC filter value."""
    if not value:
        return fallback.strftime("%Y-%m-%dT%H:%M:%SZ")
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        raise ValueError(
            "Datetime filters must be valid ISO 8601 values like 2026-04-10T14:00:00Z."
        ) from exc
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    else:
        parsed = parsed.astimezone(timezone.utc)
    return parsed.strftime("%Y-%m-%dT%H:%M:%SZ")


def _extract_sender(from_obj: Any | None) -> str:
    """Return sender display name from Graph identity-like payloads."""
    user = getattr(from_obj, "user", None) if from_obj else None
    app = getattr(from_obj, "application", None) if from_obj else None
    device = getattr(from_obj, "device", None) if from_obj else None
    conversation = getattr(from_obj, "conversation", None) if from_obj else None
    tag = getattr(from_obj, "tag", None) if from_obj else None
    name = (
        (user.display_name if user else "")
        or (app.display_name if app else "")
        or (device.display_name if device else "")
        or (conversation.display_name if conversation else "")
        or (tag.display_name if tag else "")
    )
    return name or "(unknown)"


def _team_info(team: GraphTeam) -> TeamInfo:
    return TeamInfo(
        id=team.id,
        display_name=team.display_name,
        description=team.description,
        visibility=team.visibility,
        web_url=team.web_url,
        is_archived=team.is_archived,
    )


def _channel_info(channel: GraphChannel) -> ChannelInfo:
    return ChannelInfo(
        id=channel.id,
        display_name=channel.display_name,
        description=channel.description,
        channel_type=channel.channel_type,
        web_url=channel.web_url,
        is_favorite_by_default=channel.is_favorite_by_default,
    )


def _message_info(msg: GraphChatMessage) -> TeamsMessageInfo:
    content = msg.body.content
    if len(content) > 500:
        content = content[:500] + "…"
    return TeamsMessageInfo(
        id=msg.id,
        created_at=msg.created_date_time,
        created_at_display=format_datetime_display(msg.created_date_time),
        last_modified_at=msg.last_modified_date_time,
        from_display=_extract_sender(msg.from_),
        body=content,
        body_content_type=msg.body.content_type,
        subject=msg.subject,
        web_url=msg.web_url,
        reply_to_id=msg.reply_to_id,
        importance=msg.importance,
    )


def _chat_info(chat: GraphChat) -> ChatInfo:
    return ChatInfo(
        id=chat.id,
        chat_type=chat.chat_type,
        topic=chat.topic,
        created_at=chat.created_date_time,
        created_at_display=format_datetime_display(chat.created_date_time),
        last_updated_at=chat.last_updated_date_time,
        last_updated_at_display=format_datetime_display(chat.last_updated_date_time),
        web_url=chat.web_url,
    )


def _chat_member_info(member: GraphChatMember) -> ChatMemberInfo:
    email = member.email or member.user_id or member.tenant_id
    return ChatMemberInfo(
        id=member.id,
        display_name=member.display_name,
        email=email,
        roles=member.roles,
    )


def _meeting_info(meeting: GraphOnlineMeetingDetail) -> OnlineMeetingInfo:
    return OnlineMeetingInfo(
        id=meeting.id,
        subject=meeting.subject,
        join_web_url=meeting.join_web_url,
        join_meeting_id=meeting.join_meeting_id_settings.join_meeting_id if meeting.join_meeting_id_settings else "",
        start=meeting.start_date_time,
        end=meeting.end_date_time,
        start_display=format_datetime_display(meeting.start_date_time),
        end_display=format_datetime_display(meeting.end_date_time),
        created_at=meeting.created_date_time,
    )


def _team_member_settings_info(settings) -> TeamMemberSettingsInfo:
    return TeamMemberSettingsInfo.model_validate(settings.model_dump(by_alias=False))


def _team_guest_settings_info(settings) -> TeamGuestSettingsInfo:
    return TeamGuestSettingsInfo.model_validate(settings.model_dump(by_alias=False))


def _team_fun_settings_info(settings) -> TeamFunSettingsInfo:
    return TeamFunSettingsInfo.model_validate(settings.model_dump(by_alias=False))


def _reaction_info(reaction) -> TeamsReactionInfo:
    user_display = _extract_sender(reaction.user)
    return TeamsReactionInfo.model_validate(
        {
            **reaction.model_dump(by_alias=False),
            "created_at": reaction.created_date_time,
            "user_display": user_display,
        }
    )


def _message_attachment_info(attachment) -> TeamsMessageAttachmentInfo:
    return TeamsMessageAttachmentInfo.model_validate(attachment.model_dump(by_alias=False))


def _mention_info(mention) -> TeamsMentionInfo:
    mentioned = mention.mentioned
    mentioned_display = _extract_sender(mentioned) if mentioned else ""
    return TeamsMentionInfo.model_validate(
        {
            **mention.model_dump(by_alias=False),
            "mentioned_display": mentioned_display,
        }
    )


def _meeting_participant_info(participant) -> TeamsMeetingParticipantInfo:
    identity_display = _extract_sender(participant.identity)
    return TeamsMeetingParticipantInfo.model_validate(
        {
            **participant.model_dump(by_alias=False),
            "identity_display": identity_display,
        }
    )


def _meeting_participants_info(participants) -> TeamsMeetingParticipantsInfo:
    return TeamsMeetingParticipantsInfo(
        organizer=_meeting_participant_info(participants.organizer) if participants.organizer else None,
        attendees=[_meeting_participant_info(item) for item in participants.attendees],
        producers=[_meeting_participant_info(item) for item in participants.producers],
        contributors=[_meeting_participant_info(item) for item in participants.contributors],
    )


class TeamsListJoinedInput(ToolRequestModel):
    top: int = 50
    profile: str | None = None


class TeamsGetInput(ToolRequestModel):
    team_id: str
    profile: str | None = None


class TeamsListChannelsInput(ToolRequestModel):
    team_id: str
    top: int = 50
    profile: str | None = None


class TeamsGetChannelInput(ToolRequestModel):
    team_id: str
    channel_id: str
    profile: str | None = None


class TeamsCreateChannelInput(ToolRequestModel):
    team_id: str
    display_name: str
    description: str = ""
    confirm: bool = False
    profile: str | None = None


class TeamsListChannelMessagesInput(ToolRequestModel):
    team_id: str
    channel_id: str
    top: int = 20
    profile: str | None = None


class TeamsGetChannelMessageInput(ToolRequestModel):
    team_id: str
    channel_id: str
    message_id: str
    profile: str | None = None


class TeamsSendChannelMessageInput(ToolRequestModel):
    team_id: str
    channel_id: str
    content: str
    content_type: ContentType = "text"
    subject: str | None = None
    profile: str | None = None


class TeamsReplyToChannelMessageInput(ToolRequestModel):
    team_id: str
    channel_id: str
    message_id: str
    content: str
    content_type: ContentType = "text"
    profile: str | None = None


class TeamsListMessageRepliesInput(ToolRequestModel):
    team_id: str
    channel_id: str
    message_id: str
    top: int = 20
    profile: str | None = None


class TeamsListChatsInput(ToolRequestModel):
    top: int = 20
    chat_type: Literal["", "oneOnOne", "group", "meeting"] = ""
    profile: str | None = None


class TeamsGetChatInput(ToolRequestModel):
    chat_id: str
    profile: str | None = None


class TeamsListChatMessagesInput(ToolRequestModel):
    chat_id: str
    top: int = 20
    profile: str | None = None


class TeamsSendChatMessageInput(ToolRequestModel):
    chat_id: str
    content: str
    content_type: ContentType = "text"
    profile: str | None = None


class TeamsCreateChatInput(ToolRequestModel):
    members: list[str]
    topic: str = ""
    chat_type: Literal["oneOnOne", "group"] = "group"
    profile: str | None = None


class TeamsCreateMeetingInput(ToolRequestModel):
    subject: str
    start_datetime: str
    end_datetime: str
    attendees: list[str] | None = None
    profile: str | None = None


class TeamsGetMeetingInput(ToolRequestModel):
    meeting_id: str
    profile: str | None = None


class TeamsListMeetingsInput(ToolRequestModel):
    start_after: str | None = None
    start_before: str | None = None
    top: int = 10
    profile: str | None = None


# ---------------------------------------------------------------------------
# Teams & Channels
# ---------------------------------------------------------------------------


async def teams_list_joined(
    params: TeamsListJoinedInput,
) -> TeamsListJoinedResponse:
    """
    List all Microsoft Teams the signed-in user has joined.

    Args:
        top: Maximum number of teams to return (1–100). Defaults to 50.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        dict with 'value' (list of team objects) and optional 'next_link'.
        Each team has id, displayName, description, visibility, webUrl.
    """
    g = get_graph(params.profile)
    result = await g.get(
        "/me/joinedTeams",
        params={
            "$top": params.top,
            "$select": "id,displayName,description,visibility,webUrl",
        },
    )
    teams = parse_graph_collection(result, GraphTeam)
    return TeamsListJoinedResponse(
        count=len(teams),
        teams=[_team_info(team) for team in teams],
        next_link=result.get("@odata.nextLink"),
    )


async def teams_get(
    params: TeamsGetInput,
) -> TeamDetailResponse:
    """
    Get details for a specific Microsoft Team.

    Args:
        team_id: The team's object ID (GUID).
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Team object with id, displayName, description, visibility,
        isArchived, webUrl, memberSettings, guestSettings, funSettings.
    """
    g = get_graph(params.profile)
    team = GraphTeam.model_validate(await g.get(f"/teams/{params.team_id}"))
    return TeamDetailResponse(
        **_team_info(team).model_dump(),
        member_settings=_team_member_settings_info(team.member_settings),
        guest_settings=_team_guest_settings_info(team.guest_settings),
        fun_settings=_team_fun_settings_info(team.fun_settings),
    )


async def teams_list_channels(
    params: TeamsListChannelsInput,
) -> TeamsListChannelsResponse:
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
    g = get_graph(params.profile)
    result = await g.get(
        f"/teams/{params.team_id}/channels",
        params={
            "$top": params.top,
            "$select": "id,displayName,description,channelType,webUrl,isFavoriteByDefault",
        },
    )
    channels = parse_graph_collection(result, GraphChannel)
    return TeamsListChannelsResponse(
        team_id=params.team_id,
        count=len(channels),
        channels=[_channel_info(channel) for channel in channels],
        next_link=result.get("@odata.nextLink"),
    )


async def teams_get_channel(
    params: TeamsGetChannelInput,
) -> ChannelDetailResponse:
    """
    Get details for a specific channel in a Microsoft Team.

    Args:
        team_id: The team's object ID.
        channel_id: The channel's object ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Channel object with id, displayName, description, channelType, webUrl.
    """
    g = get_graph(params.profile)
    channel = GraphChannel.model_validate(
        await g.get(f"/teams/{params.team_id}/channels/{params.channel_id}")
    )
    return ChannelDetailResponse(**_channel_info(channel).model_dump())


async def teams_create_channel(
    params: TeamsCreateChannelInput,
) -> CreateChannelResponse:
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
    if not params.confirm:
        return CreateChannelResponse(
            success=True,
            action="create_channel",
            team_id=params.team_id,
            display_name=params.display_name,
            description=params.description,
            dry_run=True,
            requires_confirmation=True,
        )

    g = get_graph(params.profile)
    payload: dict = {
        "displayName": params.display_name,
        "channelType": "standard",
    }
    if params.description:
        payload["description"] = params.description

    channel = GraphChannel.model_validate(
        await g.post(f"/teams/{params.team_id}/channels", json=payload) or {}
    )
    return CreateChannelResponse(
        success=True,
        action="create_channel",
        team_id=params.team_id,
        channel_id=channel.id,
        display_name=channel.display_name or params.display_name,
        description=channel.description or params.description,
        web_url=channel.web_url,
    )


# ---------------------------------------------------------------------------
# Channel Messages
# ---------------------------------------------------------------------------


async def teams_list_channel_messages(
    params: TeamsListChannelMessagesInput,
) -> TeamsListChannelMessagesResponse:
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
    g = get_graph(params.profile)
    result = await g.get(
        f"/teams/{params.team_id}/channels/{params.channel_id}/messages",
        params={
            "$top": params.top,
            "$select": "id,createdDateTime,lastModifiedDateTime,from,body,subject,webUrl,replyToId,importance",
        },
    )
    messages = parse_graph_collection(result, GraphChatMessage)
    return TeamsListChannelMessagesResponse(
        team_id=params.team_id,
        channel_id=params.channel_id,
        count=len(messages),
        messages=[_message_info(msg) for msg in messages],
        next_link=result.get("@odata.nextLink"),
    )


async def teams_get_channel_message(
    params: TeamsGetChannelMessageInput,
) -> ChannelMessageDetailResponse:
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
    g = get_graph(params.profile)
    message = GraphChatMessage.model_validate(await g.get(
        f"/teams/{params.team_id}/channels/{params.channel_id}/messages/{params.message_id}"
    ))
    return ChannelMessageDetailResponse(
        **_message_info(message).model_dump(),
        reactions=[_reaction_info(item) for item in message.reactions],
        attachments=[_message_attachment_info(item) for item in message.attachments],
        mentions=[_mention_info(item) for item in message.mentions],
    )


async def teams_send_channel_message(
    params: TeamsSendChannelMessageInput,
) -> SendTeamsMessageResponse:
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
    g = get_graph(params.profile)
    payload = _body_payload(params.content, params.content_type)
    if params.subject:
        payload["subject"] = params.subject

    message = GraphChatMessage.model_validate(await g.post(
        f"/teams/{params.team_id}/channels/{params.channel_id}/messages",
        json=payload,
    ) or {})
    return SendTeamsMessageResponse(
        success=True,
        action="teams_send_channel_message",
        id=message.id,
        web_url=message.web_url,
        created_at=message.created_date_time,
        created_at_display=format_datetime_display(message.created_date_time),
    )


async def teams_reply_to_channel_message(
    params: TeamsReplyToChannelMessageInput,
) -> SendTeamsMessageResponse:
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
    g = get_graph(params.profile)
    payload = _body_payload(params.content, params.content_type)
    message = GraphChatMessage.model_validate(await g.post(
        f"/teams/{params.team_id}/channels/{params.channel_id}/messages/{params.message_id}/replies",
        json=payload,
    ) or {})
    return SendTeamsMessageResponse(
        success=True,
        action="teams_reply_to_channel_message",
        id=message.id,
        parent_message_id=params.message_id,
        web_url=message.web_url,
        created_at=message.created_date_time,
        created_at_display=format_datetime_display(message.created_date_time),
    )


async def teams_list_message_replies(
    params: TeamsListMessageRepliesInput,
) -> TeamsListRepliesResponse:
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
    g = get_graph(params.profile)
    result = await g.get(
        f"/teams/{params.team_id}/channels/{params.channel_id}/messages/{params.message_id}/replies",
        params={
            "$top": params.top,
            "$select": "id,createdDateTime,from,body,webUrl,importance",
        },
    )
    replies = parse_graph_collection(result, GraphChatMessage)
    return TeamsListRepliesResponse(
        team_id=params.team_id,
        channel_id=params.channel_id,
        parent_message_id=params.message_id,
        count=len(replies),
        replies=[_message_info(reply) for reply in replies],
        next_link=result.get("@odata.nextLink"),
    )


# ---------------------------------------------------------------------------
# Chats
# ---------------------------------------------------------------------------


async def teams_list_chats(
    params: TeamsListChatsInput,
) -> TeamsListChatsResponse:
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
    g = get_graph(params.profile)
    query: dict[str, Any] = {
        "$top": params.top,
        "$select": "id,chatType,topic,createdDateTime,lastUpdatedDateTime,webUrl",
        "$orderby": "lastUpdatedDateTime desc",
    }
    if params.chat_type:
        query["$filter"] = f"chatType eq '{params.chat_type}'"

    result = await g.get("/me/chats", params=query)
    chats = parse_graph_collection(result, GraphChat)
    return TeamsListChatsResponse(
        count=len(chats),
        chats=[_chat_info(chat) for chat in chats],
        next_link=result.get("@odata.nextLink"),
    )


async def teams_get_chat(
    params: TeamsGetChatInput,
) -> ChatDetailResponse:
    """
    Get details and members for a specific chat.

    Args:
        chat_id: The chat's object ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Chat object with id, chatType, topic, createdDateTime, webUrl,
        and expanded 'members' list (each member has displayName, email, roles).
    """
    g = get_graph(params.profile)
    result = await g.get(f"/me/chats/{params.chat_id}", params={"$expand": "members"})
    chat = GraphChat.model_validate(result)
    return ChatDetailResponse(
        **_chat_info(chat).model_dump(),
        members=[_chat_member_info(GraphChatMember.model_validate(member)) for member in result.get("members") or []],
    )


async def teams_list_chat_messages(
    params: TeamsListChatMessagesInput,
) -> TeamsListChatMessagesResponse:
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
    g = get_graph(params.profile)
    result = await g.get(
        f"/me/chats/{params.chat_id}/messages",
        params={
            "$top": params.top,
            "$select": "id,createdDateTime,lastModifiedDateTime,from,body,webUrl,importance",
        },
    )
    messages = parse_graph_collection(result, GraphChatMessage)
    return TeamsListChatMessagesResponse(
        chat_id=params.chat_id,
        count=len(messages),
        messages=[_message_info(msg) for msg in messages],
        next_link=result.get("@odata.nextLink"),
    )


async def teams_send_chat_message(
    params: TeamsSendChatMessageInput,
) -> SendTeamsMessageResponse:
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
    g = get_graph(params.profile)
    payload = _body_payload(params.content, params.content_type)
    message = GraphChatMessage.model_validate(
        await g.post(f"/me/chats/{params.chat_id}/messages", json=payload) or {}
    )
    return SendTeamsMessageResponse(
        success=True,
        action="teams_send_chat_message",
        chat_id=params.chat_id,
        id=message.id,
        web_url=message.web_url,
        created_at=message.created_date_time,
        created_at_display=format_datetime_display(message.created_date_time),
    )


async def teams_create_chat(
    params: TeamsCreateChatInput,
) -> CreateTeamsChatResponse:
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
    g = get_graph(params.profile)

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

    for upn in params.members:
        member_objs.append(_build_member(upn.strip(), []))

    payload: dict = {
        "chatType": params.chat_type,
        "members": member_objs,
    }
    if params.topic and params.chat_type != "oneOnOne":
        payload["topic"] = params.topic

    chat = GraphChat.model_validate(await g.post("/chats", json=payload) or {})
    return CreateTeamsChatResponse(
        success=True,
        action="teams_create_chat",
        id=chat.id,
        chat_type=chat.chat_type or params.chat_type,
        topic=chat.topic or params.topic,
        web_url=chat.web_url,
        created_at=chat.created_date_time,
        created_at_display=format_datetime_display(chat.created_date_time),
    )


# ---------------------------------------------------------------------------
# Online Meetings
# ---------------------------------------------------------------------------


async def teams_create_meeting(
    params: TeamsCreateMeetingInput,
) -> CreateTeamsMeetingResponse:
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
    g = get_graph(params.profile)
    payload: dict = {
        "subject": params.subject,
        "startDateTime": params.start_datetime,
        "endDateTime": params.end_datetime,
    }

    if params.attendees:
        payload["participants"] = {
            "attendees": [
                {"upn": upn.strip(), "role": "attendee"}
                for upn in params.attendees
                if upn.strip()
            ]
        }

    raw_meeting = (await g.post("/me/onlineMeetings", json=payload)) or {}
    meeting = GraphOnlineMeetingDetail.model_validate(
        {
            **raw_meeting,
            "subject": raw_meeting.get("subject", params.subject),
            "startDateTime": raw_meeting.get("startDateTime", params.start_datetime),
            "endDateTime": raw_meeting.get("endDateTime", params.end_datetime),
        }
    )
    return CreateTeamsMeetingResponse(
        success=True,
        action="teams_create_meeting",
        **_meeting_info(meeting).model_dump(),
    )


async def teams_get_meeting(
    params: TeamsGetMeetingInput,
) -> MeetingDetailResponse:
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
    g = get_graph(params.profile)
    meeting = GraphOnlineMeetingDetail.model_validate(
        await g.get(f"/me/onlineMeetings/{params.meeting_id}")
    )
    return MeetingDetailResponse(
        **_meeting_info(meeting).model_dump(),
        participants=_meeting_participants_info(meeting.participants),
        video_teleconference_id=meeting.video_teleconference_id,
    )


async def teams_list_meetings(
    params: TeamsListMeetingsInput,
) -> TeamsListMeetingsResponse:
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

    start_after = _normalize_filter_datetime(params.start_after, now - timedelta(days=7))
    start_before = _normalize_filter_datetime(params.start_before, now + timedelta(days=7))

    filter_str = (
        f"startDateTime ge '{start_after}' and startDateTime le '{start_before}'"
    )

    g = get_graph(params.profile)
    result = await g.get(
        "/me/onlineMeetings",
        params={
            "$filter": filter_str,
            "$top": params.top,
            "$select": "id,subject,startDateTime,endDateTime,joinWebUrl,createdDateTime",
        },
    )
    meetings = parse_graph_collection(result, GraphOnlineMeetingDetail)
    return TeamsListMeetingsResponse(
        start_after=start_after,
        start_before=start_before,
        count=len(meetings),
        meetings=[_meeting_info(meeting) for meeting in meetings],
        next_link=result.get("@odata.nextLink"),
    )


def register(server) -> None:
    """Register all Teams tools with the given FastMCP server instance."""
    # Teams & Channels
    register_tool(server, teams_list_joined, annotations=READ_ONLY_TOOL)
    register_tool(server, teams_get, annotations=READ_ONLY_TOOL)
    register_tool(server, teams_list_channels, annotations=READ_ONLY_TOOL)
    register_tool(server, teams_get_channel, annotations=READ_ONLY_TOOL)
    register_tool(server, teams_create_channel, annotations=WRITE_TOOL)
    # Channel Messages
    register_tool(server, teams_list_channel_messages, annotations=READ_ONLY_TOOL)
    register_tool(server, teams_get_channel_message, annotations=READ_ONLY_TOOL)
    register_tool(server, teams_send_channel_message, annotations=WRITE_TOOL)
    register_tool(server, teams_reply_to_channel_message, annotations=WRITE_TOOL)
    register_tool(server, teams_list_message_replies, annotations=READ_ONLY_TOOL)
    # Chats
    register_tool(server, teams_list_chats, annotations=READ_ONLY_TOOL)
    register_tool(server, teams_get_chat, annotations=READ_ONLY_TOOL)
    register_tool(server, teams_list_chat_messages, annotations=READ_ONLY_TOOL)
    register_tool(server, teams_send_chat_message, annotations=WRITE_TOOL)
    register_tool(server, teams_create_chat, annotations=WRITE_TOOL)
    # Online Meetings
    register_tool(server, teams_create_meeting, annotations=WRITE_TOOL)
    register_tool(server, teams_get_meeting, annotations=READ_ONLY_TOOL)
    register_tool(server, teams_list_meetings, annotations=READ_ONLY_TOOL)
