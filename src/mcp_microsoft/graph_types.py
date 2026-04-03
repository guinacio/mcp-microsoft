from __future__ import annotations

from typing import Any, TypeVar

from pydantic import BaseModel, ConfigDict, Field


class GraphModel(BaseModel):
    model_config = ConfigDict(populate_by_name=True, extra="ignore")


class FlexibleGraphModel(GraphModel):
    model_config = ConfigDict(populate_by_name=True, extra="allow")


class GraphEmailAddress(GraphModel):
    name: str = ""
    address: str = ""

    @property
    def display(self) -> str:
        return f"{self.name} <{self.address}>" if self.name else self.address


class GraphRecipient(GraphModel):
    email_address: GraphEmailAddress = Field(
        default_factory=GraphEmailAddress,
        alias="emailAddress",
    )


class GraphItemBody(GraphModel):
    content_type: str = Field(default="text", alias="contentType")
    content: str = ""


class GraphAttachment(GraphModel):
    id: str = ""
    name: str = "unnamed"
    size: int = 0
    content_type: str = Field(default="", alias="contentType")
    is_inline: bool = Field(default=False, alias="isInline")
    content_bytes: str | None = Field(default=None, alias="contentBytes")


class GraphSender(GraphModel):
    email_address: GraphEmailAddress = Field(
        default_factory=GraphEmailAddress,
        alias="emailAddress",
    )


class GraphMessage(GraphModel):
    id: str = ""
    subject: str = ""
    from_: GraphSender | None = Field(default=None, alias="from")
    to_recipients: list[GraphRecipient] = Field(default_factory=list, alias="toRecipients")
    cc_recipients: list[GraphRecipient] = Field(default_factory=list, alias="ccRecipients")
    bcc_recipients: list[GraphRecipient] = Field(default_factory=list, alias="bccRecipients")
    body: GraphItemBody = Field(default_factory=GraphItemBody)
    body_preview: str = Field(default="", alias="bodyPreview")
    received_date_time: str | None = Field(default=None, alias="receivedDateTime")
    is_read: bool = Field(default=False, alias="isRead")
    conversation_id: str = Field(default="", alias="conversationId")
    has_attachments: bool = Field(default=False, alias="hasAttachments")
    importance: str = ""
    attachments: list[GraphAttachment] = Field(default_factory=list)
    last_modified_date_time: str | None = Field(default=None, alias="lastModifiedDateTime")


class GraphCalendar(GraphModel):
    id: str = ""
    name: str = ""
    color: str = ""
    is_default_calendar: bool = Field(default=False, alias="isDefaultCalendar")
    can_edit: bool = Field(default=False, alias="canEdit")


class GraphDateTimeTimeZone(GraphModel):
    date_time: str | None = Field(default=None, alias="dateTime")
    time_zone: str = Field(default="", alias="timeZone")


class GraphLocation(GraphModel):
    display_name: str = Field(default="", alias="displayName")


class GraphResponseStatus(GraphModel):
    response: str = ""


class GraphAttendeeStatus(GraphModel):
    response: str = ""


class GraphAttendee(GraphModel):
    email_address: GraphEmailAddress = Field(
        default_factory=GraphEmailAddress,
        alias="emailAddress",
    )
    type: str = ""
    status: GraphAttendeeStatus = Field(default_factory=GraphAttendeeStatus)


class GraphOrganizer(GraphModel):
    email_address: GraphEmailAddress = Field(
        default_factory=GraphEmailAddress,
        alias="emailAddress",
    )


class GraphOnlineMeeting(GraphModel):
    join_url: str = Field(default="", alias="joinUrl")


class GraphRecurrencePattern(GraphModel):
    type: str = ""
    interval: int = 1


class GraphRecurrenceRange(FlexibleGraphModel):
    type: str = ""
    start_date: str = Field(default="", alias="startDate")
    end_date: str = Field(default="", alias="endDate")
    recurrence_time_zone: str = Field(default="", alias="recurrenceTimeZone")
    number_of_occurrences: int | None = Field(default=None, alias="numberOfOccurrences")


class GraphPatternedRecurrence(GraphModel):
    pattern: GraphRecurrencePattern | None = None
    range_: GraphRecurrenceRange | None = Field(default=None, alias="range")


class GraphEvent(GraphModel):
    id: str = ""
    subject: str = ""
    body: GraphItemBody = Field(default_factory=GraphItemBody)
    start: GraphDateTimeTimeZone = Field(default_factory=GraphDateTimeTimeZone)
    end: GraphDateTimeTimeZone = Field(default_factory=GraphDateTimeTimeZone)
    location: GraphLocation = Field(default_factory=GraphLocation)
    organizer: GraphOrganizer = Field(default_factory=GraphOrganizer)
    attendees: list[GraphAttendee] = Field(default_factory=list)
    is_all_day: bool = Field(default=False, alias="isAllDay")
    is_cancelled: bool = Field(default=False, alias="isCancelled")
    recurrence: GraphPatternedRecurrence | None = None
    online_meeting: GraphOnlineMeeting | None = Field(default=None, alias="onlineMeeting")
    web_link: str = Field(default="", alias="webLink")
    response_status: GraphResponseStatus = Field(default_factory=GraphResponseStatus, alias="responseStatus")
    importance: str = ""
    sensitivity: str = ""
    show_as: str = Field(default="", alias="showAs")


class GraphIdentity(GraphModel):
    id: str = ""
    display_name: str = Field(default="", alias="displayName")
    tenant_id: str = Field(default="", alias="tenantId")


class GraphIdentitySet(GraphModel):
    user: GraphIdentity | None = None
    application: GraphIdentity | None = None
    device: GraphIdentity | None = None


class GraphTeamMemberSettings(FlexibleGraphModel):
    allow_create_update_channels: bool | None = Field(default=None, alias="allowCreateUpdateChannels")
    allow_delete_channels: bool | None = Field(default=None, alias="allowDeleteChannels")
    allow_add_remove_apps: bool | None = Field(default=None, alias="allowAddRemoveApps")
    allow_create_update_remove_tabs: bool | None = Field(default=None, alias="allowCreateUpdateRemoveTabs")
    allow_create_update_remove_connectors: bool | None = Field(default=None, alias="allowCreateUpdateRemoveConnectors")


class GraphTeamGuestSettings(FlexibleGraphModel):
    allow_create_update_channels: bool | None = Field(default=None, alias="allowCreateUpdateChannels")
    allow_delete_channels: bool | None = Field(default=None, alias="allowDeleteChannels")


class GraphTeamFunSettings(FlexibleGraphModel):
    allow_giphy: bool | None = Field(default=None, alias="allowGiphy")
    giphy_content_rating: str = Field(default="", alias="giphyContentRating")
    allow_stickers_and_memes: bool | None = Field(default=None, alias="allowStickersAndMemes")
    allow_custom_memes: bool | None = Field(default=None, alias="allowCustomMemes")


class GraphTeam(GraphModel):
    id: str = ""
    display_name: str = Field(default="", alias="displayName")
    description: str = ""
    visibility: str = ""
    web_url: str = Field(default="", alias="webUrl")
    is_archived: bool | None = Field(default=None, alias="isArchived")
    member_settings: GraphTeamMemberSettings = Field(default_factory=GraphTeamMemberSettings, alias="memberSettings")
    guest_settings: GraphTeamGuestSettings = Field(default_factory=GraphTeamGuestSettings, alias="guestSettings")
    fun_settings: GraphTeamFunSettings = Field(default_factory=GraphTeamFunSettings, alias="funSettings")


class GraphChannel(GraphModel):
    id: str = ""
    display_name: str = Field(default="", alias="displayName")
    description: str = ""
    channel_type: str = Field(default="", alias="channelType")
    web_url: str = Field(default="", alias="webUrl")
    is_favorite_by_default: bool = Field(default=False, alias="isFavoriteByDefault")


class GraphChatMessageReaction(FlexibleGraphModel):
    reaction_type: str = Field(default="", alias="reactionType")
    created_date_time: str = Field(default="", alias="createdDateTime")
    user: GraphIdentitySet | None = None


class GraphChatMessageAttachment(FlexibleGraphModel):
    id: str = ""
    name: str = ""
    content_type: str = Field(default="", alias="contentType")
    content_url: str = Field(default="", alias="contentUrl")
    thumbnail_url: str = Field(default="", alias="thumbnailUrl")


class GraphChatMessageMentionedIdentitySet(FlexibleGraphModel):
    user: GraphIdentity | None = None
    application: GraphIdentity | None = None
    device: GraphIdentity | None = None
    conversation: GraphIdentity | None = None
    tag: GraphIdentity | None = None


class GraphChatMessageMention(FlexibleGraphModel):
    id: int | str | None = None
    mention_text: str = Field(default="", alias="mentionText")
    mentioned: GraphChatMessageMentionedIdentitySet | None = None


class GraphChatMessage(GraphModel):
    id: str = ""
    created_date_time: str = Field(default="", alias="createdDateTime")
    last_modified_date_time: str = Field(default="", alias="lastModifiedDateTime")
    from_: GraphIdentitySet | None = Field(default=None, alias="from")
    body: GraphItemBody = Field(default_factory=GraphItemBody)
    subject: str = ""
    web_url: str = Field(default="", alias="webUrl")
    reply_to_id: str = Field(default="", alias="replyToId")
    importance: str = ""
    reactions: list[GraphChatMessageReaction] = Field(default_factory=list)
    attachments: list[GraphChatMessageAttachment] = Field(default_factory=list)
    mentions: list[GraphChatMessageMention] = Field(default_factory=list)


class GraphChat(GraphModel):
    id: str = ""
    chat_type: str = Field(default="", alias="chatType")
    topic: str = ""
    created_date_time: str = Field(default="", alias="createdDateTime")
    last_updated_date_time: str = Field(default="", alias="lastUpdatedDateTime")
    web_url: str = Field(default="", alias="webUrl")


class GraphChatMember(GraphModel):
    id: str = ""
    display_name: str = Field(default="", alias="displayName")
    email: str = ""
    user_id: str = Field(default="", alias="userId")
    tenant_id: str = Field(default="", alias="tenantId")
    roles: list[str] = Field(default_factory=list)


class GraphJoinMeetingIdSettings(GraphModel):
    join_meeting_id: str = Field(default="", alias="joinMeetingId")


class GraphMeetingParticipantInfo(FlexibleGraphModel):
    upn: str = ""
    role: str = ""
    identity: GraphIdentitySet | None = None


class GraphMeetingParticipants(FlexibleGraphModel):
    organizer: GraphMeetingParticipantInfo | None = None
    attendees: list[GraphMeetingParticipantInfo] = Field(default_factory=list)
    producers: list[GraphMeetingParticipantInfo] = Field(default_factory=list)
    contributors: list[GraphMeetingParticipantInfo] = Field(default_factory=list)


class GraphOnlineMeetingDetail(GraphModel):
    id: str = ""
    subject: str = ""
    join_web_url: str = Field(default="", alias="joinWebUrl")
    join_meeting_id_settings: GraphJoinMeetingIdSettings | None = Field(
        default=None,
        alias="joinMeetingIdSettings",
    )
    start_date_time: str = Field(default="", alias="startDateTime")
    end_date_time: str = Field(default="", alias="endDateTime")
    created_date_time: str = Field(default="", alias="createdDateTime")
    participants: GraphMeetingParticipants = Field(default_factory=GraphMeetingParticipants)
    video_teleconference_id: str = Field(default="", alias="videoTeleconferenceId")


GraphModelT = TypeVar("GraphModelT", bound=GraphModel)


def parse_graph_collection(payload: dict[str, Any], model: type[GraphModelT]) -> list[GraphModelT]:
    return [model.model_validate(item) for item in payload.get("value") or []]
