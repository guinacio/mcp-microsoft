from __future__ import annotations

from typing import Any, TypeVar, get_args, get_origin

from pydantic import BaseModel, ConfigDict, Field
from pydantic_core import PydanticUndefined
from pydantic import model_validator


def _annotation_allows_none(annotation: Any) -> bool:
    origin = get_origin(annotation)
    if origin is None:
        return annotation is type(None)
    return any(arg is type(None) for arg in get_args(annotation))


class GraphModel(BaseModel):
    model_config = ConfigDict(populate_by_name=True, extra="ignore")

    @model_validator(mode="before")
    @classmethod
    def _coerce_nulls_to_defaults(cls, data: Any) -> Any:
        if not isinstance(data, dict):
            return data

        normalized = dict(data)
        for field_name, field in cls.model_fields.items():
            candidate_keys = [field_name]
            if field.alias and field.alias not in candidate_keys:
                candidate_keys.append(field.alias)

            for key in candidate_keys:
                if key not in normalized or normalized[key] is not None:
                    continue
                if _annotation_allows_none(field.annotation):
                    break
                if field.default_factory is not None:
                    normalized[key] = field.default_factory()
                    break
                if field.default is not PydanticUndefined:
                    normalized[key] = field.default
                    break
        return normalized


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


class GraphContactEmailAddress(GraphModel):
    address: str = ""
    name: str = ""


class GraphContact(GraphModel):
    id: str = ""
    display_name: str | None = Field(default=None, alias="displayName")
    given_name: str | None = Field(default=None, alias="givenName")
    surname: str | None = None
    email_addresses: list[GraphContactEmailAddress] | None = Field(default=None, alias="emailAddresses")
    mobile_phone: str | None = Field(default=None, alias="mobilePhone")
    business_phones: list[str] | None = Field(default=None, alias="businessPhones")
    job_title: str | None = Field(default=None, alias="jobTitle")
    company_name: str | None = Field(default=None, alias="companyName")
    department: str | None = None
    personal_notes: str | None = Field(default=None, alias="personalNotes")


class GraphContactFolder(GraphModel):
    id: str = ""
    display_name: str = Field(default="", alias="displayName")
    parent_folder_id: str | None = Field(default=None, alias="parentFolderId")


class GraphSender(GraphModel):
    email_address: GraphEmailAddress = Field(
        default_factory=GraphEmailAddress,
        alias="emailAddress",
    )


class GraphMessage(GraphModel):
    id: str = ""
    subject: str | None = None
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
    is_draft: bool = Field(default=False, alias="isDraft")


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


class GraphMailFolder(GraphModel):
    id: str = ""
    display_name: str = Field(default="", alias="displayName")
    total_item_count: int | None = Field(default=None, alias="totalItemCount")
    unread_item_count: int | None = Field(default=None, alias="unreadItemCount")
    child_folder_count: int | None = Field(default=None, alias="childFolderCount")


class GraphFileFacet(FlexibleGraphModel):
    mime_type: str = Field(default="", alias="mimeType")


class GraphFolderFacet(FlexibleGraphModel):
    child_count: int = Field(default=0, alias="childCount")


class GraphParentReference(FlexibleGraphModel):
    drive_id: str = Field(default="", alias="driveId")
    site_id: str = Field(default="", alias="siteId")
    path: str = ""
    id: str = ""


class GraphDriveItem(GraphModel):
    id: str = ""
    name: str = ""
    size: int | None = None  # null on folders, notebooks, and shortcuts
    created_date_time: str | None = Field(default=None, alias="createdDateTime")
    last_modified_date_time: str | None = Field(default=None, alias="lastModifiedDateTime")
    web_url: str = Field(default="", alias="webUrl")
    file: GraphFileFacet | None = None
    folder: GraphFolderFacet | None = None
    parent_reference: GraphParentReference | None = Field(default=None, alias="parentReference")
    created_by: GraphIdentitySet | None = Field(default=None, alias="createdBy")
    last_modified_by: GraphIdentitySet | None = Field(default=None, alias="lastModifiedBy")


class GraphSite(GraphModel):
    id: str = ""
    display_name: str = Field(default="", alias="displayName")
    description: str = ""
    web_url: str = Field(default="", alias="webUrl")
    created_date_time: str | None = Field(default=None, alias="createdDateTime")
    last_modified_date_time: str | None = Field(default=None, alias="lastModifiedDateTime")


class GraphDrive(GraphModel):
    id: str = ""
    name: str = ""
    description: str = ""
    drive_type: str = Field(default="", alias="driveType")
    web_url: str = Field(default="", alias="webUrl")


class GraphSharePointListFacet(FlexibleGraphModel):
    template: str = ""


class GraphSharePointList(GraphModel):
    id: str = ""
    display_name: str = Field(default="", alias="displayName")
    description: str = ""
    web_url: str = Field(default="", alias="webUrl")
    list_: GraphSharePointListFacet | None = Field(default=None, alias="list")


class GraphSharePointListItem(FlexibleGraphModel):
    id: str = ""
    created_date_time: str | None = Field(default=None, alias="createdDateTime")
    last_modified_date_time: str | None = Field(default=None, alias="lastModifiedDateTime")
    fields: dict[str, Any] = Field(default_factory=dict)


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


class GraphCallTranscript(GraphModel):
    id: str = ""
    meeting_id: str = Field(default="", alias="meetingId")
    call_id: str = Field(default="", alias="callId")
    created_date_time: str = Field(default="", alias="createdDateTime")
    end_date_time: str = Field(default="", alias="endDateTime")
    content_correlation_id: str = Field(default="", alias="contentCorrelationId")
    transcript_content_url: str = Field(default="", alias="transcriptContentUrl")
    meeting_organizer: GraphIdentitySet | None = Field(default=None, alias="meetingOrganizer")


class GraphCallRecording(GraphModel):
    id: str = ""
    meeting_id: str = Field(default="", alias="meetingId")
    call_id: str = Field(default="", alias="callId")
    created_date_time: str = Field(default="", alias="createdDateTime")
    end_date_time: str = Field(default="", alias="endDateTime")
    content_correlation_id: str = Field(default="", alias="contentCorrelationId")
    recording_content_url: str = Field(default="", alias="recordingContentUrl")
    meeting_organizer: GraphIdentitySet | None = Field(default=None, alias="meetingOrganizer")


class GraphCallAiInsightNoteSubpoint(GraphModel):
    title: str = ""
    text: str = ""


class GraphCallAiInsightNote(GraphModel):
    title: str = ""
    text: str = ""
    subpoints: list[GraphCallAiInsightNoteSubpoint] = Field(default_factory=list)


class GraphCallAiInsightActionItem(GraphModel):
    title: str = ""
    text: str = ""
    owner_display_name: str = Field(default="", alias="ownerDisplayName")


class GraphCallAiInsightMentionEvent(GraphModel):
    speaker: GraphIdentitySet | None = None
    event_date_time: str = Field(default="", alias="eventDateTime")
    transcript_utterance: str = Field(default="", alias="transcriptUtterance")


class GraphCallAiInsightViewpoint(FlexibleGraphModel):
    mention_events: list[GraphCallAiInsightMentionEvent] = Field(default_factory=list, alias="mentionEvents")


class GraphCallAiInsight(GraphModel):
    id: str = ""
    call_id: str = Field(default="", alias="callId")
    content_correlation_id: str = Field(default="", alias="contentCorrelationId")
    created_date_time: str = Field(default="", alias="createdDateTime")
    end_date_time: str = Field(default="", alias="endDateTime")
    meeting_notes: list[GraphCallAiInsightNote] = Field(default_factory=list, alias="meetingNotes")
    action_items: list[GraphCallAiInsightActionItem] = Field(default_factory=list, alias="actionItems")
    viewpoint: GraphCallAiInsightViewpoint | None = None


GraphModelT = TypeVar("GraphModelT", bound=GraphModel)


def parse_graph_collection(payload: dict[str, Any], model: type[GraphModelT]) -> list[GraphModelT]:
    return [model.model_validate(item) for item in payload.get("value") or []]


def graph_identity_display(identity_set: GraphIdentitySet | None) -> str:
    if identity_set is None:
        return ""
    if identity_set.user and identity_set.user.display_name:
        return identity_set.user.display_name
    if identity_set.application and identity_set.application.display_name:
        return identity_set.application.display_name
    if identity_set.device and identity_set.device.display_name:
        return identity_set.device.display_name
    return ""
