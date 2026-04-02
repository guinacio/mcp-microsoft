from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, RootModel


class MCPModel(BaseModel):
    model_config = ConfigDict(populate_by_name=True)


class Address(MCPModel):
    name: str = ""
    address: str = ""


class DisplayAddress(Address):
    display: str = ""


class AttachmentInfo(MCPModel):
    id: str = ""
    name: str = ""
    size_bytes: int = 0
    size_display: str = ""
    content_type: str = ""
    is_inline: bool = False
    size_kb: int = 0


class MessageSummary(MCPModel):
    id: str = ""
    subject: str = ""
    from_: DisplayAddress = Field(default_factory=DisplayAddress, alias="from")
    received_at: str | None = None
    received_at_display: str = ""
    is_read: bool = False
    has_attachments: bool = False
    importance: str = ""
    preview: str = ""


class ListEmailsResponse(MCPModel):
    folder: str = ""
    count: int = 0
    messages: list[MessageSummary] = Field(default_factory=list)
    next_page_token: int | None = None
    has_more: bool = False


class ReadEmailResponse(MCPModel):
    id: str = ""
    subject: str = ""
    from_: DisplayAddress = Field(default_factory=DisplayAddress, alias="from")
    to: list[Address] = Field(default_factory=list)
    to_display: str = ""
    cc: list[Address] = Field(default_factory=list)
    cc_display: str = ""
    received_at: str | None = None
    received_at_display: str = ""
    is_read: bool = False
    conversation_id: str = ""
    importance: str = ""
    body: str = ""
    body_content_type: str = ""
    attachments: list[AttachmentInfo] = Field(default_factory=list)


class ReadEmailSummaryResponse(MCPModel):
    id: str = ""
    subject: str = ""
    from_: DisplayAddress = Field(default_factory=DisplayAddress, alias="from")
    received_at: str | None = None
    received_at_display: str = ""
    is_read: bool = False
    preview: str = ""


class SearchEmailsResponse(MCPModel):
    query: str = ""
    folder: str | None = None
    count: int = 0
    messages: list[MessageSummary] = Field(default_factory=list)


class ActionResult(MCPModel):
    success: bool
    action: str
    error: str | None = None


class SendEmailResponse(ActionResult):
    to: list[str] = Field(default_factory=list)
    cc: list[str] = Field(default_factory=list)
    bcc: list[str] = Field(default_factory=list)
    reply_to: list[str] = Field(default_factory=list)
    subject: str = ""
    body_type: str = ""
    saved_to_sent_items: bool = False


class ReplyEmailResponse(ActionResult):
    message_id: str = ""
    body_type: str = ""


class ForwardEmailResponse(ActionResult):
    message_id: str = ""
    to: list[str] = Field(default_factory=list)
    comment: str = ""


class MarkEmailReadResponse(ActionResult):
    message_id: str = ""
    is_read: bool = False


class MoveEmailResponse(ActionResult):
    message_id: str = ""
    new_message_id: str = ""
    destination_folder: str = ""


class TrashEmailResponse(MoveEmailResponse):
    soft_delete: bool = True
    profile: str | None = None


class DeleteEmailResponse(ActionResult):
    message_id: str = ""
    irreversible: bool = True


class DraftSummary(MCPModel):
    id: str = ""
    subject: str = ""
    to: list[Address] = Field(default_factory=list)
    last_modified_at: str | None = None
    last_modified_at_display: str = ""
    preview: str = ""


class CreateDraftResponse(ActionResult):
    draft_id: str = ""
    to: list[str] = Field(default_factory=list)
    cc: list[str] = Field(default_factory=list)
    subject: str = ""
    body_type: str = ""


class ListDraftsResponse(MCPModel):
    count: int = 0
    drafts: list[DraftSummary] = Field(default_factory=list)


class DraftDetailResponse(MCPModel):
    id: str = ""
    subject: str = ""
    to: list[Address] = Field(default_factory=list)
    cc: list[Address] = Field(default_factory=list)
    bcc: list[Address] = Field(default_factory=list)
    last_modified_at: str | None = None
    last_modified_at_display: str = ""
    body: str = ""
    body_content_type: str = ""
    is_draft: bool = True


class UpdateDraftResponse(ActionResult):
    draft_id: str = ""
    updated_fields: list[str] = Field(default_factory=list)
    updated_fields_display: str = ""


class SendDraftResponse(ActionResult):
    draft_id: str = ""


class MailFolderInfo(MCPModel):
    id: str = ""
    display_name: str = ""
    display_label: str = ""
    unread_count: int = 0
    total_count: int = 0
    child_folder_count: int = 0
    is_child: bool = False


class ListFoldersResponse(MCPModel):
    count: int = 0
    include_child_folders: bool = False
    folders: list[MailFolderInfo] = Field(default_factory=list)


class CreateFolderResponse(ActionResult):
    folder_id: str = ""
    display_name: str = ""
    parent_folder_id: str | None = None


class DeleteFolderResponse(ActionResult):
    folder_id: str = ""
    irreversible: bool = True


class ListAttachmentsResponse(MCPModel):
    message_id: str = ""
    count: int = 0
    attachments: list[AttachmentInfo] = Field(default_factory=list)


class DownloadAttachmentResponse(ActionResult):
    message_id: str = ""
    attachment_id: str = ""
    filename: str = ""
    path: str | None = None
    size_bytes: int = 0
    size_display: str = ""
    content_type: str = ""


class CalendarInfo(MCPModel):
    id: str = ""
    name: str = ""
    color: str = ""
    is_default: bool = False
    can_edit: bool = False


class ListCalendarsResponse(MCPModel):
    count: int = 0
    calendars: list[CalendarInfo] = Field(default_factory=list)


class AttendeeInfo(MCPModel):
    name: str = ""
    address: str = ""
    type: str = ""
    response: str = ""


class EventSummary(MCPModel):
    id: str = ""
    subject: str = ""
    start: str | None = None
    start_display: str = ""
    end: str | None = None
    end_display: str = ""
    timezone: str = ""
    location: str = ""
    is_all_day: bool = False
    is_cancelled: bool = False
    response_status: str = ""


class ListEventsResponse(MCPModel):
    calendar_id: str | None = None
    filter_start: str | None = None
    count: int = 0
    events: list[EventSummary] = Field(default_factory=list)


class ListUpcomingEventsResponse(MCPModel):
    calendar_id: str | None = None
    start_datetime: str = ""
    end_datetime: str = ""
    count: int = 0
    events: list[EventSummary] = Field(default_factory=list)


class EventDetailResponse(MCPModel):
    id: str = ""
    subject: str = ""
    start: str | None = None
    start_display: str = ""
    end: str | None = None
    end_display: str = ""
    timezone: str = ""
    location: str = ""
    organizer: DisplayAddress = Field(default_factory=DisplayAddress)
    attendees: list[AttendeeInfo] = Field(default_factory=list)
    attendees_display: str = ""
    is_all_day: bool = False
    is_cancelled: bool = False
    show_as: str = ""
    web_link: str = ""
    join_url: str = ""
    body: str = ""
    body_content_type: str = ""
    recurrence: dict[str, Any] | None = None
    recurrence_display: str = ""
    importance: str = ""
    sensitivity: str = ""


class CreateEventResponse(ActionResult):
    event_id: str = ""
    subject: str = ""
    start_datetime: str = ""
    end_datetime: str = ""
    timezone: str = ""
    calendar_id: str | None = None
    web_link: str = ""
    join_url: str = ""


class UpdateEventResponse(ActionResult):
    event_id: str = ""
    updated_fields: list[str] = Field(default_factory=list)
    updated_fields_display: str = ""


class DeleteEventResponse(ActionResult):
    event_id: str = ""


class RsvpEventResponse(ActionResult):
    event_id: str = ""
    response: str = ""
    comment: str | None = None
    send_response: bool = True


class ScheduleItemInfo(MCPModel):
    status: str = ""
    start: str | None = None
    start_display: str = ""
    end: str | None = None
    end_display: str = ""
    subject: str = ""


class PersonSchedule(MCPModel):
    email: str = ""
    availability_view: str = ""
    legend: dict[str, str] = Field(default_factory=dict)
    schedule_items: list[ScheduleItemInfo] = Field(default_factory=list)


class FreeBusyResponse(MCPModel):
    start_datetime: str = ""
    end_datetime: str = ""
    timezone: str = ""
    people: list[PersonSchedule] = Field(default_factory=list)


class AttendeeAvailabilityInfo(MCPModel):
    email: str = ""
    availability: str = ""


class MeetingSuggestion(MCPModel):
    confidence: float | int = 0
    start: str | None = None
    start_display: str = ""
    end: str | None = None
    end_display: str = ""
    attendee_availability: list[AttendeeAvailabilityInfo] = Field(default_factory=list)


class MeetingSuggestionsResponse(MCPModel):
    count: int = 0
    suggestions: list[MeetingSuggestion] = Field(default_factory=list)
    empty_suggestions_reason: str | None = None
    timezone: str = ""


class DriveItemInfo(MCPModel):
    id: str = ""
    name: str = ""
    size_bytes: int = 0
    size_display: str = ""
    last_modified_at: str | None = None
    last_modified_at_display: str = ""
    web_url: str = ""
    is_folder: bool = False
    child_count: int = 0
    mime_type: str = ""
    parent_path: str = ""


class ListDriveItemsResponse(MCPModel):
    folder_id: str | None = None
    count: int = 0
    items: list[DriveItemInfo] = Field(default_factory=list)
    has_more: bool = False


class DriveItemDetailResponse(MCPModel):
    id: str = ""
    name: str = ""
    type: str = ""
    size_bytes: int = 0
    size_display: str = ""
    created_at: str | None = None
    created_at_display: str = ""
    created_by: str = ""
    modified_at: str | None = None
    modified_at_display: str = ""
    modified_by: str = ""
    path: str = ""
    web_url: str = ""
    child_count: int = 0
    mime_type: str = ""


class SearchDriveResponse(MCPModel):
    query: str = ""
    count: int = 0
    items: list[DriveItemInfo] = Field(default_factory=list)


class CreateDriveFolderResponse(ActionResult):
    folder_id: str = ""
    name: str = ""
    parent_folder_id: str | None = None
    web_url: str = ""


class UploadFileResponse(ActionResult):
    filename: str = ""
    size_bytes: int = 0
    size_display: str = ""
    file_id: str = ""
    web_url: str = ""
    parent_folder_id: str | None = None
    path: str | None = None


class DownloadFileResponse(ActionResult):
    item_id: str = ""
    path: str = ""
    filename: str = ""
    size_bytes: int = 0
    size_display: str = ""
    expected_size_bytes: int = 0


class DeleteDriveItemResponse(ActionResult):
    item_id: str = ""
    soft_delete: bool = True


class MoveOrCopyItemResponse(ActionResult):
    item_id: str = ""
    new_item_id: str = ""
    destination_folder_id: str = ""
    name: str | None = None
    status: str | None = None


class SharePointSiteInfo(MCPModel):
    id: str = ""
    display_name: str = ""
    description: str = ""
    web_url: str = ""
    created_at: str | None = None
    created_at_display: str = ""
    last_modified_at: str | None = None
    last_modified_at_display: str = ""


class SearchSharePointSitesResponse(BaseModel):
    query: str = ""
    count: int = 0
    sites: list[SharePointSiteInfo] = Field(default_factory=list)
    has_more: bool = False


class SharePointSiteDetailResponse(SharePointSiteInfo):
    pass


class SharePointLibraryInfo(MCPModel):
    id: str = ""
    name: str = ""
    description: str = ""
    drive_type: str = ""
    web_url: str = ""


class ListSiteLibrariesResponse(MCPModel):
    site_id: str = ""
    count: int = 0
    libraries: list[SharePointLibraryInfo] = Field(default_factory=list)


class ListSiteFilesResponse(MCPModel):
    site_id: str = ""
    drive_id: str = ""
    folder_id: str | None = None
    count: int = 0
    items: list[DriveItemInfo] = Field(default_factory=list)
    has_more: bool = False


class SiteFileDetailResponse(MCPModel):
    site_id: str = ""
    drive_id: str = ""
    id: str = ""
    name: str = ""
    type: str = ""
    size_bytes: int = 0
    size_display: str = ""
    created_at: str | None = None
    created_at_display: str = ""
    created_by: str = ""
    modified_at: str | None = None
    modified_at_display: str = ""
    modified_by: str = ""
    path: str = ""
    child_count: int = 0
    mime_type: str = ""
    web_url: str = ""


class UploadSiteFileResponse(ActionResult):
    site_id: str = ""
    drive_id: str = ""
    folder_id: str | None = None
    filename: str = ""
    size_bytes: int = 0
    size_display: str = ""
    file_id: str = ""
    web_url: str = ""
    path: str | None = None


class DownloadSiteFileResponse(ActionResult):
    site_id: str = ""
    drive_id: str = ""
    item_id: str = ""
    path: str = ""
    filename: str = ""
    size_bytes: int = 0
    size_display: str = ""


class SharePointListInfo(MCPModel):
    id: str = ""
    display_name: str = ""
    description: str = ""
    web_url: str = ""
    template: str = ""


class ListSiteListsResponse(MCPModel):
    site_id: str = ""
    count: int = 0
    lists: list[SharePointListInfo] = Field(default_factory=list)


class SharePointListItemInfo(MCPModel):
    id: str = ""
    title: str = ""
    created_at: str | None = None
    created_at_display: str = ""
    modified_at: str | None = None
    modified_at_display: str = ""
    fields: dict[str, Any] = Field(default_factory=dict)


class GetListItemsResponse(MCPModel):
    site_id: str = ""
    list_id: str = ""
    count: int = 0
    items: list[SharePointListItemInfo] = Field(default_factory=list)
    has_more: bool = False


class SharePointFields(RootModel[dict[str, Any]]):
    root: dict[str, Any]


class CreateListItemResponse(ActionResult):
    site_id: str = ""
    list_id: str = ""
    item_id: str = ""
    fields: dict[str, Any] = Field(default_factory=dict)


class UpdateListItemResponse(ActionResult):
    site_id: str = ""
    list_id: str = ""
    item_id: str = ""
    updated_fields: list[str] = Field(default_factory=list)
    fields: dict[str, Any] = Field(default_factory=dict)


class DeleteListItemResponse(ActionResult):
    site_id: str = ""
    list_id: str = ""
    item_id: str = ""


class ProfileInfo(MCPModel):
    name: str = ""
    client_id_masked: str = ""
    tenant_id: str = ""
    is_default: bool = False
    is_authenticated: bool | None = None
    cache_path: str = ""


class ListProfilesResponse(MCPModel):
    default_profile: str | None = None
    count: int = 0
    profiles: list[ProfileInfo] = Field(default_factory=list)


class AddedProfileInfo(MCPModel):
    name: str = ""
    client_id: str = ""
    tenant_id: str = ""
    cache_path: str = ""
    is_default: bool = False


class AddProfileResponse(ActionResult):
    profile: AddedProfileInfo | None = None


class RemoveProfileResponse(ActionResult):
    profile: str = ""


class AuthenticateProfileResponse(ActionResult):
    profile: str | None = None
    tenant_id: str = ""
    cache_path: str = ""


class SetDefaultProfileResponse(ActionResult):
    profile: str = ""


# ---------------------------------------------------------------------------
# Contacts models
# ---------------------------------------------------------------------------


class ContactEmailAddress(MCPModel):
    address: str = ""
    name: str = ""


class ContactInfo(MCPModel):
    id: str = ""
    display_name: str = ""
    given_name: str = ""
    surname: str = ""
    email_addresses: list[ContactEmailAddress] = Field(default_factory=list)
    mobile_phone: str = ""
    business_phones: list[str] = Field(default_factory=list)
    job_title: str = ""
    company_name: str = ""
    department: str = ""


class ListContactsResponse(MCPModel):
    count: int = 0
    folder_id: str | None = None
    contacts: list[ContactInfo] = Field(default_factory=list)


class GetContactResponse(MCPModel):
    id: str = ""
    display_name: str = ""
    given_name: str = ""
    surname: str = ""
    email_addresses: list[ContactEmailAddress] = Field(default_factory=list)
    mobile_phone: str = ""
    business_phones: list[str] = Field(default_factory=list)
    job_title: str = ""
    company_name: str = ""
    department: str = ""
    notes: str = ""


class CreateContactResponse(ActionResult):
    contact_id: str = ""
    display_name: str = ""
    folder_id: str | None = None


class UpdateContactResponse(ActionResult):
    contact_id: str = ""
    updated_fields: list[str] = Field(default_factory=list)
    updated_fields_display: str = ""


class DeleteContactResponse(ActionResult):
    contact_id: str = ""


class ContactFolderInfo(MCPModel):
    id: str = ""
    display_name: str = ""
    parent_folder_id: str = ""
    total_item_count: int = 0


class ListContactFoldersResponse(MCPModel):
    count: int = 0
    folders: list[ContactFolderInfo] = Field(default_factory=list)


class SearchContactsResponse(MCPModel):
    query: str = ""
    count: int = 0
    contacts: list[ContactInfo] = Field(default_factory=list)


class GetContactPhotoResponse(ActionResult):
    contact_id: str = ""
    photo_base64: str = ""
    saved_path: str = ""
    size_bytes: int = 0
