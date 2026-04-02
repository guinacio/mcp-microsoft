"""
Calendar tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.
Endpoints live under /me/calendar and /me/calendars in Graph v1.0.

Implemented:
  - list_calendars
  - list_events
  - list_upcoming_events
  - get_event
  - create_event
  - update_event
  - delete_event
  - rsvp_event
  - get_free_busy
  - find_meeting_times
"""

from __future__ import annotations

from datetime import datetime
from typing import Any, Literal, Optional, Union

from mcp.types import ToolAnnotations

from mcp_microsoft.models import (
    AttendeeAvailabilityInfo,
    AttendeeInfo,
    CreateEventResponse,
    DeleteEventResponse,
    DisplayAddress,
    EventDetailResponse,
    EventSummary,
    FreeBusyResponse,
    ListCalendarsResponse,
    ListEventsResponse,
    ListUpcomingEventsResponse,
    MeetingSuggestion,
    MeetingSuggestionsResponse,
    PersonSchedule,
    RsvpEventResponse,
    ScheduleItemInfo,
    CalendarInfo,
    UpdateEventResponse,
)
from mcp_microsoft.graph import get_graph
from mcp_microsoft.server import mcp

# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

BodyType = Literal["text", "html"]
RsvpResponse = Literal["accept", "decline", "tentativelyAccept"]
_READ_ONLY = ToolAnnotations(readOnlyHint=True, openWorldHint=True)
_WRITE = ToolAnnotations(destructiveHint=False, openWorldHint=True)
_DESTRUCTIVE = ToolAnnotations(destructiveHint=True, openWorldHint=True)


def _fmt_dt(iso: Optional[str]) -> str:
    """Format an ISO 8601 datetime to a human-readable form."""
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return iso


def _fmt_attendees(attendees: list[dict]) -> str:
    """Format Graph attendee objects as a readable string."""
    parts = []
    for att in attendees or []:
        ea = att.get("emailAddress", {})
        name = ea.get("name", "")
        addr = ea.get("address", "")
        status = att.get("status", {}).get("response", "")
        label = f"{name} <{addr}>" if name else addr
        if status and status != "none":
            label += f" [{status}]"
        parts.append(label)
    return ", ".join(parts) if parts else "(none)"


def _build_datetime(dt_str: str, tz: str) -> dict:
    """Build a Graph dateTimeTimeZone object."""
    return {"dateTime": dt_str, "timeZone": tz}


def _parse_attendees(
    emails: Optional[Union[str, list[str]]],
    attendee_type: str = "required",
) -> list[dict]:
    """Convert email addresses to Graph attendee format."""
    if not emails:
        return []
    if isinstance(emails, str):
        addresses = [addr.strip() for addr in emails.split(",") if addr.strip()]
    else:
        addresses = [addr.strip() for addr in emails if addr.strip()]
    return [
        {
            "emailAddress": {"address": addr},
            "type": attendee_type,
        }
        for addr in addresses
    ]


def _attendee_values(attendees: list[dict]) -> list[AttendeeInfo]:
    """Normalize Graph attendee objects into simple dictionaries."""
    values: list[AttendeeInfo] = []
    for attendee in attendees or []:
        email_address = attendee.get("emailAddress", {})
        values.append(
            AttendeeInfo(
                name=email_address.get("name", ""),
                address=email_address.get("address", ""),
                type=attendee.get("type", ""),
                response=attendee.get("status", {}).get("response", ""),
            )
        )
    return values


def _event_summary(event: dict[str, Any]) -> EventSummary:
    """Normalize a Graph event into a summary payload."""
    return EventSummary(
        id=event.get("id", ""),
        subject=event.get("subject") or "(no subject)",
        start=event.get("start", {}).get("dateTime"),
        start_display=_fmt_dt(event.get("start", {}).get("dateTime")),
        end=event.get("end", {}).get("dateTime"),
        end_display=_fmt_dt(event.get("end", {}).get("dateTime")),
        timezone=event.get("start", {}).get("timeZone", ""),
        location=event.get("location", {}).get("displayName", ""),
        is_all_day=event.get("isAllDay", False),
        is_cancelled=event.get("isCancelled", False),
        response_status=event.get("responseStatus", {}).get("response", ""),
    )


# ---------------------------------------------------------------------------
# list_calendars
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def list_calendars(profile: str | None = None) -> ListCalendarsResponse:
    """
    List all calendars in the user's mailbox.

    Args:
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured calendar metadata.
    """
    g = get_graph(profile)
    params = {
        "$select": "id,name,color,isDefaultCalendar,canEdit",
        "$top": 50,
    }

    result = await g.get("/me/calendars", params=params)
    calendars = result.get("value", [])

    return ListCalendarsResponse(
        count=len(calendars),
        calendars=[
            CalendarInfo(
                id=cal.get("id", ""),
                name=cal.get("name", "(unnamed)"),
                color=cal.get("color", "auto"),
                is_default=cal.get("isDefaultCalendar", False),
                can_edit=cal.get("canEdit", False),
            )
            for cal in calendars
        ],
    )


# ---------------------------------------------------------------------------
# list_events
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def list_events(
    max_results: int = 10,
    filter_start: Optional[str] = None,
    calendar_id: Optional[str] = None,
    profile: str | None = None,
) -> ListEventsResponse:
    """
    List events from a calendar.

    For recurring events use list_upcoming_events instead, which correctly
    expands recurring event instances via the calendarView endpoint.

    Args:
        max_results: Maximum number of events to return (1-100). Defaults to 10.
        filter_start: Optional ISO 8601 datetime to filter events starting after
                      this time. E.g. '2026-03-31T00:00:00'.
        calendar_id: Optional calendar ID. Defaults to the primary calendar.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured event summaries.
    """
    g = get_graph(profile)
    params: dict = {
        "$top": max_results,
        "$select": "id,subject,start,end,location,organizer,isAllDay,isCancelled,responseStatus",
        "$orderby": "start/dateTime",
    }
    if filter_start:
        # Sanitize: strip quotes to prevent OData injection
        safe_start = filter_start.replace("'", "")
        params["$filter"] = f"start/dateTime ge '{safe_start}'"

    if calendar_id:
        path = f"/me/calendars/{calendar_id}/events"
    else:
        path = "/me/calendar/events"

    result = await g.get(path, params=params)
    events = result.get("value", [])

    return ListEventsResponse(
        calendar_id=calendar_id,
        filter_start=filter_start,
        count=len(events),
        events=[_event_summary(ev) for ev in events],
    )


# ---------------------------------------------------------------------------
# list_upcoming_events (calendarView)
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def list_upcoming_events(
    start_datetime: str,
    end_datetime: str,
    max_results: int = 25,
    calendar_id: Optional[str] = None,
    profile: str | None = None,
) -> ListUpcomingEventsResponse:
    """
    List upcoming events using the calendarView endpoint, which correctly
    expands recurring event instances within the specified time window.

    Args:
        start_datetime: Start of the time window (ISO 8601, e.g. '2026-03-31T00:00:00').
        end_datetime: End of the time window (ISO 8601, e.g. '2026-04-07T23:59:59').
        max_results: Maximum number of events to return (1-100). Defaults to 25.
        calendar_id: Optional calendar ID. Defaults to the primary calendar.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured event summaries within the time window.
    """
    g = get_graph(profile)
    params: dict = {
        "startDateTime": start_datetime,
        "endDateTime": end_datetime,
        "$top": max_results,
        "$select": "id,subject,start,end,location,organizer,isAllDay,isCancelled,responseStatus",
        "$orderby": "start/dateTime",
    }

    if calendar_id:
        path = f"/me/calendars/{calendar_id}/calendarView"
    else:
        path = "/me/calendar/calendarView"

    result = await g.get(path, params=params)
    events = result.get("value", [])

    return ListUpcomingEventsResponse(
        calendar_id=calendar_id,
        start_datetime=start_datetime,
        end_datetime=end_datetime,
        count=len(events),
        events=[_event_summary(ev) for ev in events],
    )


# ---------------------------------------------------------------------------
# get_event
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def get_event(event_id: str, profile: str | None = None) -> EventDetailResponse:
    """
    Fetch a calendar event by ID with full details.

    Args:
        event_id: The Graph event ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured event details.
    """
    g = get_graph(profile)
    params = {
        "$select": (
            "id,subject,body,start,end,location,organizer,attendees,"
            "isAllDay,isCancelled,recurrence,onlineMeeting,webLink,"
            "responseStatus,importance,sensitivity,showAs"
        ),
    }

    ev = await g.get(f"/me/events/{event_id}", params=params)

    subject = ev.get("subject") or "(no subject)"
    start = _fmt_dt(ev.get("start", {}).get("dateTime"))
    start_tz = ev.get("start", {}).get("timeZone", "")
    end = _fmt_dt(ev.get("end", {}).get("dateTime"))
    location = ev.get("location", {}).get("displayName", "")
    organizer_ea = ev.get("organizer", {}).get("emailAddress", {})
    organizer = f"{organizer_ea.get('name', '')} <{organizer_ea.get('address', '')}>".strip()
    attendees_str = _fmt_attendees(ev.get("attendees", []))
    show_as = ev.get("showAs", "")
    web_link = ev.get("webLink", "")

    # Online meeting info
    online = ev.get("onlineMeeting") or {}
    join_url = online.get("joinUrl", "")

    # Body
    body_obj = ev.get("body", {})
    content_type = (body_obj.get("contentType") or "text").lower()
    raw_body = body_obj.get("content", "")
    if content_type == "html" and raw_body:
        import html as html_module
        import re
        text = re.sub(r"<(style|script)[^>]*>.*?</\1>", "", raw_body, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
        text = re.sub(r"<[^>]+>", "", text)
        body_text = html_module.unescape(text).strip()
    else:
        body_text = raw_body

    # Recurrence
    recurrence = ev.get("recurrence")
    recurrence_str = ""
    if recurrence:
        pattern = recurrence.get("pattern", {})
        recurrence_str = f"{pattern.get('type', '')} (interval: {pattern.get('interval', 1)})"

    return EventDetailResponse(
        id=event_id,
        subject=subject,
        start=ev.get("start", {}).get("dateTime"),
        start_display=start,
        end=ev.get("end", {}).get("dateTime"),
        end_display=end,
        timezone=start_tz,
        location=location,
        organizer=DisplayAddress(
            display=organizer,
            name=organizer_ea.get("name", ""),
            address=organizer_ea.get("address", ""),
        ),
        attendees=_attendee_values(ev.get("attendees", [])),
        attendees_display=attendees_str,
        is_all_day=ev.get("isAllDay", False),
        is_cancelled=ev.get("isCancelled", False),
        show_as=show_as,
        web_link=web_link,
        join_url=join_url,
        body=body_text,
        body_content_type=content_type,
        recurrence=recurrence,
        recurrence_display=recurrence_str,
        importance=ev.get("importance", ""),
        sensitivity=ev.get("sensitivity", ""),
    )


# ---------------------------------------------------------------------------
# create_event
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_WRITE)
async def create_event(
    subject: str,
    start_datetime: str,
    end_datetime: str,
    timezone: str = "UTC",
    body: Optional[str] = None,
    body_type: BodyType = "text",
    location: Optional[str] = None,
    attendees: Optional[Union[str, list[str]]] = None,
    optional_attendees: Optional[Union[str, list[str]]] = None,
    is_all_day: bool = False,
    is_online_meeting: bool = False,
    calendar_id: Optional[str] = None,
    profile: str | None = None,
) -> CreateEventResponse:
    """
    Create a new calendar event.

    Args:
        subject: Event title.
        start_datetime: Start time in ISO 8601 format (e.g. '2026-04-01T10:00:00').
        end_datetime: End time in ISO 8601 format (e.g. '2026-04-01T11:00:00').
        timezone: IANA timezone string (e.g. 'America/Sao_Paulo', 'UTC'). Defaults to 'UTC'.
        body: Optional event description/body.
        body_type: 'text' or 'html'. Defaults to 'text'.
        location: Optional location display name.
        attendees: Optional required attendee email(s). Comma-separated string or list.
        optional_attendees: Optional attendee email(s). Comma-separated string or list.
        is_all_day: When True, create an all-day event. Defaults to False.
        is_online_meeting: When True, create as Teams online meeting. Defaults to False.
        calendar_id: Optional calendar ID. Defaults to the primary calendar.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured event creation confirmation.
    """
    g = get_graph(profile)
    event: dict = {
        "subject": subject,
        "start": _build_datetime(start_datetime, timezone),
        "end": _build_datetime(end_datetime, timezone),
        "isAllDay": is_all_day,
    }

    if body:
        event["body"] = {
            "contentType": "HTML" if body_type.lower() == "html" else "Text",
            "content": body,
        }

    if location:
        event["location"] = {"displayName": location}

    all_attendees = []
    all_attendees.extend(_parse_attendees(attendees, "required"))
    all_attendees.extend(_parse_attendees(optional_attendees, "optional"))
    if all_attendees:
        event["attendees"] = all_attendees

    if is_online_meeting:
        event["isOnlineMeeting"] = True
        event["onlineMeetingProvider"] = "teamsForBusiness"

    if calendar_id:
        path = f"/me/calendars/{calendar_id}/events"
    else:
        path = "/me/calendar/events"

    result = await g.post(path, json=event)

    event_id = (result or {}).get("id", "unknown")
    web_link = (result or {}).get("webLink", "")
    join_url = ""
    online_meeting = (result or {}).get("onlineMeeting") or {}
    if online_meeting:
        join_url = online_meeting.get("joinUrl", "")

    return CreateEventResponse(
        success=True,
        action="create_event",
        event_id=event_id,
        subject=subject,
        start_datetime=start_datetime,
        end_datetime=end_datetime,
        timezone=timezone,
        calendar_id=calendar_id,
        web_link=web_link,
        join_url=join_url,
    )


# ---------------------------------------------------------------------------
# update_event
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_WRITE)
async def update_event(
    event_id: str,
    subject: Optional[str] = None,
    start_datetime: Optional[str] = None,
    end_datetime: Optional[str] = None,
    timezone: str = "UTC",
    body: Optional[str] = None,
    body_type: BodyType = "text",
    location: Optional[str] = None,
    attendees: Optional[Union[str, list[str]]] = None,
    is_all_day: Optional[bool] = None,
    profile: str | None = None,
) -> UpdateEventResponse:
    """
    Update an existing calendar event. Only provided fields are changed.

    Args:
        event_id: The Graph event ID to update.
        subject: Replace event title (omit to leave unchanged).
        start_datetime: Replace start time (ISO 8601, omit to leave unchanged).
        end_datetime: Replace end time (ISO 8601, omit to leave unchanged).
        timezone: IANA timezone for start/end times. Defaults to 'UTC'.
        body: Replace event body (omit to leave unchanged).
        body_type: 'text' or 'html'. Defaults to 'text'.
        location: Replace location (omit to leave unchanged).
        attendees: Replace attendee list (omit to leave unchanged).
        is_all_day: Set all-day flag (omit to leave unchanged).
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured event update confirmation.
    """
    g = get_graph(profile)
    patch: dict = {}

    if subject is not None:
        patch["subject"] = subject
    if start_datetime is not None:
        patch["start"] = _build_datetime(start_datetime, timezone)
    if end_datetime is not None:
        patch["end"] = _build_datetime(end_datetime, timezone)
    if body is not None:
        patch["body"] = {
            "contentType": "HTML" if body_type.lower() == "html" else "Text",
            "content": body,
        }
    if location is not None:
        patch["location"] = {"displayName": location}
    if attendees is not None:
        patch["attendees"] = _parse_attendees(attendees, "required")
    if is_all_day is not None:
        patch["isAllDay"] = is_all_day

    if not patch:
        return UpdateEventResponse(
            success=False,
            action="update_event",
            event_id=event_id,
            updated_fields=[],
            error="No fields to update.",
        )

    result = await g.patch(f"/me/events/{event_id}", json=patch)

    updated_id = (result or {}).get("id", event_id)
    updated_fields = ", ".join(patch.keys())
    return UpdateEventResponse(
        success=True,
        action="update_event",
        event_id=updated_id,
        updated_fields=list(patch.keys()),
        updated_fields_display=updated_fields,
    )


# ---------------------------------------------------------------------------
# delete_event
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_DESTRUCTIVE)
async def delete_event(event_id: str, profile: str | None = None) -> DeleteEventResponse:
    """
    Delete a calendar event.

    Args:
        event_id: The Graph event ID to delete.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured event deletion confirmation.
    """
    g = get_graph(profile)
    await g.delete(f"/me/events/{event_id}")
    return DeleteEventResponse(success=True, action="delete_event", event_id=event_id)


# ---------------------------------------------------------------------------
# rsvp_event
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_WRITE)
async def rsvp_event(
    event_id: str,
    response: RsvpResponse,
    comment: Optional[str] = None,
    send_response: bool = True,
    profile: str | None = None,
) -> RsvpEventResponse:
    """
    Respond to a calendar event invitation (accept, decline, or tentatively accept).

    Args:
        event_id: The Graph event ID to respond to.
        response: One of 'accept', 'decline', or 'tentativelyAccept'.
        comment: Optional comment to include with the response.
        send_response: When True (default), send the response to the organizer.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured RSVP confirmation.
    """
    g = get_graph(profile)
    payload: dict = {"sendResponse": send_response}
    if comment:
        payload["comment"] = comment

    await g.post(f"/me/events/{event_id}/{response}", json=payload)
    return RsvpEventResponse(
        success=True,
        action="rsvp_event",
        event_id=event_id,
        response=response,
        comment=comment,
        send_response=send_response,
    )


# ---------------------------------------------------------------------------
# get_free_busy
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def get_free_busy(
    email_addresses: Union[str, list[str]],
    start_datetime: str,
    end_datetime: str,
    timezone: str = "UTC",
    profile: str | None = None,
) -> FreeBusyResponse:
    """
    Check free/busy availability for one or more people.

    Args:
        email_addresses: Email address(es) to check. Comma-separated string or list.
        start_datetime: Start of the time window (ISO 8601).
        end_datetime: End of the time window (ISO 8601).
        timezone: IANA timezone. Defaults to 'UTC'.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured free/busy data for each requested person.
    """
    g = get_graph(profile)
    if isinstance(email_addresses, str):
        schedules = [addr.strip() for addr in email_addresses.split(",") if addr.strip()]
    else:
        schedules = [addr.strip() for addr in email_addresses if addr.strip()]

    payload = {
        "schedules": schedules,
        "startTime": _build_datetime(start_datetime, timezone),
        "endTime": _build_datetime(end_datetime, timezone),
        "availabilityViewInterval": 30,  # 30-minute slots
    }

    result = await g.post("/me/calendar/getSchedule", json=payload)
    schedule_items = (result or {}).get("value", [])

    people: list[PersonSchedule] = []
    for sched in schedule_items:
        email = sched.get("scheduleId", "unknown")
        availability = sched.get("availabilityView", "")
        items = sched.get("scheduleItems", [])
        people.append(
            PersonSchedule(
                email=email,
                availability_view=availability,
                legend={
                    "0": "Free",
                    "1": "Tentative",
                    "2": "Busy",
                    "3": "OOF",
                    "4": "Working Elsewhere",
                },
                schedule_items=[
                    ScheduleItemInfo(
                        status=item.get("status", ""),
                        start=item.get("start", {}).get("dateTime"),
                        start_display=_fmt_dt(item.get("start", {}).get("dateTime")),
                        end=item.get("end", {}).get("dateTime"),
                        end_display=_fmt_dt(item.get("end", {}).get("dateTime")),
                        subject=item.get("subject", ""),
                    )
                    for item in items
                ],
            )
        )

    return FreeBusyResponse(
        start_datetime=start_datetime,
        end_datetime=end_datetime,
        timezone=timezone,
        people=people,
    )


# ---------------------------------------------------------------------------
# find_meeting_times
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
async def find_meeting_times(
    attendees: Union[str, list[str]],
    duration_minutes: int = 60,
    start_datetime: Optional[str] = None,
    end_datetime: Optional[str] = None,
    timezone: str = "UTC",
    max_candidates: int = 5,
    is_organizer_optional: bool = False,
    profile: str | None = None,
) -> MeetingSuggestionsResponse:
    """
    Find available meeting time suggestions for a set of attendees.

    Uses the Microsoft Graph findMeetingTimes endpoint to suggest times
    when all (or most) attendees are available.

    Args:
        attendees: Attendee email address(es). Comma-separated string or list.
        duration_minutes: Meeting duration in minutes. Defaults to 60.
        start_datetime: Optional start of the search window (ISO 8601).
        end_datetime: Optional end of the search window (ISO 8601).
        timezone: IANA timezone. Defaults to 'UTC'.
        max_candidates: Maximum number of time suggestions to return. Defaults to 5.
        is_organizer_optional: When True, organizer's schedule is not considered. Defaults to False.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured meeting-time suggestions.
    """
    g = get_graph(profile)
    attendee_list = _parse_attendees(attendees, "required")

    payload: dict = {
        "attendees": attendee_list,
        "meetingDuration": f"PT{duration_minutes}M",
        "maxCandidates": max_candidates,
        "isOrganizerOptional": is_organizer_optional,
    }

    if start_datetime and end_datetime:
        payload["timeConstraint"] = {
            "activityDomain": "work",
            "timeSlots": [
                {
                    "start": _build_datetime(start_datetime, timezone),
                    "end": _build_datetime(end_datetime, timezone),
                }
            ],
        }

    result = await g.post("/me/findMeetingTimes", json=payload)

    suggestions = (result or {}).get("meetingTimeSuggestions", [])
    emptiness = (result or {}).get("emptySuggestionsReason", "")

    normalized: list[MeetingSuggestion] = []
    for sug in suggestions:
        confidence = sug.get("confidence", 0)
        meeting_ts = sug.get("meetingTimeSlot", {})
        normalized.append(
            MeetingSuggestion(
                confidence=confidence,
                start=meeting_ts.get("start", {}).get("dateTime"),
                start_display=_fmt_dt(meeting_ts.get("start", {}).get("dateTime")),
                end=meeting_ts.get("end", {}).get("dateTime"),
                end_display=_fmt_dt(meeting_ts.get("end", {}).get("dateTime")),
                attendee_availability=[
                    AttendeeAvailabilityInfo(
                        email=aa.get("attendee", {}).get("emailAddress", {}).get("address", ""),
                        availability=aa.get("availability", ""),
                    )
                    for aa in sug.get("attendeeAvailability", [])
                ],
            )
        )

    return MeetingSuggestionsResponse(
        count=len(normalized),
        suggestions=normalized,
        empty_suggestions_reason=emptiness or None,
        timezone=timezone,
    )
