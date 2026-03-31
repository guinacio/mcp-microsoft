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
from typing import Optional, Union

from mcp_microsoft.graph import get_graph
from mcp_microsoft.server import mcp

# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------


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


# ---------------------------------------------------------------------------
# list_calendars
# ---------------------------------------------------------------------------


@mcp.tool()
async def list_calendars(profile: str | None = None) -> str:
    """
    List all calendars in the user's mailbox.

    Args:
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted table of calendars with name, color, and ID.
    """
    g = get_graph(profile)
    params = {
        "$select": "id,name,color,isDefaultCalendar,canEdit",
        "$top": 50,
    }

    result = await g.get("/me/calendars", params=params)
    calendars = result.get("value", [])

    if not calendars:
        return "No calendars found."

    lines = ["## Calendars\n"]
    lines.append("| Calendar | Color | Default | Editable | ID |")
    lines.append("|---|---|---|---|---|")

    for cal in calendars:
        name = cal.get("name", "(unnamed)")
        color = cal.get("color", "auto")
        is_default = "Yes" if cal.get("isDefaultCalendar") else "No"
        can_edit = "Yes" if cal.get("canEdit") else "No"
        cal_id = cal.get("id", "")
        lines.append(f"| {name} | {color} | {is_default} | {can_edit} | `{cal_id}` |")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# list_events
# ---------------------------------------------------------------------------


@mcp.tool()
async def list_events(
    max_results: int = 10,
    filter_start: Optional[str] = None,
    calendar_id: Optional[str] = None,
    profile: str | None = None,
) -> str:
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
        Markdown-formatted list of event summaries.
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

    if not events:
        return "No events found."

    lines = [f"## Events ({len(events)} found)\n"]
    for ev in events:
        subject = ev.get("subject") or "(no subject)"
        start = _fmt_dt(ev.get("start", {}).get("dateTime"))
        end = _fmt_dt(ev.get("end", {}).get("dateTime"))
        location = ev.get("location", {}).get("displayName", "")
        all_day = " [All Day]" if ev.get("isAllDay") else ""
        cancelled = " [CANCELLED]" if ev.get("isCancelled") else ""
        loc_str = f" | Location: {location}" if location else ""
        lines.append(
            f"- **{subject}**{all_day}{cancelled}\n"
            f"  {start} — {end}{loc_str}\n"
            f"  ID: `{ev.get('id')}`\n"
        )

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# list_upcoming_events (calendarView)
# ---------------------------------------------------------------------------


@mcp.tool()
async def list_upcoming_events(
    start_datetime: str,
    end_datetime: str,
    max_results: int = 25,
    calendar_id: Optional[str] = None,
    profile: str | None = None,
) -> str:
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
        Markdown-formatted list of event summaries within the time window.
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

    if not events:
        return f"No events found between {start_datetime} and {end_datetime}."

    lines = [f"## Upcoming Events ({len(events)} found)\n"]
    for ev in events:
        subject = ev.get("subject") or "(no subject)"
        start = _fmt_dt(ev.get("start", {}).get("dateTime"))
        end = _fmt_dt(ev.get("end", {}).get("dateTime"))
        location = ev.get("location", {}).get("displayName", "")
        all_day = " [All Day]" if ev.get("isAllDay") else ""
        cancelled = " [CANCELLED]" if ev.get("isCancelled") else ""
        loc_str = f" | Location: {location}" if location else ""
        lines.append(
            f"- **{subject}**{all_day}{cancelled}\n"
            f"  {start} — {end}{loc_str}\n"
            f"  ID: `{ev.get('id')}`\n"
        )

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# get_event
# ---------------------------------------------------------------------------


@mcp.tool()
async def get_event(event_id: str, profile: str | None = None) -> str:
    """
    Fetch a calendar event by ID with full details.

    Args:
        event_id: The Graph event ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted event details including attendees, body, and location.
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
    all_day = "Yes" if ev.get("isAllDay") else "No"
    cancelled = "Yes" if ev.get("isCancelled") else "No"
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

    lines = [
        f"## {subject}",
        f"**When:** {start} — {end} ({start_tz})",
        f"**All Day:** {all_day}",
    ]
    if location:
        lines.append(f"**Location:** {location}")
    lines.append(f"**Organizer:** {organizer}")
    lines.append(f"**Attendees:** {attendees_str}")
    lines.append(f"**Show As:** {show_as}")
    if cancelled == "Yes":
        lines.append("**Status:** CANCELLED")
    if recurrence_str:
        lines.append(f"**Recurrence:** {recurrence_str}")
    if join_url:
        lines.append(f"**Join URL:** {join_url}")
    if web_link:
        lines.append(f"**Web Link:** {web_link}")
    lines.append(f"**Event ID:** `{event_id}`")

    if body_text:
        lines += ["", "---", "", body_text]

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# create_event
# ---------------------------------------------------------------------------


@mcp.tool()
async def create_event(
    subject: str,
    start_datetime: str,
    end_datetime: str,
    timezone: str = "UTC",
    body: Optional[str] = None,
    body_type: str = "text",
    location: Optional[str] = None,
    attendees: Optional[Union[str, list[str]]] = None,
    optional_attendees: Optional[Union[str, list[str]]] = None,
    is_all_day: bool = False,
    is_online_meeting: bool = False,
    calendar_id: Optional[str] = None,
    profile: str | None = None,
) -> str:
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
        Confirmation string with the new event ID.
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

    lines = [
        f"Event created successfully.",
        f"**Subject:** {subject}",
        f"**When:** {start_datetime} — {end_datetime} ({timezone})",
        f"**Event ID:** `{event_id}`",
    ]
    if web_link:
        lines.append(f"**Web Link:** {web_link}")
    if join_url:
        lines.append(f"**Join URL:** {join_url}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# update_event
# ---------------------------------------------------------------------------


@mcp.tool()
async def update_event(
    event_id: str,
    subject: Optional[str] = None,
    start_datetime: Optional[str] = None,
    end_datetime: Optional[str] = None,
    timezone: str = "UTC",
    body: Optional[str] = None,
    body_type: str = "text",
    location: Optional[str] = None,
    attendees: Optional[Union[str, list[str]]] = None,
    is_all_day: Optional[bool] = None,
    profile: str | None = None,
) -> str:
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
        Confirmation string with updated event ID.
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
        return "No fields to update — provide at least one field to change."

    result = await g.patch(f"/me/events/{event_id}", json=patch)

    updated_id = (result or {}).get("id", event_id)
    updated_fields = ", ".join(patch.keys())
    return (
        f"Event updated successfully.\n"
        f"**Event ID:** `{updated_id}`\n"
        f"**Updated fields:** {updated_fields}"
    )


# ---------------------------------------------------------------------------
# delete_event
# ---------------------------------------------------------------------------


@mcp.tool()
async def delete_event(event_id: str, profile: str | None = None) -> str:
    """
    Delete a calendar event.

    Args:
        event_id: The Graph event ID to delete.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Confirmation string.
    """
    g = get_graph(profile)
    await g.delete(f"/me/events/{event_id}")
    return f"Event `{event_id}` deleted successfully."


# ---------------------------------------------------------------------------
# rsvp_event
# ---------------------------------------------------------------------------


@mcp.tool()
async def rsvp_event(
    event_id: str,
    response: str,
    comment: Optional[str] = None,
    send_response: bool = True,
    profile: str | None = None,
) -> str:
    """
    Respond to a calendar event invitation (accept, decline, or tentatively accept).

    Args:
        event_id: The Graph event ID to respond to.
        response: One of 'accept', 'decline', or 'tentativelyAccept'.
        comment: Optional comment to include with the response.
        send_response: When True (default), send the response to the organizer.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Confirmation string.
    """
    g = get_graph(profile)
    valid_responses = {"accept", "decline", "tentativelyAccept"}
    if response not in valid_responses:
        return f"Invalid response '{response}'. Must be one of: {', '.join(sorted(valid_responses))}"

    payload: dict = {"sendResponse": send_response}
    if comment:
        payload["comment"] = comment

    await g.post(f"/me/events/{event_id}/{response}", json=payload)
    return f"Event `{event_id}` — response: **{response}** sent successfully."


# ---------------------------------------------------------------------------
# get_free_busy
# ---------------------------------------------------------------------------


@mcp.tool()
async def get_free_busy(
    email_addresses: Union[str, list[str]],
    start_datetime: str,
    end_datetime: str,
    timezone: str = "UTC",
    profile: str | None = None,
) -> str:
    """
    Check free/busy availability for one or more people.

    Args:
        email_addresses: Email address(es) to check. Comma-separated string or list.
        start_datetime: Start of the time window (ISO 8601).
        end_datetime: End of the time window (ISO 8601).
        timezone: IANA timezone. Defaults to 'UTC'.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Markdown-formatted schedule availability for each person.
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

    if not schedule_items:
        return "No schedule data returned."

    lines = ["## Free/Busy Schedule\n"]
    for sched in schedule_items:
        email = sched.get("scheduleId", "unknown")
        availability = sched.get("availabilityView", "")
        items = sched.get("scheduleItems", [])

        lines.append(f"### {email}")
        if availability:
            # Availability view: 0=free, 1=tentative, 2=busy, 3=oof, 4=working elsewhere
            legend = {"0": "Free", "1": "Tentative", "2": "Busy", "3": "OOF", "4": "Working Elsewhere"}
            lines.append(f"Availability slots (30 min each): `{availability}`")
            lines.append(f"Legend: {' | '.join(f'{k}={v}' for k, v in legend.items())}")

        if items:
            for item in items:
                status = item.get("status", "")
                start = _fmt_dt(item.get("start", {}).get("dateTime"))
                end = _fmt_dt(item.get("end", {}).get("dateTime"))
                subj = item.get("subject", "")
                subj_str = f" — {subj}" if subj else ""
                lines.append(f"  - [{status}] {start} — {end}{subj_str}")
        lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# find_meeting_times
# ---------------------------------------------------------------------------


@mcp.tool()
async def find_meeting_times(
    attendees: Union[str, list[str]],
    duration_minutes: int = 60,
    start_datetime: Optional[str] = None,
    end_datetime: Optional[str] = None,
    timezone: str = "UTC",
    max_candidates: int = 5,
    is_organizer_optional: bool = False,
    profile: str | None = None,
) -> str:
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
        Markdown-formatted list of suggested meeting times with attendee availability.
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

    if not suggestions:
        reason = f" Reason: {emptiness}" if emptiness else ""
        return f"No meeting time suggestions found.{reason}"

    lines = [f"## Meeting Time Suggestions ({len(suggestions)} found)\n"]
    for i, sug in enumerate(suggestions, 1):
        confidence = sug.get("confidence", 0)
        confidence_pct = f"{confidence * 100:.0f}%" if isinstance(confidence, float) else str(confidence)
        meeting_ts = sug.get("meetingTimeSlot", {})
        start = _fmt_dt(meeting_ts.get("start", {}).get("dateTime"))
        end = _fmt_dt(meeting_ts.get("end", {}).get("dateTime"))

        att_avail = sug.get("attendeeAvailability", [])
        avail_parts = []
        for aa in att_avail:
            aa_email = aa.get("attendee", {}).get("emailAddress", {}).get("address", "")
            aa_status = aa.get("availability", "")
            avail_parts.append(f"{aa_email}: {aa_status}")

        lines.append(
            f"### Option {i} (confidence: {confidence_pct})\n"
            f"  **Time:** {start} — {end}\n"
        )
        if avail_parts:
            lines.append("  **Attendee availability:**")
            for part in avail_parts:
                lines.append(f"  - {part}")
        lines.append("")

    return "\n".join(lines)
