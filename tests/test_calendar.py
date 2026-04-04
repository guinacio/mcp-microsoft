"""
tests/test_calendar.py — Test coverage for the Calendar module.

Uses monkeypatch to mock the Graph API client, following the
pattern established in test_contacts.py.
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from mcp_microsoft.models import (
    CalendarInfo,
    CreateEventResponse,
    DeleteEventResponse,
    EventDetailResponse,
    FreeBusyResponse,
    ListCalendarsResponse,
    ListEventsResponse,
    UpdateEventResponse,
)
import mcp_microsoft.server as server
from mcp_microsoft.tools import calendar


# ---------------------------------------------------------------------------
# Tool Registration Test
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_calendar_tools_are_registered() -> None:
    """Verify all calendar tools are registered with the MCP server."""
    tool_names = {tool.name for tool in await server.mcp.list_tools(run_middleware=False)}
    assert "list_calendars" in tool_names
    assert "list_events" in tool_names
    assert "list_upcoming_events" in tool_names
    assert "get_event" in tool_names
    assert "create_event" in tool_names
    assert "update_event" in tool_names
    assert "delete_event" in tool_names
    assert "rsvp_event" in tool_names
    assert "get_free_busy" in tool_names
    assert "find_meeting_times" in tool_names


# ---------------------------------------------------------------------------
# list_events
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_list_events_returns_structured_response(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_events returns ListEventsResponse with normalized event summaries."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "value": [
                    {
                        "id": "event-1",
                        "subject": "Team Standup",
                        "start": {"dateTime": "2026-04-01T09:00:00", "timeZone": "UTC"},
                        "end": {"dateTime": "2026-04-01T09:30:00", "timeZone": "UTC"},
                        "location": {"displayName": "Conference Room A"},
                        "organizer": {"emailAddress": {"name": "Alice", "address": "alice@example.com"}},
                        "isAllDay": False,
                        "isCancelled": False,
                        "responseStatus": {"response": "accepted"},
                    }
                ]
            }

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    result = await calendar.list_events(calendar.ListEventsInput(max_results=10))

    assert isinstance(result, ListEventsResponse)
    assert result.count == 1
    assert len(result.events) == 1
    assert result.events[0].subject == "Team Standup"
    assert result.events[0].location == "Conference Room A"
    assert result.events[0].response_status == "accepted"


@pytest.mark.asyncio
async def test_list_events_empty_result(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_events handles an empty value array gracefully."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {"value": []}

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    result = await calendar.list_events(calendar.ListEventsInput())

    assert isinstance(result, ListEventsResponse)
    assert result.count == 0
    assert result.events == []


@pytest.mark.asyncio
async def test_list_events_uses_calendar_id_in_path(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_events constructs the correct path when calendar_id is provided."""
    captured: dict[str, str] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["path"] = path
            return {"value": []}

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    await calendar.list_events(calendar.ListEventsInput(calendar_id="cal-abc"))

    assert captured["path"] == "/me/calendars/cal-abc/events"


@pytest.mark.asyncio
async def test_list_events_default_path(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_events uses /me/calendar/events when no calendar_id given."""
    captured: dict[str, str] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["path"] = path
            return {"value": []}

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    await calendar.list_events(calendar.ListEventsInput())

    assert captured["path"] == "/me/calendar/events"


@pytest.mark.asyncio
async def test_list_events_filter_start_adds_odata_filter(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_events adds $filter param when filter_start is provided."""
    captured: dict[str, dict] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["params"] = params or {}
            return {"value": []}

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    await calendar.list_events(calendar.ListEventsInput(filter_start="2026-04-01T00:00:00"))

    assert "$filter" in captured["params"]
    assert "start/dateTime ge" in captured["params"]["$filter"]
    assert "2026-04-01T00:00:00" in captured["params"]["$filter"]


@pytest.mark.asyncio
async def test_list_events_datetime_display_formatting(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_events formats start_display/end_display from ISO datetimes."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "value": [
                    {
                        "id": "event-2",
                        "subject": "Lunch",
                        "start": {"dateTime": "2026-04-01T12:00:00", "timeZone": "UTC"},
                        "end": {"dateTime": "2026-04-01T13:00:00", "timeZone": "UTC"},
                        "location": {"displayName": ""},
                        "organizer": {"emailAddress": {"name": "", "address": ""}},
                        "isAllDay": False,
                        "isCancelled": False,
                        "responseStatus": {"response": "none"},
                    }
                ]
            }

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    result = await calendar.list_events(calendar.ListEventsInput())

    ev = result.events[0]
    assert ev.start_display == "2026-04-01 12:00"
    assert ev.end_display == "2026-04-01 13:00"


# ---------------------------------------------------------------------------
# get_event
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_get_event_returns_full_details(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify get_event returns EventDetailResponse with all major fields populated."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "id": "event-1",
                "subject": "Project Review",
                "body": {"contentType": "text", "content": "Quarterly review meeting."},
                "start": {"dateTime": "2026-04-10T14:00:00", "timeZone": "UTC"},
                "end": {"dateTime": "2026-04-10T15:00:00", "timeZone": "UTC"},
                "location": {"displayName": "Board Room"},
                "organizer": {"emailAddress": {"name": "Bob", "address": "bob@example.com", "display": "Bob"}},
                "attendees": [
                    {
                        "emailAddress": {"name": "Alice", "address": "alice@example.com"},
                        "type": "required",
                        "status": {"response": "accepted"},
                    }
                ],
                "isAllDay": False,
                "isCancelled": False,
                "responseStatus": {"response": "accepted"},
                "importance": "normal",
                "sensitivity": "normal",
                "showAs": "busy",
                "webLink": "https://outlook.office.com/calendar/item/event-1",
                "recurrence": None,
                "onlineMeeting": None,
            }

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    result = await calendar.get_event(calendar.GetEventInput(event_id="event-1"))

    assert isinstance(result, EventDetailResponse)
    assert result.id == "event-1"
    assert result.subject == "Project Review"
    assert result.location == "Board Room"
    assert result.body == "Quarterly review meeting."
    assert result.body_content_type == "text"
    assert result.is_all_day is False
    assert len(result.attendees) == 1
    assert result.attendees[0].address == "alice@example.com"
    assert result.attendees[0].response == "accepted"


@pytest.mark.asyncio
async def test_get_event_null_safe_optional_fields(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify get_event handles missing optional fields (onlineMeeting, recurrence) safely."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "id": "event-minimal",
                "subject": "Minimal Event",
                "body": {"contentType": "text", "content": ""},
                "start": {"dateTime": "2026-04-11T10:00:00", "timeZone": "UTC"},
                "end": {"dateTime": "2026-04-11T11:00:00", "timeZone": "UTC"},
                "location": {"displayName": ""},
                "organizer": {"emailAddress": {"name": "", "address": "", "display": ""}},
                "attendees": [],
                "isAllDay": False,
                "isCancelled": False,
                "responseStatus": {"response": "none"},
                "importance": "",
                "sensitivity": "",
                "showAs": "",
                "webLink": "",
                "recurrence": None,
                "onlineMeeting": None,
            }

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    result = await calendar.get_event(calendar.GetEventInput(event_id="event-minimal"))

    assert result.join_url == ""
    assert result.recurrence is None
    assert result.recurrence_display == ""
    assert result.attendees_display == "(none)"


# ---------------------------------------------------------------------------
# create_event
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_create_event_required_fields_only(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify create_event posts correct body and returns CreateEventResponse."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {
                "id": "new-event-id",
                "webLink": "https://outlook.office.com/calendar/item/new-event-id",
                "onlineMeeting": None,
            }

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    result = await calendar.create_event(
        calendar.CreateEventInput(
            subject="Kickoff Meeting",
            start_datetime="2026-04-15T10:00:00",
            end_datetime="2026-04-15T11:00:00",
        )
    )

    assert isinstance(result, CreateEventResponse)
    assert result.success is True
    assert result.event_id == "new-event-id"
    assert result.subject == "Kickoff Meeting"
    assert captured["path"] == "/me/calendar/events"
    body = captured["json"]
    assert body["subject"] == "Kickoff Meeting"
    assert body["start"]["dateTime"] == "2026-04-15T10:00:00"
    assert body["start"]["timeZone"] == "UTC"
    assert body["isAllDay"] is False
    assert "body" not in body
    assert "location" not in body
    assert "attendees" not in body


@pytest.mark.asyncio
async def test_create_event_with_optional_fields(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify create_event includes body, location, and attendees when provided."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["json"] = json
            return {"id": "full-event-id", "webLink": "", "onlineMeeting": None}

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    await calendar.create_event(
        calendar.CreateEventInput(
            subject="Full Event",
            start_datetime="2026-04-20T09:00:00",
            end_datetime="2026-04-20T10:00:00",
            body="Agenda: quarterly planning",
            body_type="text",
            location="Main Office",
            attendees=["alice@example.com", "bob@example.com"],
            optional_attendees="carol@example.com",
        )
    )

    body = captured["json"]
    assert body["body"]["content"] == "Agenda: quarterly planning"
    assert body["body"]["contentType"] == "Text"
    assert body["location"]["displayName"] == "Main Office"
    assert len(body["attendees"]) == 3
    required = [a for a in body["attendees"] if a["type"] == "required"]
    optional = [a for a in body["attendees"] if a["type"] == "optional"]
    assert len(required) == 2
    assert len(optional) == 1
    assert optional[0]["emailAddress"]["address"] == "carol@example.com"


@pytest.mark.asyncio
async def test_create_event_uses_calendar_id_in_path(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify create_event posts to calendar-specific path when calendar_id is given."""
    captured: dict[str, str] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            return {"id": "ev-1", "webLink": "", "onlineMeeting": None}

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    await calendar.create_event(
        calendar.CreateEventInput(
            subject="Cal-specific Event",
            start_datetime="2026-04-21T10:00:00",
            end_datetime="2026-04-21T11:00:00",
            calendar_id="cal-xyz",
        )
    )

    assert captured["path"] == "/me/calendars/cal-xyz/events"


@pytest.mark.asyncio
async def test_create_event_online_meeting_sets_teams_fields(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify is_online_meeting=True adds Teams meeting fields to the payload."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["json"] = json
            return {
                "id": "teams-event-id",
                "webLink": "",
                "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/join/abc"},
            }

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    result = await calendar.create_event(
        calendar.CreateEventInput(
            subject="Teams Meeting",
            start_datetime="2026-04-22T10:00:00",
            end_datetime="2026-04-22T11:00:00",
            is_online_meeting=True,
        )
    )

    body = captured["json"]
    assert body["isOnlineMeeting"] is True
    assert body["onlineMeetingProvider"] == "teamsForBusiness"
    assert result.join_url == "https://teams.microsoft.com/join/abc"


# ---------------------------------------------------------------------------
# update_event
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_update_event_patches_subset_of_fields(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify update_event sends PATCH with only the provided fields."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def patch(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {"id": "event-1"}

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    result = await calendar.update_event(
        calendar.UpdateEventInput(
            event_id="event-1",
            subject="Updated Title",
            location="New Location",
        )
    )

    assert isinstance(result, UpdateEventResponse)
    assert result.success is True
    assert set(result.updated_fields) == {"subject", "location"}
    assert captured["json"]["subject"] == "Updated Title"
    assert captured["json"]["location"]["displayName"] == "New Location"
    assert "start" not in captured["json"]
    assert "end" not in captured["json"]


@pytest.mark.asyncio
async def test_update_event_no_fields_returns_error(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify update_event returns an error response when no fields are provided."""
    monkeypatch.setattr(calendar, "get_graph", lambda _profile: None)
    result = await calendar.update_event(
        calendar.UpdateEventInput(event_id="event-1")
    )

    assert result.success is False
    assert result.error == "No fields to update."
    assert result.updated_fields == []


# ---------------------------------------------------------------------------
# delete_event
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_delete_event_sends_delete_to_correct_path(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify delete_event sends DELETE request and returns success response."""
    captured_path: dict[str, str] = {}

    class DummyGraph:
        async def delete(self, path: str):
            captured_path["path"] = path

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    result = await calendar.delete_event(calendar.DeleteEventInput(event_id="event-99"))

    assert isinstance(result, DeleteEventResponse)
    assert result.success is True
    assert result.event_id == "event-99"
    assert captured_path["path"] == "/me/events/event-99"


# ---------------------------------------------------------------------------
# list_calendars
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_list_calendars_returns_structured_response(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_calendars returns ListCalendarsResponse with normalized calendar info."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "value": [
                    {
                        "id": "cal-1",
                        "name": "My Calendar",
                        "color": "lightBlue",
                        "isDefaultCalendar": True,
                        "canEdit": True,
                    },
                    {
                        "id": "cal-2",
                        "name": "Birthdays",
                        "color": "auto",
                        "isDefaultCalendar": False,
                        "canEdit": False,
                    },
                ]
            }

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    result = await calendar.list_calendars(calendar.ListCalendarsInput())

    assert isinstance(result, ListCalendarsResponse)
    assert result.count == 2
    default_cal = next(c for c in result.calendars if c.is_default)
    assert default_cal.name == "My Calendar"
    assert default_cal.color == "lightBlue"
    readonly_cal = next(c for c in result.calendars if not c.can_edit)
    assert readonly_cal.name == "Birthdays"


# ---------------------------------------------------------------------------
# rsvp_event
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_rsvp_event_posts_comment_when_send_response_true(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {}

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())

    result = await calendar.rsvp_event(
        calendar.RsvpEventInput(
            event_id="event-1",
            response="accept",
            comment="See you there",
            send_response=True,
        )
    )

    assert result.success is True
    assert captured["path"] == "/me/events/event-1/accept"
    assert captured["json"] == {"sendResponse": True, "comment": "See you there"}


def test_rsvp_event_rejects_comment_when_send_response_false() -> None:
    with pytest.raises(ValidationError, match="comment requires send_response=true"):
        calendar.RsvpEventInput(
            event_id="event-1",
            response="accept",
            comment="Do not send this",
            send_response=False,
        )


# ---------------------------------------------------------------------------
# get_free_busy
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_get_free_busy_returns_schedule_for_multiple_users(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify get_free_busy returns FreeBusyResponse with PersonSchedule for each user."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            captured["payload"] = json
            return {
                "value": [
                    {
                        "scheduleId": "alice@example.com",
                        "availabilityView": "0020",
                        "scheduleItems": [
                            {
                                "status": "busy",
                                "start": {"dateTime": "2026-04-15T10:00:00", "timeZone": "UTC"},
                                "end": {"dateTime": "2026-04-15T10:30:00", "timeZone": "UTC"},
                                "subject": "Standup",
                            }
                        ],
                    },
                    {
                        "scheduleId": "bob@example.com",
                        "availabilityView": "0000",
                        "scheduleItems": [],
                    },
                ]
            }

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    result = await calendar.get_free_busy(
        calendar.GetFreeBusyInput(
            email_addresses=["alice@example.com", "bob@example.com"],
            start_datetime="2026-04-15T09:00:00",
            end_datetime="2026-04-15T17:00:00",
        )
    )

    assert isinstance(result, FreeBusyResponse)
    assert len(result.people) == 2
    assert result.start_datetime == "2026-04-15T09:00:00"
    assert result.end_datetime == "2026-04-15T17:00:00"
    assert result.timezone == "UTC"

    alice = next(p for p in result.people if p.email == "alice@example.com")
    assert alice.availability_view == "0020"
    assert len(alice.schedule_items) == 1
    assert alice.schedule_items[0].status == "busy"
    assert alice.schedule_items[0].subject == "Standup"
    assert alice.schedule_items[0].start_display == "2026-04-15 10:00"

    bob = next(p for p in result.people if p.email == "bob@example.com")
    assert bob.schedule_items == []

    assert captured["path"] == "/me/calendar/getSchedule"
    assert set(captured["payload"]["schedules"]) == {"alice@example.com", "bob@example.com"}


@pytest.mark.asyncio
async def test_get_free_busy_accepts_comma_separated_string(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify get_free_busy splits a comma-separated email string into schedule list."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["schedules"] = json["schedules"]
            return {"value": []}

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    await calendar.get_free_busy(
        calendar.GetFreeBusyInput(
            email_addresses="alice@example.com, bob@example.com",
            start_datetime="2026-04-15T09:00:00",
            end_datetime="2026-04-15T17:00:00",
        )
    )

    assert captured["schedules"] == ["alice@example.com", "bob@example.com"]


# ---------------------------------------------------------------------------
# ToolRequestModel input validation
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_input_camelcase_normalization() -> None:
    """Verify ToolRequestModel accepts camelCase keys and normalises to snake_case."""
    params = calendar.ListEventsInput.model_validate(
        {"maxResults": 25, "calendarId": "cal-123", "filterStart": "2026-04-01T00:00:00"}
    )
    assert params.max_results == 25
    assert params.calendar_id == "cal-123"
    assert params.filter_start == "2026-04-01T00:00:00"


@pytest.mark.asyncio
async def test_input_stringified_json_payload() -> None:
    """Verify ToolRequestModel parses a JSON-encoded string as the full payload."""
    import json

    raw = json.dumps({"subject": "Meeting", "startDatetime": "2026-04-01T10:00:00", "endDatetime": "2026-04-01T11:00:00"})
    params = calendar.CreateEventInput.model_validate(raw)
    assert params.subject == "Meeting"
    assert params.start_datetime == "2026-04-01T10:00:00"


@pytest.mark.asyncio
async def test_input_attendees_comma_separated_string() -> None:
    """Verify CreateEventInput coerces a comma-separated attendees string into a list."""
    params = calendar.CreateEventInput.model_validate(
        {
            "subject": "All Hands",
            "start_datetime": "2026-05-01T09:00:00",
            "end_datetime": "2026-05-01T10:00:00",
            "attendees": "alice@example.com, bob@example.com, carol@example.com",
        }
    )
    assert isinstance(params.attendees, list)
    assert len(params.attendees) == 3
    assert "alice@example.com" in params.attendees
