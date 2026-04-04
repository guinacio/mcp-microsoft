from __future__ import annotations

import pytest

from mcp_microsoft.graph_types import GraphChatMessage, GraphEvent, GraphMessage
from mcp_microsoft.tools import calendar, mail, teams


def test_graph_message_accepts_null_wire_fields() -> None:
    message = GraphMessage.model_validate(
        {
            "id": "m1",
            "subject": None,
            "toRecipients": None,
            "ccRecipients": None,
            "bccRecipients": None,
            "body": None,
            "bodyPreview": None,
            "conversationId": None,
            "importance": None,
            "attachments": None,
        }
    )

    assert message.subject is None
    assert message.to_recipients == []
    assert message.cc_recipients == []
    assert message.bcc_recipients == []
    assert message.body.content == ""
    assert message.body_preview == ""
    assert message.conversation_id == ""
    assert message.importance == ""
    assert message.attachments == []


@pytest.mark.asyncio
async def test_get_event_handles_nullable_graph_fields(monkeypatch: pytest.MonkeyPatch) -> None:
    class DummyGraph:
        async def get(self, _path: str, params: dict | None = None):
            return {
                "id": "event-null",
                "subject": None,
                "body": None,
                "start": {"dateTime": "2026-04-11T10:00:00Z", "timeZone": None},
                "end": {"dateTime": "2026-04-11T11:00:00Z", "timeZone": None},
                "location": None,
                "organizer": None,
                "attendees": None,
                "isAllDay": False,
                "isCancelled": False,
                "responseStatus": None,
                "importance": None,
                "sensitivity": None,
                "showAs": None,
                "webLink": None,
                "recurrence": None,
                "onlineMeeting": None,
            }

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())

    result = await calendar.get_event(calendar.GetEventInput(event_id="event-null"))

    assert result.subject == "(no subject)"
    assert result.timezone == ""
    assert result.location == ""
    assert result.organizer.display == ""
    assert result.attendees == []
    assert result.attendees_display == "(none)"
    assert result.body == ""
    assert result.body_content_type == "text"
    assert result.importance == ""
    assert result.sensitivity == ""
    assert result.show_as == ""
    assert result.web_link == ""


@pytest.mark.asyncio
async def test_teams_channel_message_handles_nullable_graph_fields(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyGraph:
        async def get(self, _path: str, params: dict | None = None):
            return {
                "id": "msg-null",
                "createdDateTime": None,
                "lastModifiedDateTime": None,
                "from": None,
                "body": None,
                "subject": None,
                "webUrl": None,
                "replyToId": None,
                "importance": None,
                "reactions": None,
                "attachments": None,
                "mentions": None,
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())

    result = await teams.teams_get_channel_message(
        teams.TeamsGetChannelMessageInput(
            team_id="team-1",
            channel_id="channel-1",
            message_id="msg-null",
        )
    )

    assert result.id == "msg-null"
    assert result.body == ""
    assert result.body_content_type == "text"
    assert result.subject == ""
    assert result.web_url == ""
    assert result.reply_to_id == ""
    assert result.importance == ""
    assert result.reactions == []
    assert result.attachments == []
    assert result.mentions == []


@pytest.mark.asyncio
async def test_read_email_handles_nullable_graph_fields(monkeypatch: pytest.MonkeyPatch) -> None:
    class DummyGraph:
        async def get(self, _path: str, params: dict | None = None):
            return {
                "id": "mail-null",
                "subject": None,
                "from": None,
                "toRecipients": None,
                "ccRecipients": None,
                "receivedDateTime": None,
                "body": None,
                "bodyPreview": None,
                "attachments": None,
                "isRead": False,
                "conversationId": None,
                "importance": None,
            }

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())

    result = await mail.read_email(mail.ReadEmailInput(message_id="mail-null"))

    assert result.subject == "(no subject)"
    assert result.to == []
    assert result.cc == []
    assert result.conversation_id == ""
    assert result.importance == ""
    assert result.body == ""
    assert result.body_content_type == "text"
    assert result.attachments == []


def test_graph_chat_message_accepts_null_wire_fields() -> None:
    message = GraphChatMessage.model_validate(
        {
            "id": "chat-1",
            "createdDateTime": None,
            "lastModifiedDateTime": None,
            "body": None,
            "subject": None,
            "webUrl": None,
            "replyToId": None,
            "importance": None,
            "reactions": None,
            "attachments": None,
            "mentions": None,
        }
    )

    assert message.created_date_time == ""
    assert message.last_modified_date_time == ""
    assert message.body.content == ""
    assert message.subject == ""
    assert message.web_url == ""
    assert message.reply_to_id == ""
    assert message.importance == ""
    assert message.reactions == []
    assert message.attachments == []
    assert message.mentions == []


def test_graph_event_accepts_null_wire_fields() -> None:
    event = GraphEvent.model_validate(
        {
            "id": "event-1",
            "subject": None,
            "body": None,
            "location": None,
            "organizer": None,
            "attendees": None,
            "responseStatus": None,
            "importance": None,
            "sensitivity": None,
            "showAs": None,
            "webLink": None,
        }
    )

    assert event.subject == ""
    assert event.body.content == ""
    assert event.location.display_name == ""
    assert event.organizer.email_address.address == ""
    assert event.attendees == []
    assert event.response_status.response == ""
    assert event.importance == ""
    assert event.sensitivity == ""
    assert event.show_as == ""
    assert event.web_link == ""
