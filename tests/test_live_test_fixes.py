"""Regressions for issues found while exercising the server on a live account.

- OneDrive listings must return a cursor that can be passed back.
- Free/busy must tell "free" apart from "no data for this address".
- Calendar tools must not silently treat naive datetimes as UTC.
- search_emails must return importance like list_emails does.
- get_event must not report plain-text bodies as HTML.
"""

from __future__ import annotations

from typing import Any

import pytest
from fastmcp.exceptions import ToolError

from mcp_microsoft.config import AppConfig
from mcp_microsoft.tools import calendar, mail, onedrive


class _CaptureGraph:
    def __init__(self, response: Any) -> None:
        self.response = response
        self.calls: list[dict[str, Any]] = []

    async def get(self, path: str, params: dict | None = None, headers: dict | None = None):
        self.calls.append({"path": path, "params": params or {}, "headers": headers or {}})
        return self.response

    async def post(self, path: str, json: dict | None = None):
        self.calls.append({"path": path, "json": json})
        return self.response


@pytest.mark.asyncio
async def test_list_drive_items_returns_and_accepts_skip_token(monkeypatch: pytest.MonkeyPatch) -> None:
    graph = _CaptureGraph(
        {
            "value": [{"id": "1", "name": "a.txt", "file": {}}],
            "@odata.nextLink": (
                "https://graph.microsoft.com/v1.0/me/drive/root/children"
                "?$top=1&%24skipToken=page-2-cursor"
            ),
        }
    )
    monkeypatch.setattr(onedrive, "get_graph", lambda _profile: graph)

    first = await onedrive.list_drive_items(onedrive.ListDriveItemsInput(max_results=1))
    assert first.has_more is True
    assert first.next_page_token == "page-2-cursor"

    graph.response = {"value": [{"id": "2", "name": "b.txt", "file": {}}]}
    second = await onedrive.list_drive_items(
        onedrive.ListDriveItemsInput(max_results=1, skip_token=first.next_page_token)
    )
    assert graph.calls[-1]["params"]["$skiptoken"] == "page-2-cursor"
    assert second.has_more is False
    assert second.next_page_token is None


@pytest.mark.asyncio
async def test_free_busy_marks_addresses_without_data(monkeypatch: pytest.MonkeyPatch) -> None:
    graph = _CaptureGraph(
        {
            "value": [
                {"scheduleId": "me@outlook.com", "availabilityView": "00", "scheduleItems": []},
                {
                    "scheduleId": "someone@gmail.com",
                    "availabilityView": "",
                    "error": {"message": "The specified object was not found.", "responseCode": "ErrorItemNotFound"},
                },
            ]
        }
    )
    monkeypatch.setattr(calendar, "get_graph", lambda _profile: graph)

    result = await calendar.get_free_busy(
        calendar.GetFreeBusyInput(
            email_addresses=["me@outlook.com", "someone@gmail.com", "missing@example.com"],
            start_datetime="2026-10-19T09:00:00",
            end_datetime="2026-10-19T10:00:00",
            timezone="America/Sao_Paulo",
        )
    )

    people = {person.email: person for person in result.people}
    assert people["me@outlook.com"].has_data is True
    assert people["me@outlook.com"].error == ""
    assert people["someone@gmail.com"].has_data is False
    assert "not found" in people["someone@gmail.com"].error
    assert people["missing@example.com"].has_data is False
    assert result.timezone == "America/Sao_Paulo"


@pytest.mark.asyncio
async def test_create_event_rejects_naive_datetime_without_timezone(monkeypatch: pytest.MonkeyPatch) -> None:
    graph = _CaptureGraph({"id": "evt"})
    monkeypatch.setattr(calendar, "get_graph", lambda _profile: graph)
    monkeypatch.setattr(calendar, "get_app_config", lambda: AppConfig())

    with pytest.raises(ToolError, match="timezone"):
        await calendar.create_event(
            calendar.CreateEventInput(
                subject="Lunch",
                start_datetime="2026-10-19T12:00:00",
                end_datetime="2026-10-19T13:00:00",
            )
        )
    assert graph.calls == []


@pytest.mark.asyncio
async def test_create_event_uses_server_default_timezone(monkeypatch: pytest.MonkeyPatch) -> None:
    graph = _CaptureGraph({"id": "evt"})
    monkeypatch.setattr(calendar, "get_graph", lambda _profile: graph)
    monkeypatch.setattr(
        calendar, "get_app_config", lambda: AppConfig(default_timezone="America/Sao_Paulo")
    )

    result = await calendar.create_event(
        calendar.CreateEventInput(
            subject="Lunch",
            start_datetime="2026-10-19T12:00:00",
            end_datetime="2026-10-19T13:00:00",
        )
    )

    assert graph.calls[0]["json"]["start"] == {
        "dateTime": "2026-10-19T12:00:00",
        "timeZone": "America/Sao_Paulo",
    }
    assert result.timezone == "America/Sao_Paulo"


@pytest.mark.asyncio
async def test_create_event_converts_offset_datetimes_to_utc(monkeypatch: pytest.MonkeyPatch) -> None:
    graph = _CaptureGraph({"id": "evt"})
    monkeypatch.setattr(calendar, "get_graph", lambda _profile: graph)
    monkeypatch.setattr(calendar, "get_app_config", lambda: AppConfig())

    result = await calendar.create_event(
        calendar.CreateEventInput(
            subject="Lunch",
            start_datetime="2026-10-19T12:00:00-03:00",
            end_datetime="2026-10-19T13:00:00-03:00",
        )
    )

    assert graph.calls[0]["json"]["start"] == {"dateTime": "2026-10-19T15:00:00", "timeZone": "UTC"}
    assert graph.calls[0]["json"]["end"] == {"dateTime": "2026-10-19T16:00:00", "timeZone": "UTC"}
    assert result.start_datetime == "2026-10-19T15:00:00"
    assert result.timezone == "UTC"


@pytest.mark.asyncio
async def test_get_event_asks_graph_for_text_body(monkeypatch: pytest.MonkeyPatch) -> None:
    graph = _CaptureGraph(
        {
            "id": "evt",
            "body": {"contentType": "text", "content": "Bring the slides."},
            "start": {"dateTime": "2026-10-19T12:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-10-19T13:00:00", "timeZone": "UTC"},
        }
    )
    monkeypatch.setattr(calendar, "get_graph", lambda _profile: graph)

    result = await calendar.get_event(calendar.GetEventInput(event_id="evt"))

    assert graph.calls[0]["headers"] == {"Prefer": 'outlook.body-content-type="text"'}
    assert result.body == "Bring the slides."
    assert result.body_content_type == "text"


@pytest.mark.asyncio
async def test_search_emails_selects_importance(monkeypatch: pytest.MonkeyPatch) -> None:
    graph = _CaptureGraph(
        {
            "value": [
                {
                    "id": "m1",
                    "subject": "Hi",
                    "from": {"emailAddress": {"name": "A", "address": "a@example.com"}},
                    "receivedDateTime": "2026-10-18T10:00:00Z",
                    "importance": "normal",
                }
            ]
        }
    )
    monkeypatch.setattr(mail, "get_graph", lambda _profile: graph)

    result = await mail.search_emails(mail.SearchEmailsInput(query="hello"))

    assert "importance" in graph.calls[0]["params"]["$select"].split(",")
    assert result.messages[0].importance == "normal"
