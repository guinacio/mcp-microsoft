from __future__ import annotations

import pytest
from pydantic import ValidationError

from mcp_microsoft.tools import calendar, mail, sharepoint, teams


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("search_input", "expected_graph_query"),
    [
        ("project update", '"project update"'),
        ('"project update"', '"project update"'),
        ('to:"user@example.com"', '"to:user@example.com"'),
        ('"to:user@example.com"', '"to:user@example.com"'),
        (
            '"Denmark" AND ("LMS" OR "deployment")',
            '"Denmark AND (LMS OR deployment)"',
        ),
        (
            '("Package" OR Platform OR service.example) AND (Region OR Org)',
            '"(Package OR Platform OR service.example) AND (Region OR Org)"',
        ),
        ('from:user@example.com AND status', '"from:user@example.com AND status"'),
    ],
)
async def test_search_emails_uses_graph_search_envelope(
    monkeypatch: pytest.MonkeyPatch,
    search_input: str,
    expected_graph_query: str,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            captured["path"] = path
            captured["params"] = params or {}
            return {"value": []}

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())
    await mail.search_emails(mail.SearchEmailsInput(query=search_input))

    assert captured["path"] == "/me/messages"
    assert captured["params"]["$search"] == expected_graph_query


@pytest.mark.parametrize(
    "search_input",
    [
        'body:"project update" AND status',
        'from:"unterminated',
        "   ",
        "subject:status\nOR body:update",
    ],
)
def test_search_emails_rejects_unsupported_nested_or_unbalanced_quotes(
    search_input: str,
) -> None:
    with pytest.raises(ValueError, match="quoted|unbalanced|printable"):
        mail._normalize_email_search_query(search_input)


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("search_input", "expected_filter"),
    [
        ('subject:"Quarterly report"', "contains(subject, 'Quarterly report')"),
        ('SUBJECT:"Manager\'s update"', "contains(subject, 'Manager''s update')"),
    ],
)
async def test_search_emails_uses_filter_for_standalone_subject_phrase(
    monkeypatch: pytest.MonkeyPatch,
    search_input: str,
    expected_filter: str,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            captured["path"] = path
            captured["params"] = params or {}
            return {"value": []}

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())
    await mail.search_emails(
        mail.SearchEmailsInput(query=search_input, folder="sentitems")
    )

    assert captured["path"] == "/me/mailFolders/sentitems/messages"
    assert captured["params"]["$filter"] == expected_filter
    assert "$search" not in captured["params"]


@pytest.mark.asyncio
async def test_mail_orderby_property_leads_filter(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: list[dict[str, object]] = []

    class DummyGraph:
        async def get(self, _path: str, params: dict | None = None):
            captured.append(params or {})
            return {"value": []}

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())
    await mail.list_emails(mail.ListEmailsInput(unread_only=True))
    await mail.filter_emails(mail.FilterEmailsInput(subject_contains="status"))

    for query in captured:
        assert query["$orderby"] == "receivedDateTime desc"
        assert str(query["$filter"]).startswith("receivedDateTime ge ")


@pytest.mark.asyncio
async def test_filter_emails_uses_supported_sender_filter(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, _path: str, params: dict | None = None):
            captured.update(params or {})
            return {"value": []}

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())
    await mail.filter_emails(
        mail.FilterEmailsInput(from_address="sender@example.com")
    )

    assert "from/emailAddress/address eq 'sender@example.com'" in str(
        captured["$filter"]
    )
    assert "toRecipients" not in str(captured["$filter"])


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("folder", "expected_path"),
    [
        (None, "/me/messages"),
        ("inbox", "/me/mailFolders/inbox/messages"),
    ],
)
async def test_filter_emails_defaults_to_mailbox_wide_scope(
    monkeypatch: pytest.MonkeyPatch,
    folder: str | None,
    expected_path: str,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            captured["path"] = path
            return {"value": []}

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())
    result = await mail.filter_emails(
        mail.FilterEmailsInput(folder=folder, subject_contains="status")
    )

    assert captured["path"] == expected_path
    assert result.folder == folder


@pytest.mark.asyncio
async def test_filter_emails_preserves_validated_graph_continuation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls: list[tuple[str, dict | None]] = []

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            calls.append((path, params))
            if len(calls) == 1:
                return {
                    "value": [{"id": "first"}],
                    "@odata.nextLink": (
                        "https://graph.microsoft.com/v1.0/me/messages?"
                        "%24filter=subject%20eq%20%27status%27&%24skip=2"
                    ),
                }
            return {"value": [{"id": "second"}]}

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())
    first = await mail.filter_emails(
        mail.FilterEmailsInput(subject_contains="status", max_results=1)
    )
    second = await mail.filter_emails(
        mail.FilterEmailsInput(
            subject_contains="status",
            max_results=1,
            skip_token=first.next_page_token,
        )
    )

    assert first.next_page_token and first.next_page_token.startswith("mf1.")
    assert first.has_more is True
    assert calls[1] == (
        "/me/messages?%24filter=subject%20eq%20%27status%27&%24skip=2",
        None,
    )
    assert [message.id for message in second.messages] == ["second"]
    assert second.has_more is False


@pytest.mark.asyncio
async def test_filter_emails_rejects_untrusted_or_mismatched_continuation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class UnsafeGraph:
        async def get(self, _path: str, params: dict | None = None):
            return {
                "value": [],
                "@odata.nextLink": "https://example.com/v1.0/me/messages?$skip=2",
            }

    monkeypatch.setattr(mail, "get_graph", lambda _profile: UnsafeGraph())
    with pytest.raises(ValueError, match="invalid mail continuation"):
        await mail.filter_emails(mail.FilterEmailsInput(subject_contains="status"))

    safe_link = "https://graph.microsoft.com/v1.0/me/messages?$skip=2"
    cursor = mail._encode_mail_filter_cursor(
        fingerprint=mail._mail_filter_fingerprint(
            mail.FilterEmailsInput(subject_contains="status")
        ),
        page_link=safe_link,
        expected_path="/me/messages",
    )
    with pytest.raises(ValueError, match="does not match"):
        await mail.filter_emails(
            mail.FilterEmailsInput(
                subject_contains="different",
                skip_token=cursor,
            )
        )

    with pytest.raises(ValueError, match="invalid or expired"):
        await mail.filter_emails(
            mail.FilterEmailsInput(
                subject_contains="status",
                skip_token="mf1.not-valid-base64",
            )
        )


def test_filter_emails_rejects_recipient_search_combined_with_odata_filter() -> None:
    with pytest.raises(ValidationError, match="cannot be combined"):
        mail.FilterEmailsInput(
            to_address="recipient@example.com",
            subject_contains="status",
        )


@pytest.mark.asyncio
async def test_filter_emails_recipient_search_uses_kql_and_local_sorting(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            captured["path"] = path
            captured["params"] = params or {}
            return {
                "value": [
                    {
                        "id": "newer",
                        "subject": "Newer",
                        "receivedDateTime": "2026-08-24T12:00:00Z",
                    },
                    {
                        "id": "older",
                        "subject": "Older",
                        "receivedDateTime": "2026-08-23T12:00:00Z",
                    },
                ]
            }

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())
    result = await mail.filter_emails(
        mail.FilterEmailsInput(
            to_address="recipient@example.com",
            folder="sentitems",
            max_results=100,
            sort_order="oldest",
        )
    )

    assert captured["path"] == "/me/mailFolders/sentitems/messages"
    query = captured["params"]
    assert query["$search"] == '"to:recipient@example.com"'
    assert query["$top"] == mail._RECIPIENT_SEARCH_MAX_RESULTS
    assert "$filter" not in query
    assert "$orderby" not in query
    assert [message.id for message in result.messages] == ["older", "newer"]
    assert result.has_more is False


@pytest.mark.asyncio
async def test_filter_emails_recipient_search_pages_with_bound_cursor(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyGraph:
        async def get(self, _path: str, params: dict | None = None):
            return {
                "value": [
                    {
                        "id": str(index),
                        "receivedDateTime": f"2026-08-{index:02d}T12:00:00Z",
                    }
                    for index in range(1, 6)
                ]
            }

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())
    first = await mail.filter_emails(
        mail.FilterEmailsInput(
            to_address="recipient@example.com",
            max_results=2,
            sort_order="oldest",
        )
    )
    second = await mail.filter_emails(
        mail.FilterEmailsInput(
            to_address="recipient@example.com",
            max_results=2,
            sort_order="oldest",
            skip_token=first.next_page_token,
        )
    )
    third = await mail.filter_emails(
        mail.FilterEmailsInput(
            to_address="recipient@example.com",
            max_results=2,
            sort_order="oldest",
            skip_token=second.next_page_token,
        )
    )

    assert [message.id for message in first.messages] == ["1", "2"]
    assert [message.id for message in second.messages] == ["3", "4"]
    assert [message.id for message in third.messages] == ["5"]
    assert first.next_page_token and first.next_page_token.startswith("m1.")
    assert second.has_more is True
    assert third.has_more is False

    with pytest.raises(ValueError, match="does not match"):
        await mail.filter_emails(
            mail.FilterEmailsInput(
                to_address="different@example.com",
                max_results=2,
                sort_order="oldest",
                skip_token=first.next_page_token,
            )
        )


@pytest.mark.asyncio
async def test_teams_list_chats_uses_supported_orderby(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            captured["path"] = path
            captured["params"] = params or {}
            return {"value": []}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_list_chats(teams.TeamsListChatsInput())

    assert captured["path"] == "/me/chats"
    assert captured["params"]["$orderby"] == "lastMessagePreview/createdDateTime desc"
    assert "$select" not in captured["params"]


@pytest.mark.asyncio
async def test_joined_teams_omits_unsupported_query_options(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            captured["path"] = path
            captured["params"] = params
            return {"value": []}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_list_joined(teams.TeamsListJoinedInput())

    assert captured == {"path": "/me/joinedTeams", "params": None}


@pytest.mark.asyncio
async def test_list_meetings_follows_calendar_pages_and_resolves_join_url(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls: list[tuple[str, dict | None]] = []

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            calls.append((path, params))
            if path == "/me/calendarView" and params and "$skiptoken" not in params:
                return {
                    "value": [{"id": "event-1", "isOnlineMeeting": False}],
                    "@odata.nextLink": (
                        "https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=next"
                    ),
                }
            if path == "/me/calendarView":
                return {
                    "value": [
                        {
                            "id": "event-2",
                            "isOnlineMeeting": True,
                            "onlineMeeting": {
                                "joinUrl": "https://teams.example/join/1"
                            },
                        }
                    ]
                }
            return {
                "value": [
                    {
                        "id": "meeting-1",
                        "subject": "Standup",
                        "joinWebUrl": "https://teams.example/join/1",
                        "startDateTime": "2026-04-01T09:00:00Z",
                        "endDateTime": "2026-04-01T09:15:00Z",
                    }
                ]
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_list_meetings(
        teams.TeamsListMeetingsInput(
            start_after="2026-04-01T00:00:00Z",
            start_before="2026-04-02T00:00:00Z",
        )
    )

    assert result.count == 1
    assert result.meetings[0].id == "meeting-1"
    assert calls[1] == ("/me/calendarView", {"$skiptoken": "next"})
    assert calls[2][1] == {"$filter": "JoinWebUrl eq 'https://teams.example/join/1'"}


@pytest.mark.asyncio
async def test_default_calendar_view_uses_documented_path(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            captured["path"] = path
            return {"value": []}

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    await calendar.list_upcoming_events(
        calendar.ListUpcomingEventsInput(
            start_datetime="2026-04-01T00:00:00Z",
            end_datetime="2026-04-02T00:00:00Z",
        )
    )

    assert captured["path"] == "/me/calendarView"


@pytest.mark.asyncio
async def test_search_api_isolates_restricted_entity_types(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    with pytest.raises(ValueError, match="searched separately"):
        await sharepoint.search_content(
            sharepoint.SearchContentInput(
                query="status", entity_types=["driveItem", "message"]
            )
        )

    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict | None = None):
            captured["request"] = (json or {})["requests"][0]
            return {"value": []}

    monkeypatch.setattr(
        sharepoint, "_get_sharepoint_graph", lambda _profile: DummyGraph()
    )
    await sharepoint.search_content(
        sharepoint.SearchContentInput(
            query="status", entity_types=["message"], max_results=100
        )
    )

    assert captured["request"]["size"] == 25
    assert "fields" not in captured["request"]
