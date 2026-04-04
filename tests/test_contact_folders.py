"""
tests/test_contact_folders.py — Test coverage for the contact folder tools.

Uses monkeypatch to mock the Graph API client, following the
pattern established in test_calendar.py.
"""

from __future__ import annotations

import json

import pytest

from mcp_microsoft.models import (
    ContactFolderInfo,
    ListContactFoldersResponse,
)
import mcp_microsoft.server as server
from mcp_microsoft.tools import contacts


# ---------------------------------------------------------------------------
# Tool Registration Test
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_contact_folder_tool_is_registered() -> None:
    """Verify list_contact_folders tool is registered with the MCP server."""
    tool_names = {tool.name for tool in await server.mcp.list_tools(run_middleware=False)}
    assert "list_contact_folders" in tool_names


# ---------------------------------------------------------------------------
# list_contact_folders
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_list_contact_folders_returns_structured_response(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_contact_folders returns ListContactFoldersResponse with folder info."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None, headers: dict = None):
            return {
                "value": [
                    {
                        "id": "folder-1",
                        "displayName": "Family",
                        "parentFolderId": "root-folder",
                        "totalItemCount": 5,
                    },
                    {
                        "id": "folder-2",
                        "displayName": "Work Contacts",
                        "parentFolderId": "root-folder",
                        "totalItemCount": 42,
                    },
                ]
            }

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.list_contact_folders(contacts.ListContactFoldersInput())

    assert isinstance(result, ListContactFoldersResponse)
    assert result.count == 2
    assert len(result.folders) == 2

    family = next(f for f in result.folders if f.id == "folder-1")
    assert isinstance(family, ContactFolderInfo)
    assert family.display_name == "Family"
    assert family.parent_folder_id == "root-folder"
    assert family.total_item_count == 5

    work = next(f for f in result.folders if f.id == "folder-2")
    assert work.display_name == "Work Contacts"
    assert work.total_item_count == 42


@pytest.mark.asyncio
async def test_list_contact_folders_empty(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_contact_folders handles an empty folder list gracefully."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None, headers: dict = None):
            return {"value": []}

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.list_contact_folders(contacts.ListContactFoldersInput())

    assert isinstance(result, ListContactFoldersResponse)
    assert result.count == 0
    assert result.folders == []


@pytest.mark.asyncio
async def test_list_contact_folders_uses_correct_path(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_contact_folders requests /me/contactFolders."""
    captured: dict[str, str] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None, headers: dict = None):
            captured["path"] = path
            return {"value": []}

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    await contacts.list_contact_folders(contacts.ListContactFoldersInput())

    assert captured["path"] == "/me/contactFolders"


@pytest.mark.asyncio
async def test_list_contact_folders_selects_correct_fields(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_contact_folders sends the expected $select fields."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None, headers: dict = None):
            captured["params"] = params or {}
            return {"value": []}

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    await contacts.list_contact_folders(contacts.ListContactFoldersInput())

    select = captured["params"].get("$select", "")
    assert "id" in select
    assert "displayName" in select
    assert "totalItemCount" in select


@pytest.mark.asyncio
async def test_list_contact_folders_null_safe_fields(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_contact_folders handles None values in optional fields safely."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None, headers: dict = None):
            return {
                "value": [
                    {
                        "id": "folder-sparse",
                        "displayName": "Sparse",
                        "parentFolderId": None,
                        "totalItemCount": None,
                    }
                ]
            }

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.list_contact_folders(contacts.ListContactFoldersInput())

    folder = result.folders[0]
    assert folder.parent_folder_id == ""
    assert folder.total_item_count == 0


@pytest.mark.asyncio
async def test_list_contact_folders_single_folder(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_contact_folders works correctly with a single folder result."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None, headers: dict = None):
            return {
                "value": [
                    {
                        "id": "folder-only",
                        "displayName": "Favorites",
                        "parentFolderId": "parent-id",
                        "totalItemCount": 3,
                    }
                ]
            }

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    result = await contacts.list_contact_folders(contacts.ListContactFoldersInput())

    assert result.count == 1
    assert result.folders[0].id == "folder-only"
    assert result.folders[0].display_name == "Favorites"


@pytest.mark.asyncio
async def test_list_contact_folders_top_parameter(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify list_contact_folders sends $top=100 to limit results."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None, headers: dict = None):
            captured["params"] = params or {}
            return {"value": []}

    monkeypatch.setattr(contacts, "get_graph", lambda _profile: DummyGraph())
    await contacts.list_contact_folders(contacts.ListContactFoldersInput())

    assert captured["params"].get("$top") == 100


# ---------------------------------------------------------------------------
# ToolRequestModel input validation
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_input_camelcase_normalization() -> None:
    """Verify ListContactFoldersInput accepts camelCase keys (profile passthrough)."""
    params = contacts.ListContactFoldersInput.model_validate({"profile": None})
    assert params.profile is None


@pytest.mark.asyncio
async def test_input_stringified_json_payload() -> None:
    """Verify ToolRequestModel parses a JSON-encoded string as the full payload."""
    raw = json.dumps({})
    params = contacts.ListContactFoldersInput.model_validate(raw)
    assert params.profile is None


@pytest.mark.asyncio
async def test_list_contact_folders_profile_forwarded(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify the profile parameter is forwarded to get_graph."""
    captured_profile: list[str | None] = []

    def _get_graph(profile):
        captured_profile.append(profile)

        class DummyGraph:
            async def get(self, path: str, params: dict = None, headers: dict = None):
                return {"value": []}

        return DummyGraph()

    monkeypatch.setattr(contacts, "get_graph", _get_graph)
    await contacts.list_contact_folders(contacts.ListContactFoldersInput(profile="work"))

    assert captured_profile == ["work"]
