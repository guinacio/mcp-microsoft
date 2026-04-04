from __future__ import annotations

import pytest

from mcp_microsoft.models import DriveItemDetailResponse, ListFoldersResponse, SiteFileDetailResponse
from mcp_microsoft.tools import folders, onedrive, sharepoint


@pytest.mark.asyncio
async def test_onedrive_get_drive_item_uses_typed_drive_model(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyGraph:
        async def get(self, _path: str, params: dict | None = None):
            return {
                "id": "item-1",
                "name": "report.docx",
                "size": 2048,
                "createdDateTime": "2026-04-01T10:00:00Z",
                "lastModifiedDateTime": "2026-04-02T11:30:00Z",
                "webUrl": "https://example.invalid/report.docx",
                "parentReference": {"path": "/drive/root:/Reports"},
                "createdBy": {"user": {"displayName": "Alice"}},
                "lastModifiedBy": {"application": {"displayName": "Office"}},
                "file": {"mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
            }

    monkeypatch.setattr(onedrive, "get_graph", lambda _profile: DummyGraph())

    result = await onedrive.get_drive_item(onedrive.GetDriveItemInput(item_id="item-1"))

    assert isinstance(result, DriveItemDetailResponse)
    assert result.id == "item-1"
    assert result.created_by == "Alice"
    assert result.modified_by == "Office"
    assert result.path == "/drive/root:/Reports"
    assert result.mime_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


@pytest.mark.asyncio
async def test_sharepoint_get_site_file_uses_typed_drive_model(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyGraph:
        async def get(self, _path: str, params: dict | None = None):
            return {
                "id": "item-1",
                "name": "budget.xlsx",
                "size": 4096,
                "createdDateTime": "2026-04-01T10:00:00Z",
                "lastModifiedDateTime": "2026-04-02T11:30:00Z",
                "webUrl": "https://example.invalid/budget.xlsx",
                "parentReference": {"path": "/drives/drive-1/root:/Finance"},
                "createdBy": {"user": {"displayName": "Bob"}},
                "lastModifiedBy": {"user": {"displayName": "Carol"}},
                "file": {"mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
            }

    monkeypatch.setattr(sharepoint, "_get_sharepoint_graph", lambda _profile: DummyGraph())

    result = await sharepoint.get_site_file(
        sharepoint.GetSiteFileInput(site_id="site-1", drive_id="drive-1", item_id="item-1")
    )

    assert isinstance(result, SiteFileDetailResponse)
    assert result.id == "item-1"
    assert result.created_by == "Bob"
    assert result.modified_by == "Carol"
    assert result.path == "/drives/drive-1/root:/Finance"
    assert result.mime_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@pytest.mark.asyncio
async def test_list_folders_handles_nullable_counts_with_typed_folder_model(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            if path.endswith("/childFolders"):
                return {
                    "value": [
                        {
                            "id": "child-1",
                            "displayName": "Subfolder",
                            "totalItemCount": None,
                            "unreadItemCount": None,
                            "childFolderCount": None,
                        }
                    ]
                }
            return {
                "value": [
                    {
                        "id": "parent-1",
                        "displayName": "Projects",
                        "totalItemCount": None,
                        "unreadItemCount": None,
                        "childFolderCount": 1,
                    }
                ]
            }

    monkeypatch.setattr(folders, "get_graph", lambda _profile: DummyGraph())

    result = await folders.list_folders(
        folders.ListFoldersInput(include_child_folders=True)
    )

    assert isinstance(result, ListFoldersResponse)
    assert result.count == 2
    parent = next(folder for folder in result.folders if folder.id == "parent-1")
    child = next(folder for folder in result.folders if folder.id == "child-1")
    assert parent.total_count == 0
    assert child.unread_count == 0
    assert child.display_label.startswith("  ")
    assert child.is_child is True
