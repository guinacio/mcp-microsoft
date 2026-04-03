from __future__ import annotations

import base64
import importlib
from pathlib import Path

import pytest
from fastmcp.utilities.types import File

from mcp_microsoft.models import (
    CreateListItemResponse,
    SharePointFields,
    UpdateListItemResponse,
    UploadFileResponse,
    UploadSiteFileResponse,
)
import mcp_microsoft.server as server
from mcp_microsoft.tools import attachments, onedrive, sharepoint


@pytest.mark.asyncio
async def test_sharepoint_tools_are_registered_when_enabled(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("MCP_ENABLE_SHAREPOINT", "true")
    importlib.reload(server)
    tool_names = {tool.name for tool in await server.mcp.list_tools(run_middleware=False)}
    assert "search_sharepoint_sites" in tool_names
    monkeypatch.delenv("MCP_ENABLE_SHAREPOINT", raising=False)
    importlib.reload(server)


@pytest.mark.asyncio
async def test_download_attachment_returns_fastmcp_file(monkeypatch: pytest.MonkeyPatch) -> None:
    class DummyGraph:
        async def get(self, _path: str):
            return {
                "name": "hello.txt",
                "contentType": "text/plain",
                "contentBytes": base64.b64encode(b"hello world").decode("ascii"),
            }

    monkeypatch.setattr(attachments, "get_graph", lambda _profile: DummyGraph())

    result = await attachments.download_attachment("message-id", "attachment-id")

    assert isinstance(result, File)
    assert result.data == b"hello world"
    assert result._name == "hello.txt"


@pytest.mark.asyncio
async def test_onedrive_large_upload_does_not_read_entire_file(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    large_file = tmp_path / "large.bin"
    large_file.write_bytes(b"x" * (4 * 1024 * 1024 + 1))

    class DummyGraph:
        async def post(self, _path: str, json: dict | None = None):
            return {"uploadUrl": "https://example.invalid/upload"}

    captured: dict[str, object] = {}

    async def fake_large_upload(upload_url: str, file_path: Path, total_size: int, ctx=None):
        captured["upload_url"] = upload_url
        captured["file_path"] = file_path
        captured["total_size"] = total_size
        return {"id": "drive-item", "webUrl": "https://example.invalid/file"}

    monkeypatch.setattr(onedrive, "get_graph", lambda _profile: DummyGraph())
    monkeypatch.setattr(onedrive, "_upload_large_file", fake_large_upload)
    monkeypatch.setattr(Path, "read_bytes", lambda self: (_ for _ in ()).throw(AssertionError("read_bytes should not be used for large uploads")))

    result = await onedrive.upload_file(large_file)

    assert isinstance(result, UploadFileResponse)
    assert result.success is True
    assert result.action == "upload_file"
    assert result.file_id == "drive-item"
    assert captured["upload_url"] == "https://example.invalid/upload"
    assert captured["file_path"] == large_file
    assert captured["total_size"] == large_file.stat().st_size


@pytest.mark.asyncio
async def test_sharepoint_large_upload_does_not_read_entire_file(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    large_file = tmp_path / "large.bin"
    large_file.write_bytes(b"x" * (4 * 1024 * 1024 + 1))

    class DummyGraph:
        async def post(self, _path: str, json: dict | None = None):
            return {"uploadUrl": "https://example.invalid/upload"}

    captured: dict[str, object] = {}

    async def fake_large_upload(upload_url: str, file_path: Path, total_size: int, ctx=None):
        captured["upload_url"] = upload_url
        captured["file_path"] = file_path
        captured["total_size"] = total_size
        return {"id": "drive-item", "webUrl": "https://example.invalid/file"}

    monkeypatch.setattr(sharepoint, "_get_sharepoint_graph", lambda _profile: DummyGraph())
    monkeypatch.setattr(sharepoint, "_upload_large_file", fake_large_upload)
    monkeypatch.setattr(Path, "read_bytes", lambda self: (_ for _ in ()).throw(AssertionError("read_bytes should not be used for large uploads")))

    result = await sharepoint.upload_to_site("site-id", "drive-id", large_file)

    assert isinstance(result, UploadSiteFileResponse)
    assert result.success is True
    assert result.action == "upload_to_site"
    assert result.file_id == "drive-item"
    assert captured["upload_url"] == "https://example.invalid/upload"
    assert captured["file_path"] == large_file
    assert captured["total_size"] == large_file.stat().st_size


@pytest.mark.asyncio
async def test_sharepoint_list_item_tools_accept_structured_fields(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[dict] = []

    class DummyGraph:
        async def post(self, _path: str, json: dict | None = None):
            calls.append({"method": "post", "json": json})
            return {"id": "item-123"}

        async def patch(self, _path: str, json: dict | None = None):
            calls.append({"method": "patch", "json": json})
            return {}

    monkeypatch.setattr(sharepoint, "_get_sharepoint_graph", lambda _profile: DummyGraph())

    create_result = await sharepoint.create_list_item("site-id", "list-id", SharePointFields({"Title": "Hello"}))
    update_result = await sharepoint.update_list_item("site-id", "list-id", "item-123", SharePointFields({"Status": "Done"}))

    assert isinstance(create_result, CreateListItemResponse)
    assert create_result.success is True
    assert create_result.fields == {"Title": "Hello"}
    assert isinstance(update_result, UpdateListItemResponse)
    assert update_result.success is True
    assert update_result.updated_fields == ["Status"]
    assert calls == [
        {"method": "post", "json": {"fields": {"Title": "Hello"}}},
        {"method": "patch", "json": {"Status": "Done"}},
    ]
