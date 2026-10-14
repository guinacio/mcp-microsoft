"""Confirmation flows across MCP protocol eras.

MCP 2026-07-28 removed server-initiated requests, so ``confirm=True`` on a
modern connection must ask through an ``InputRequiredResult`` round trip
(SEP-2322) instead of ``ctx.elicit()``. These tests drive the real FastMCP
server through ``fastmcp.Client`` on both eras, and pin the fail-closed
binding rules of ``common.confirmation`` directly.
"""

from __future__ import annotations

from types import SimpleNamespace
from typing import Any

import pytest
from fastmcp import Client, FastMCP
from fastmcp.client.elicitation import ElicitResult
from mcp import types as mcp_types
from mcp_types.version import LATEST_HANDSHAKE_VERSION, LATEST_MODERN_VERSION

from mcp_microsoft.common import confirmation
from mcp_microsoft.common.confirmation import (
    CONFIRMATION_INPUT_KEY,
    ConfirmationDenied,
    ConfirmationGranted,
    request_confirmation,
)
from mcp_microsoft.runtime import reset_runtime_state
from mcp_microsoft.tools import mail

CONFIRM_TOOLS = {
    "send_email": "SendEmailResponse",
    "reply_email": "ReplyEmailResponse",
    "forward_email": "ForwardEmailResponse",
    "delete_email": "DeleteEmailResponse",
    "bulk_trash_emails": "BulkTrashEmailsResponse",
    "bulk_delete_emails": "BulkDeleteEmailsResponse",
}


@pytest.fixture(autouse=True)
def _stdio_runtime(monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
    monkeypatch.setenv("MCP_TRANSPORT", "stdio")
    monkeypatch.setenv("MS365_CREDENTIALS_DIR", str(tmp_path))
    monkeypatch.delenv("MCP_DISABLE_DELETION_TOOLS", raising=False)
    reset_runtime_state()
    yield
    reset_runtime_state()


class _SendGraph:
    def __init__(self) -> None:
        self.posts: list[str] = []

    async def post(self, path: str, json: dict | None = None):
        self.posts.append(path)
        return None


class _ShrinkingFolderGraph:
    """Folder that loses a message between the asking and answering rounds."""

    def __init__(self) -> None:
        self.listings = 0
        self.batched: list[str] = []

    async def get(self, path: str, params: dict | None = None):
        self.listings += 1
        ids = ["m1", "m2", "m3"] if self.listings == 1 else ["m1", "m2"]
        return {"value": [{"id": mid} for mid in ids]}

    async def batch(self, requests: list[dict]):
        self.batched.extend(request["url"] for request in requests)
        return [{"id": request["id"], "status": 204} for request in requests]


def _server() -> FastMCP:
    mcp = FastMCP("confirm-test")
    mail.register(mcp)
    return mcp


def _answer(confirmed: bool, prompts: list[str]):
    """Elicitation handler that answers on either era's schema shape."""

    async def handler(message, response_type, params, ctx):
        prompts.append(message)
        properties = params.requested_schema.get("properties", {})
        field = "confirmed" if "confirmed" in properties else "value"
        return ElicitResult(action="accept", content={field: confirmed})

    return handler


@pytest.mark.asyncio
async def test_confirm_tools_keep_their_declared_output_schema() -> None:
    mcp = _server()
    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}

    for name, model in CONFIRM_TOOLS.items():
        schema = tools[name].output_schema
        assert schema is not None, name
        rendered = repr(schema)
        assert "InputRequiredResult" not in rendered, name
        assert "input_required" not in rendered, name
        assert model in rendered or "success" in schema.get("properties", {}), name


@pytest.mark.asyncio
@pytest.mark.parametrize("mode", ["auto", "legacy"])
async def test_send_email_confirmation_accept_sends_once(
    monkeypatch: pytest.MonkeyPatch, mode: str
) -> None:
    graph = _SendGraph()
    monkeypatch.setattr(mail, "get_graph", lambda _profile: graph)
    prompts: list[str] = []

    async with Client(
        _server(), mode=mode, elicitation_handler=_answer(True, prompts)
    ) as client:
        result = await client.call_tool(
            "send_email",
            {"params": {"to": "a@example.com", "subject": "Hi", "body": "Body", "confirm": True}},
        )

    assert result.structured_content["success"] is True
    assert graph.posts == ["/me/sendMail"]
    assert len(prompts) == 1
    assert prompts[0].startswith("Send this email?")


@pytest.mark.asyncio
@pytest.mark.parametrize("mode", ["auto", "legacy"])
async def test_send_email_confirmation_decline_never_sends(
    monkeypatch: pytest.MonkeyPatch, mode: str
) -> None:
    graph = _SendGraph()
    monkeypatch.setattr(mail, "get_graph", lambda _profile: graph)

    async with Client(
        _server(), mode=mode, elicitation_handler=_answer(False, [])
    ) as client:
        result = await client.call_tool(
            "send_email",
            {"params": {"to": "a@example.com", "subject": "Hi", "body": "Body", "confirm": True}},
        )

    assert result.structured_content["success"] is False
    assert result.structured_content["error"] == "Cancelled by user."
    assert graph.posts == []


@pytest.mark.asyncio
async def test_modern_client_without_elicitation_fails_closed(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    graph = _SendGraph()
    monkeypatch.setattr(mail, "get_graph", lambda _profile: graph)

    async with Client(_server()) as client:
        result = await client.call_tool(
            "delete_email", {"params": {"message_id": "m1", "confirm": True}}
        )

    assert result.structured_content["success"] is False
    assert "negotiated form elicitation" in result.structured_content["error"]
    assert graph.posts == []


@pytest.mark.asyncio
async def test_modern_bulk_delete_reasks_when_folder_changes_between_rounds(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    graph = _ShrinkingFolderGraph()
    monkeypatch.setattr(mail, "get_graph", lambda _profile: graph)
    prompts: list[str] = []

    async with Client(_server(), elicitation_handler=_answer(True, prompts)) as client:
        result = await client.call_tool(
            "bulk_delete_emails",
            {"params": {"folder": "junkemail", "confirm": True}},
        )

    # The first approval was for 3 messages; the folder then held 2, so the
    # user was asked again and only the 2 current messages were deleted.
    assert [p.split(" messages")[0] for p in prompts] == [
        "Permanently delete 3",
        "Permanently delete 2",
    ]
    assert result.structured_content["success"] is True
    assert result.structured_content["succeeded"] == 2
    assert len(graph.batched) == 2


# ---------------------------------------------------------------------------
# request_confirmation unit rules (modern era)
# ---------------------------------------------------------------------------


def _modern_ctx(
    *,
    responses: dict[str, Any] | None = None,
    request_state: str | None = None,
    elicitation: bool = True,
) -> SimpleNamespace:
    capabilities = SimpleNamespace(
        elicitation=SimpleNamespace(form=SimpleNamespace(), url=None)
        if elicitation
        else None
    )
    return SimpleNamespace(
        session=SimpleNamespace(
            client_capabilities=capabilities,
            protocol_version=LATEST_MODERN_VERSION,
        ),
        input_responses=responses,
        request_state=request_state,
    )


def _accepted(confirmed: object = True) -> mcp_types.ElicitResult:
    return mcp_types.ElicitResult(action="accept", content={"confirmed": confirmed})


@pytest.mark.asyncio
async def test_first_modern_round_asks_with_bound_state() -> None:
    outcome = await request_confirmation(_modern_ctx(), message="Go?", binding="b1")

    assert isinstance(outcome, mcp_types.InputRequiredResult)
    request = outcome.input_requests[CONFIRMATION_INPUT_KEY]
    assert request.params.message == "Go?"
    assert outcome.request_state == confirmation._binding_state("b1")
    assert "b1" not in outcome.request_state  # only a digest leaves the server


@pytest.mark.asyncio
async def test_answer_without_minted_state_is_not_trusted() -> None:
    ctx = _modern_ctx(responses={CONFIRMATION_INPUT_KEY: _accepted()})

    outcome = await request_confirmation(ctx, message="Go?", binding="b1")

    assert isinstance(outcome, mcp_types.InputRequiredResult)


@pytest.mark.asyncio
async def test_answer_for_a_different_action_is_not_trusted() -> None:
    ctx = _modern_ctx(
        responses={CONFIRMATION_INPUT_KEY: _accepted()},
        request_state=confirmation._binding_state("other-action"),
    )

    outcome = await request_confirmation(ctx, message="Go?", binding="b1")

    assert isinstance(outcome, mcp_types.InputRequiredResult)


@pytest.mark.asyncio
async def test_matching_accepted_answer_is_granted() -> None:
    ctx = _modern_ctx(
        responses={CONFIRMATION_INPUT_KEY: _accepted()},
        request_state=confirmation._binding_state("b1"),
    )

    outcome = await request_confirmation(ctx, message="Go?", binding="b1")

    assert isinstance(outcome, ConfirmationGranted)


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "answer",
    [
        _accepted(False),
        _accepted("true"),
        mcp_types.ElicitResult(action="decline"),
        mcp_types.ElicitResult(action="cancel"),
        mcp_types.ListRootsResult(roots=[]),
    ],
)
async def test_anything_but_explicit_true_is_denied(answer: object) -> None:
    ctx = _modern_ctx(
        responses={CONFIRMATION_INPUT_KEY: answer},
        request_state=confirmation._binding_state("b1"),
    )

    outcome = await request_confirmation(ctx, message="Go?", binding="b1")

    assert isinstance(outcome, ConfirmationDenied)


@pytest.mark.asyncio
async def test_no_declared_elicitation_is_denied_on_both_eras() -> None:
    modern = _modern_ctx(elicitation=False)
    handshake = _modern_ctx(elicitation=False)
    handshake.session.protocol_version = LATEST_HANDSHAKE_VERSION

    for ctx in (modern, handshake, None):
        outcome = await request_confirmation(ctx, message="Go?", binding="b1")
        assert isinstance(outcome, ConfirmationDenied)
