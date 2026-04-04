# Changelog

All notable changes to `mcp-microsoft` are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).
Versioning follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [0.6.0] — 2026-04-02

### Added

**Teams module** (`tools/teams.py`) — 18 new Microsoft Graph tools across 4 areas:
- Teams & Channels: `teams_list_joined`, `teams_get`, `teams_list_channels`, `teams_get_channel`, `teams_create_channel`
- Channel Messages: `teams_list_channel_messages`, `teams_get_channel_message`, `teams_send_channel_message`, `teams_reply_to_channel_message`, `teams_list_message_replies`
- Chats: `teams_list_chats`, `teams_get_chat`, `teams_list_chat_messages`, `teams_send_chat_message`, `teams_create_chat`
- Online Meetings: `teams_create_meeting`, `teams_get_meeting`, `teams_list_meetings`
- 6 new OAuth scopes in `profiles.py`: `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `Chat.ReadWrite`, `OnlineMeetings.ReadWrite`

**Mail Drafts** (`tools/drafts.py`) — 5 tools: `create_draft`, `list_drafts`, `get_draft`, `update_draft`, `send_draft`; all wrapped with `ToolRequestModel` input validation.

**Bulk mail operations** (`tools/mail.py`) — 3 tools: `bulk_move_emails`, `bulk_trash_emails`, `bulk_delete_emails`; execute up to 20 Graph API requests per `/$batch` call with per-item failure reporting.

**`filter_emails` tool** — OData `$filter` with full `$skiptoken` cursor-based pagination; replaces the prior `$skip` integer paging approach.

**SharePoint `search_content` tool** — full-text tenant-wide search via the Microsoft Search API (`POST /search/query`); supports KQL queries, multi-entity types (`driveItem`, `listItem`, `site`, `message`, `event`), optional single-site scoping, and result excerpts stripped of HTML.

**FastMCP server factory** (`config.py`, `server.py`) — `create_mcp_server()` factory function injects an `AppConfig` frozen dataclass; removes global state from module initialisation.

**Typed Graph response models** (`graph_types.py`) — 300+ line module of Pydantic models covering drive items, messages, events, channels, chats, and search hits; replaces raw `dict` access throughout all tool modules.

**Shared `common/` utilities**:
- `common/tooling.py` — `READ_ONLY_TOOL`, `WRITE_TOOL`, `DESTRUCTIVE_TOOL` annotation constants; `apply_tool_annotations()` post-registration pass that infers `readOnlyHint`/`destructiveHint`/`idempotentHint`/`openWorldHint` from tool-name prefix for all registered tools; auto-title generation via `_humanize_title()` (snake_case → Title Case with proper casing for MS, SharePoint, OneDrive).
- `common/formatting.py` — centralised `_fmt_dt`, `_fmt_size`, and other display helpers (deduplicated from four tool modules).
- `common/text.py` — shared text-truncation utilities.
- `common/transfer.py` — shared `_upload_large_file` and `_drive_item_payload` helpers.
- `common/mail_utils.py` — shared mail-formatting helpers.

**Corporate account gating** (`feature_flags.py`, `server.py`) — `is_corporate_account()` auto-detects work/school tenants; `_should_register_corporate_service()` in `server.py` unifies Teams and SharePoint gating — respects explicit `MCP_ENABLE_TEAMS` / `MCP_ENABLE_SHAREPOINT` env vars (MCPB toggle), falls back to corporate auto-detect when unset. `feature_flags.py` exposes `resolve_optional_service_enabled()` and a 3-state `env_flag()` helper (on / off / unset).

**`ToolRequestModel` base class** (`common/request_model.py`) — Pydantic base for all tool input models; adds camelCase normalisation, JSON-string coercion, comma-separated list parsing, and required-field enforcement used across mail, Teams, contacts, and drafts tools.

**`content_base64` upload fallback** (OneDrive, SharePoint) — agents in containers can pass file content as base64; decoded to a system temp file and uploaded, with cleanup in a `finally` block.

**Opt-in `confirm` guard for destructive mail ops** — `send_email` and `delete_email` accept `confirm: bool = False`; when `True`, calls `ctx.elicit()` to display a preview / confirmation prompt before executing.

**GitHub Actions CI** (`.github/workflows/ci.yml`) — runs the full test suite on every push to `master` and on pull requests.

**Test suite** — expanded from 39 tests to **137 tests across 9 test files**:
- `test_drafts.py` (18 tests) — `create_draft`, `list_drafts`, `get_draft`, `update_draft`, `send_draft`
- `test_teams.py` (47 tests) — all 18 Teams tools plus 6 `ToolRequestModel` validation tests
- `test_calendar.py` (22 tests) — all calendar tools plus `ToolRequestModel` validation
- `test_contact_folders.py` (11 tests) — `list_contact_folders` with null-safe field handling
- `test_request_model_and_teams_guards.py`, `test_tool_request_models.py` — cross-module input model tests
- `test_optional_services_and_upload_safety.py` — conditional service registration, upload path safety

### Notable details (Teams module)
- `teams_create_channel` has `confirm: bool = False` dry-run guard — no API call unless `confirm=True`
- `teams_create_chat` fetches caller's ID from `/me` to satisfy Graph API owner membership requirement
- List tools truncate message bodies at 500 chars to protect LLM context
- `teams_list_meetings` always applies a `startDateTime` date-range filter (Graph requires it); defaults to today ± 7 days. May 400 on tenants without OData `$filter` support on `/me/onlineMeetings` — documented in tool docstring
- No `retry_on_429` yet — Teams endpoints throttle at ~4 req/s; `GraphClient` raises immediately on 429

### Changed
- `server.py` — Teams module registered; tool count now **82** (was 64); all modules converted to explicit `register(server)` pattern removing import side-effect coupling; `create_mcp_server()` factory introduced.
- `graph.py` — `GraphClient.batch()` added: POST `/$batch`, auto-chunks into groups of 20, inherits existing retry/auth logic.
- `profiles.py` — `TEAMS_SCOPES` split out of `DEFAULT_SCOPES`; new `build_default_scopes()` only requests Teams/SharePoint OAuth scopes when those services are enabled at runtime (fixes scope-bloat on personal accounts).
- `manifest.json` — bumped to 0.6.0; `enable_teams` / `enable_sharepoint` config options changed from string to boolean so Claude Desktop renders on/off toggles; Teams tools and `teams` keyword added.
- `pyproject.toml` — bumped to 0.6.0.

### Fixed
- `apply_tool_annotations(mcp)` was imported but never called — annotations were silently unset for all tools; fixed by calling it after all `register()` calls complete.
- Tool display titles were missing (`title=None`) — `_humanize_title()` now auto-generates them.
- `ctx.elicit()` called with `schema=` instead of `response_type=` (FastMCP API mismatch) — corrected in `send_email` and `delete_email`.
- `filter_emails` used integer `$skip` pagination, which Graph truncates at low page counts — replaced with `$skiptoken` opaque cursor from `@odata.nextLink`.
- `search_content` used `contentSources` for site scoping (external-connector API) — replaced with `path:"<webUrl>"` KQL injection into the query string.
- Null-field crashes across all Graph API tools (20 bugs) — `.get()` guards added throughout calendar, drafts, mail, OneDrive, SharePoint, and Teams response parsing.
- `_fmt_size` `NameError` in `onedrive.py` — moved to shared `common/formatting.py`.
- Unterminated f-string literal in `mail.py` `send_email` preview block.
- `content_base64` upload wrote to a caller-controlled path — now writes to a `tempfile`-managed path with cleanup in `finally`.

---

## [0.5.0] — 2026-04-01

### Added
- **Contacts module** (`tools/contacts.py`) — 8 new Microsoft Graph tools:
  - `list_contacts` — list contacts with optional folder scope and OData `$select`
  - `get_contact` — retrieve a single contact by ID
  - `create_contact` — create a contact with name, email, phone, org, and notes fields
  - `update_contact` — patch any subset of contact fields (returns early if no fields provided)
  - `delete_contact` — delete a contact by ID
  - `list_contact_folders` — enumerate contact folders
  - `search_contacts` — OData `startswith()` search across displayName / email; single-quote-sanitised to prevent injection
  - `get_contact_photo` — fetch profile photo as base64; optional `save_to` disk path
- **Pydantic models** (`models/contacts.py`) — `ContactCreate`, `ContactUpdate`, `ContactFolder`, `ContactPhoto`
- **`Contacts.ReadWrite` scope** added to `profiles.py` for all profile types (personal, enterprise, delegated)
- **14 tests** (`tests/test_contacts.py`) — covers tool registration, API paths, PATCH semantics, no-fields early-return, OData injection prevention, base64 encoding, disk save

### Changed
- `server.py` — Contacts module registered; tool count now **64** (was 56)
- `README.md` — Contacts section added; multi-profile guide expanded

---

## [0.4.0] — 2026-03-31

### Added
- **Mail module** (`tools/mail.py`) — 14 tools: list/get/send/reply/forward/move/delete messages, manage folders, handle attachments, search mail
- **Calendar module** (`tools/calendar.py`) — 14 tools: list/get/create/update/delete events, manage calendars, find free slots, accept/decline invitations
- **OneDrive module** (`tools/onedrive.py`) — 14 tools: list/get/upload/download/move/copy/delete files and folders, manage sharing links
- **SharePoint module** (`tools/sharepoint.py`) — 14 tools: list sites/drives/items, get/upload/download files, manage permissions
- **Multi-profile authentication** (`profiles.py`, `graph.py`) — named profiles via `~/.mcp-microsoft/<profile>/` directories; `--profile` flag on CLI; `SENTINEL_PROFILE` env var support
- **MCPB packaging** — `manifest.json` + `mcp-microsoft.mcpb` Desktop Extension bundle for one-click Claude Desktop install
- **HTTP client connection pool** (`graph.py`) — shared `httpx.Client` per profile for persistent keep-alive across Graph API calls
- **`mcp-microsoft-setup` CLI** — interactive Azure App Registration wizard; MSAL device-code + client-credentials auth flows

### Changed
- Initial release — project renamed from `mcp-outlook` (personal-only prototype) to `mcp-microsoft` (multi-tenant, multi-module)

---

## [0.1.0] — 2026-03-30

### Added
- Initial prototype (`mcp-outlook`) — personal Outlook Mail read/send via Microsoft Graph, single-profile MSAL device-code flow
