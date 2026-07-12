# Changelog

All notable changes to `mcp-microsoft` are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).
Versioning follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [0.8.0] — 2026-07-12

### Added

- **Multi-user Streamable HTTP server mode** (`MCP_TRANSPORT=http`), alongside the existing single-user `stdio` mode (default, unchanged). Any number of users connect concurrently over MCP Streamable HTTP (spec 2025-11-25, `/mcp`) and each authenticates with their own Microsoft work/school account via Microsoft Entra ID OAuth; every Graph call then runs under that user's delegated identity.
  - **Auth**: FastMCP's `AzureProvider` (OAuth-proxy pattern, bridging Entra's lack of Dynamic Client Registration for MCP clients) plus a per-user **On-Behalf-Of** token exchange for Microsoft Graph (`src/mcp_microsoft/identity.py`: `TokenProvider` protocol, `ProfileTokenProvider` for stdio, `OboTokenProvider` for http). `graph.get_graph()` is now transport-aware; the `profile` argument on every Graph tool is silently ignored in http mode — identity always comes from the caller's bearer token.
  - **New config surface** (`config.py`): `MCP_TRANSPORT`, `MCP_HTTP_HOST` (default `127.0.0.1`), `MCP_HTTP_PORT` (default `8000`), `MCP_HTTP_STATELESS` (default `false`), `MCP_BASE_URL`, `MCP_AUTH_CLIENT_ID`/`MCP_AUTH_CLIENT_SECRET`/`MCP_AUTH_TENANT_ID`, `MCP_AUTH_REQUIRED_SCOPE` (default `mcp-access`), `MCP_RATE_LIMIT_RPS` (default `10`). `validate_http_config()` fails startup fast and clearly when http mode is missing required auth config; stdio mode ignores all of it.
  - **Tool registration differences in http mode**: profile-management tools (`add_ms_profile`, `list_ms_profiles`, `remove_ms_profile`, `authenticate_ms_profile`, `set_default_ms_profile`) are not registered — identity management is a server-operator/stdio concern, not something a remote user should be able to do. Feature flags (`MCP_ENABLE_TEAMS`, `MCP_ENABLE_SHAREPOINT`, `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS`, `MCP_ENABLE_TEAMS_AI_INSIGHTS`) must be set explicitly in http mode — the corporate-account auto-detect fallback used by manual stdio installs doesn't apply (there's no single profile to inspect).
  - **Audience**: http mode targets a single work/school tenant. `MCP_AUTH_TENANT_ID` must be the concrete **tenant GUID** — pseudo-tenants (`organizations`, `common`, `consumers`) and verified domains (`contoso.onmicrosoft.com`) are rejected at startup, because fastmcp's `AzureProvider` pins the accepted token issuer to a literal URL built from this value and real Entra tokens carry the concrete GUID in `iss`, so only the GUID validates. Personal Microsoft accounts remain stdio-only — On-Behalf-Of and custom API scopes aren't reliably supported for consumer accounts. Multi-tenant deployments are future work (issuer-skip + per-tenant OBO authority).

### Security

- **Local-disk tool gating in http mode** (`tools/onedrive.py`, `tools/sharepoint.py`, `tools/teams.py`, `tools/attachments.py`, `tools/contacts.py`). The server's local disk is not the caller's disk in a multi-user deployment: `download_file`, `download_from_site`, and `teams_download_meeting_recording` are not registered at all in http mode, and `upload_file`/`upload_to_site` (`local_path`), `download_attachment` (`save_path`), and `get_contact_photo` (`save_path`) reject those parameters at call time with an explanatory error instead of touching the filesystem.
- **Rate limiting** (`middleware.py`: `UserRateLimitMiddleware`) — bounded per-user (`tid`+`oid`) token-bucket limit, `MCP_RATE_LIMIT_RPS` (default 10 req/s, burst 2×); set to `0` or negative to disable. Replaces fastmcp's `RateLimitingMiddleware`, which keyed every request under one shared `"global"` bucket (one user could throttle all) and backed its per-client limiters with a `defaultdict` that never evicted; ours keys per user and bounds memory (LRU cap of 10,000 keys + idle-TTL pruning). Client-visible behavior is unchanged (same token-bucket semantics; over-limit raises the same `RateLimitError`).
- **Audit logging** (`middleware.py`: `AuditLoggingMiddleware`) — one line per tool call in http mode: tool name, caller `oid` + `preferred_username` (from the validated bearer token's claims), duration, and success/error outcome. Arguments, results, and the token itself are never logged.
- **Error masking** — `mask_error_details=True` in http mode so internal exception details never leak to remote clients (stdio mode is unaffected).
- **Unauthenticated `GET /health`** (`server.py`) for load balancers/container healthchecks — returns `200 {"status": "ok", "transport": "http"}` without a bearer token, since it's mounted outside FastMCP's auth-wrapped MCP route.
- **Token-gated observability** (`metrics.py`, `middleware.py`, `server.py`) — DevOps traffic/usage metrics for http mode, **off by default** and enabled only when `MCP_STATS_TOKEN` is set. An in-process `MetricsRegistry` aggregates global totals, a rolling 60-minute per-minute traffic timeline, per-tool latency (p50/p95/avg over the last 256 calls), and per-user activity (capped at 1000 identities with least-recently-seen eviction). A `MetricsMiddleware` (registered after audit logging) feeds it without ever affecting the observed call. Three same-server routes require the token (Bearer or HTTP Basic password, timing-safe comparison, never logged): `GET /metrics` (Prometheus text exposition — `mcp_uptime_seconds`, `mcp_calls_total`, `mcp_errors_total`, `mcp_users_tracked`, `mcp_users_evicted_total`, and per-tool `mcp_tool_calls_total`/`mcp_tool_errors_total`/`mcp_tool_duration_ms`, with **no per-user label series** to avoid unbounded cardinality), `GET /stats` (JSON snapshot), and `GET /dashboard` (one self-contained HTML page, no external requests, polling `/stats` every 10s). Metrics are in-memory and reset on restart; no tool arguments, results, or tokens are ever recorded — only `oid`/`username`, the same exposure as the audit log. When `MCP_STATS_TOKEN` is unset the routes are not registered at all. stdio mode is entirely unaffected.
- The existing deletion kill-switch (`MCP_DISABLE_DELETION_TOOLS`) and Teams/SharePoint feature gating both continue to work identically in http mode.
- **Single-worker constraint, documented**: the OAuth-proxy client store and the per-user OBO credential cache both live in-process memory. Run exactly one worker in http mode; horizontal scaling needs fastmcp's external `client_storage` backend, which is not wired up here (documented as future work in README.md).

### Changed

- Dependency floor: `fastmcp[azure]>=3.4.4` (pulls in `azure-identity` for the OBO exchange).
- `docs/azure-setup.md` — added a new "App registration for the remote (http) server" walkthrough: confidential client, Web platform, redirect URI, Expose an API / `mcp-access` scope, `requestedAccessTokenVersion: 2`, client secret, delegated Graph permissions, admin consent. This is explicitly a **separate** App Registration from the stdio public-client one (different platform type).
- `README.md` — new "Remote server — multi-user (Streamable HTTP)" section: what it is, quickstart, how MCP clients connect, the full list of http-mode behavioral differences, and a security-notes subsection.
- Added `Dockerfile`, `docker-compose.yml`, and an expanded `.env.template` for running the http server (image runs as a non-root user; healthcheck against `/health`; compose file documents TLS termination via a reverse proxy — `MCP_BASE_URL` must equal the proxy's public HTTPS URL).

### Compatibility

- **stdio mode (MCPB / Claude Desktop / from-source) is unchanged.** No new required configuration, no behavior change, no new dependencies pulled onto the request path. `manifest.json` and the MCPB bundles remain stdio-only.

---

## [0.7.0] — 2026-05-26

### Security

- **Encrypted token cache** (`profiles.py`). MSAL token caches were previously written as plaintext JSON to `~/.microsoft-mcp/msal_cache_*.json`, exposing refresh tokens to any process running as the same user. Caches are now persisted via [msal-extensions](https://github.com/AzureAD/microsoft-authentication-extensions-for-python) using OS-native encryption: DPAPI on Windows, Keychain on macOS, libsecret on Linux. A permission-restricted plaintext fallback (mode `0600`) covers headless Linux without a keyring. Legacy `msal_cache_*.json` files are migrated to the encrypted `.bin` format on first run and the originals are deleted.
- **`profiles.json` permissions tightened** to `0600` on POSIX; the credentials directory to `0700`. Windows relies on the existing user-profile ACL.
- **`delete_email` confirm bug fix** (`tools/mail.py`). The pre-flight guard read `if params.confirm and ctx:`, silently skipping the elicitation prompt when an MCP host did not supply a `Context` and proceeding straight to permanent deletion. The tool now fails closed when `confirm=True` but `ctx` is unavailable.
- **OData injection fix in `search_contacts`** (`tools/contacts.py`). User input was previously sanitised by stripping single quotes — bypassable via OData operators (e.g. `') and (true`). Replaced with proper OData quote-doubling (`'` → `''`). `list_contacts` `$search` now also rejects control characters and caps input at 256 chars.

### Added

- **`MCP_DISABLE_DELETION_TOOLS` env var** (`config.py`, `feature_flags.py`). When truthy, the server omits registration of all permanent-delete tools: `delete_email`, `bulk_delete_emails`, `delete_event`, `delete_contact`, `delete_folder`, `delete_drive_item`, `delete_list_item`, `remove_ms_profile`. Recoverable variants (`trash_email`, `bulk_trash_emails`, `move_or_copy_item`) remain available. Surfaced as a "Disable Permanent-Delete Tools" toggle in the MCPB extension settings.
- **`mcp-microsoft-nodelete.mcpb`** companion bundle — identical capability surface but with `MCP_DISABLE_DELETION_TOOLS=true` hardcoded and the toggle hidden from extension settings. Installable side-by-side with the main bundle.

### Changed

- New runtime dependency: `msal-extensions>=1.2`.
- Token-cache file extension changed from `.json` to `.bin` to signal the new opaque encrypted format.

### Tests

- `tests/test_security_hardening.py` — confirm fail-closed path, deletion-tool gating in both directions, persisted-cache fallback writes through to disk.
- `tests/test_contacts.py` — updated to assert OData quote-doubling instead of quote-stripping; added a break-out attempt to verify the literal cannot be escaped.

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
