# Changelog

All notable changes to `mcp-microsoft` are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).
Versioning follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
