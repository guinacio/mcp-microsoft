# mcp-microsoft

![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)
![License MIT](https://img.shields.io/badge/license-MIT-green)
![MCP](https://img.shields.io/badge/MCP-compatible-purple)

Microsoft 365 MCP server — Mail, Calendar, OneDrive, and SharePoint via the Microsoft Graph API, with multi-account support.

## Overview

`mcp-microsoft` is a [Model Context Protocol](https://modelcontextprotocol.io) server that gives Claude (and any other MCP client) full access to your Microsoft 365 account. It covers the four most-used surface areas of the Microsoft Graph API: email, calendar, OneDrive file storage, and SharePoint — 56 tools in total.

The server works with both personal Microsoft accounts (Outlook.com, Live) and enterprise accounts (Azure AD / Entra ID) using a single App Registration. SharePoint tools are included automatically for work accounts and excluded for personal-only tenants, since the `Sites.ReadWrite.All` scope is unavailable to consumer accounts.

Multi-account support is a first-class feature. Named profiles let you configure separate client IDs for each account and switch between them on any tool call by passing `profile="work"`. Profiles and MSAL token caches are stored in `~/.microsoft-mcp/` and survive server restarts without re-authentication.

The server ships as an MCPB bundle (`mcp-microsoft.mcpb`) for zero-friction installation through the Claude Desktop Extension installer. It can also be run from source or wired directly into `claude_desktop_config.json`. Built with [FastMCP](https://github.com/jlowin/fastmcp), MSAL, and async httpx.

## Features

### Tools (56 total)

#### Mail (21 tools)

- `list_emails` — list messages from any folder with pagination and unread filter
- `read_email` — fetch the full body of a message by ID (supports summary mode)
- `search_emails` — search using Microsoft Graph KQL `$search` syntax
- `send_email` — compose and send a new message (to/cc/bcc, HTML or plain text)
- `reply_email` — reply or reply-all to an existing message
- `forward_email` — forward a message to one or more recipients
- `mark_as_read` / `mark_as_unread` — toggle read state
- `move_email` — move to any folder by well-known name or folder ID
- `trash_email` — soft-delete to Deleted Items (recoverable)
- `delete_email` — permanently delete a message (irreversible)
- `create_draft` / `get_draft` / `list_drafts` / `update_draft` / `send_draft` — full draft lifecycle
- `list_folders` / `create_folder` / `delete_folder` — manage mailbox folders
- `list_attachments` / `download_attachment` — inspect and save attachments

#### Calendar (10 tools)

- `list_calendars` — enumerate all calendars in the mailbox
- `list_events` — list events from a calendar with optional date filtering
- `list_upcoming_events` — list events using calendarView with recurring-instance expansion
- `get_event` — fetch full event details including attendees, body, and recurrence
- `create_event` — create an event (subject, datetime, timezone, attendees, location, online meeting flag)
- `update_event` / `delete_event` — modify or remove an event
- `rsvp_event` — accept, tentatively accept, or decline an invitation
- `get_free_busy` — check availability for one or more people in a time window
- `find_meeting_times` — get meeting time suggestions for a set of attendees

#### OneDrive (8 tools)

- `list_drive_items` — browse files and folders by path or item ID
- `get_drive_item` — get metadata for a specific file or folder
- `search_drive` — full-text search across OneDrive
- `upload_file` — upload a local file (auto-switches to resumable upload for files over 4 MB)
- `download_file` — download a file to a local path
- `create_drive_folder` — create a new folder at any path
- `move_or_copy_item` — move or copy items within OneDrive
- `delete_drive_item` — delete a file or folder (moves to recycle bin)

#### SharePoint (12 tools)

> SharePoint tools require a work or school account (Azure AD / Entra ID). They are not available for personal Outlook.com / Live accounts, which do not support the `Sites.ReadWrite.All` Graph permission. `Sites.ReadWrite.All` requires one-time admin consent in enterprise tenants.

- `search_sharepoint_sites` — search or list SharePoint sites the user can access
- `get_sharepoint_site` — get details of a specific site
- `list_site_libraries` — list document libraries in a site
- `list_site_files` / `get_site_file` — browse files in a document library
- `upload_to_site` / `download_from_site` — transfer files to/from SharePoint
- `list_site_lists` — list all SharePoint lists in a site
- `get_list_items` / `create_list_item` / `update_list_item` / `delete_list_item` — manage list records

#### Profile Management (5 tools)

- `list_ms_profiles` — list all configured profiles and which is the default
- `add_ms_profile` — add a new account (name, client_id, tenant_id)
- `remove_ms_profile` — remove a profile and delete its cached tokens
- `authenticate_ms_profile` — trigger interactive OAuth for a profile
- `set_default_ms_profile` — change which profile is used when none is specified

## Installation

### Option A: Claude Desktop Extension (MCPB) — Recommended

```bash
npx @anthropic-ai/mcpb install mcp-microsoft-0.4.0.mcpb
```

The installer prompts for your Azure App Registration details (see [Azure Setup](#azure-setup)):

| Prompt | Description |
|---|---|
| **Azure Client ID** | Application (client) ID from your App Registration |
| **Tenant ID** | `common` for personal + work, `consumers` for personal only, or your org's tenant ID/domain |
| **Credentials Directory** | Optional. Defaults to `~/.microsoft-mcp/` |

A `default` profile is created automatically from these values.

### Option B: From Source

```bash
git clone https://github.com/guilhermeinacio/mcp-microsoft.git
cd mcp-microsoft
uv sync
export MS365_CLIENT_ID=your-client-id
export MS365_TENANT_ID=common
uv run mcp-microsoft
```

### Option C: Add to claude_desktop_config.json

```json
{
  "mcpServers": {
    "mcp-microsoft": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/mcp-microsoft", "mcp-microsoft"],
      "env": {
        "MS365_CLIENT_ID": "your-client-id",
        "MS365_TENANT_ID": "common"
      }
    }
  }
}
```

## Azure Setup

You need an Azure App Registration to get a `client_id`. This is a one-time step.

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**.
2. Name it anything (e.g., `mcp-microsoft`).
3. Under **Supported account types**, choose based on your use case:
   - *Personal Microsoft accounts only* — Outlook.com / Live users
   - *Accounts in any organizational directory and personal Microsoft accounts* — personal and work
4. Under **Redirect URI**, select **Mobile and desktop applications** and enter `http://localhost`.
5. Under **Authentication**, enable **Allow public client flows** (required for the interactive loopback OAuth flow — no client secret needed).
6. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** and add:
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Calendars.ReadWrite`
   - `Files.ReadWrite`
   - `Sites.ReadWrite.All` *(work accounts only — required for SharePoint)*
   - `offline_access` *(usually pre-added)*
7. For `Sites.ReadWrite.All`: click **Grant admin consent**. Your IT administrator must approve this once per tenant.
8. From the **Overview** page, copy the **Application (client) ID** and, if targeting a specific tenant, the **Directory (tenant) ID**.

For a detailed walkthrough with screenshots, see [`docs/azure-setup.md`](docs/azure-setup.md).

## Profile Management

The server supports multiple Microsoft 365 accounts as named profiles. Each profile has its own `client_id`, `tenant_id`, and MSAL token cache.

**Bootstrap:** On first start, if `MS365_CLIENT_ID` is set (via the MCPB installer or environment variable), a `default` profile is created and persisted to `profiles.json` automatically. If the variable is not set, the server starts with zero profiles and you must call `add_ms_profile`.

**Add accounts:**

```
add_ms_profile(name="personal", client_id="...", tenant_id="consumers")
add_ms_profile(name="work", client_id="...", tenant_id="mycompany.onmicrosoft.com")
```

**Use a specific profile** on any tool call:

```
list_emails(folder="Inbox", profile="work")
search_drive(query="Q1 report", profile="personal")
```

**Authenticate** (opens a browser window for OAuth the first time):

```
authenticate_ms_profile(profile="work")
```

**List profiles:**

```
list_ms_profiles()
```

**Change the default:**

```
set_default_ms_profile(profile="work")
```

Profiles are stored in `~/.microsoft-mcp/profiles.json`. Token caches are stored as `~/.microsoft-mcp/msal_cache_{name}.json`. After the first interactive login, MSAL handles token refresh silently.

> **Security note:** `profiles.json` and `msal_cache_*.json` contain refresh tokens. Do not commit them to version control. `MS365_CLIENT_ID` is not a secret and can be committed.

## Configuration

| Variable | Required | Default | Description |
|---|---|---|---|
| `MS365_CLIENT_ID` | Yes (for bootstrap) | — | Azure App Registration client ID for the default profile |
| `MS365_TENANT_ID` | No | `common` | Tenant ID for the default profile |
| `MS365_CREDENTIALS_DIR` | No | `~/.microsoft-mcp/` | Directory for `profiles.json` and token caches |

These variables are only used to bootstrap the `default` profile on first run. Once `profiles.json` exists they have no effect. Use the profile management tools to modify accounts.

## Development

```bash
# Install dependencies
uv sync

# Start the MCP server (stdio mode)
uv run mcp-microsoft

# Run the test suite
uv run pytest -q

# Rebuild the MCPB bundle
npx @anthropic-ai/mcpb pack
```

Tool implementations are organized by surface area under `src/mcp_microsoft/tools/`. Authentication is handled by [MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-python) with a per-profile serializable token cache. HTTP calls go through a shared async `httpx` client initialized at server startup.

## License

MIT
