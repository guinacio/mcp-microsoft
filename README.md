# mcp-microsoft

![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)
![License MIT](https://img.shields.io/badge/license-MIT-green)
![MCP](https://img.shields.io/badge/MCP-compatible-purple)
[![Tests](https://github.com/guilhermeinacio/mcp-microsoft/actions/workflows/ci.yml/badge.svg)](https://github.com/guilhermeinacio/mcp-microsoft/actions/workflows/ci.yml)
[![SafeSkill 91/100](https://img.shields.io/badge/SafeSkill-91%2F100_Verified%20Safe-brightgreen)](https://safeskill.dev/scan/guinacio-mcp-microsoft)

Microsoft 365 MCP server — Mail, Calendar, OneDrive, SharePoint, Contacts, and Teams via the Microsoft Graph API, with multi-account support.

## Overview

`mcp-microsoft` is a [Model Context Protocol](https://modelcontextprotocol.io) server that gives Claude (and any other MCP client) full access to your Microsoft 365 account. It covers six surface areas of the Microsoft Graph API: email, calendar, OneDrive file storage, SharePoint, contacts, and Teams — **93 tools** in total.

The server works with both personal Microsoft accounts (Outlook.com, Live) and enterprise accounts (Azure AD / Entra ID) using a single App Registration. Teams and SharePoint require a work or school account and are gated behind feature flags (`MCP_ENABLE_TEAMS` / `MCP_ENABLE_SHAREPOINT`). On manual installs, they auto-enable for corporate-oriented tenant values (`common`, `organizations`, or a specific tenant ID). In Claude Desktop / MCPB, the installer toggles remain authoritative. You can always override the default with the environment flags to force either service on or off. Teams meeting transcripts/recordings and Copilot AI insights are separate explicit opt-ins so the server does not request those additional scopes unless you enable them.

Multi-account support is a first-class feature. Named profiles let you configure separate client IDs for each account and switch between them on any tool call by passing `profile="work"`. Profiles and MSAL token caches are stored in `~/.microsoft-mcp/` and survive server restarts without re-authentication.

The server ships as an MCPB bundle (`mcp-microsoft.mcpb`) for zero-friction installation through the Claude Desktop Extension installer. It can also be run from source or wired directly into `claude_desktop_config.json`. Built with [FastMCP](https://github.com/jlowin/fastmcp), MSAL, and async httpx.

## Features

### Tools (93 total)

#### Mail (25 tools)

- `list_emails` — list messages from any folder with pagination and unread filter
- `read_email` — fetch the full body of a message by ID (supports summary mode)
- `search_emails` — search using Microsoft Graph KQL `$search` syntax (max 25 results)
- `filter_emails` — find emails by sender, recipient, subject, date range, or attachments with full pagination
- `send_email` — compose and send a new message (to/cc/bcc, HTML or plain text)
- `reply_email` — reply or reply-all to an existing message
- `forward_email` — forward a message to one or more recipients
- `mark_as_read` / `mark_as_unread` — toggle read state
- `move_email` — move to any folder by well-known name or folder ID
- `trash_email` — soft-delete to Deleted Items (recoverable)
- `delete_email` — permanently delete a message (irreversible)
- `bulk_move_emails` — move multiple messages to a folder in one operation
- `bulk_trash_emails` — move multiple messages to Deleted Items
- `bulk_delete_emails` — permanently delete multiple messages (irreversible)
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

#### Contacts (8 tools)

- `list_contacts` — list contacts with optional folder scope and field selection
- `get_contact` — retrieve a single contact by ID
- `create_contact` — create a new contact with name, email, phone, org, and notes
- `update_contact` — update any subset of contact fields
- `delete_contact` — delete a contact by ID (irreversible)
- `list_contact_folders` — enumerate contact folders in the mailbox
- `search_contacts` — search contacts by display name or email
- `get_contact_photo` — fetch a contact's profile photo as base64 or save to disk

#### Teams (25 tools)

> Teams tools require a work or school account (Azure AD / Entra ID) with a specific tenant ID. They are not available for personal Outlook.com / Live accounts.

- `teams_list_joined` / `teams_get` — list and inspect teams
- `teams_list_channels` / `teams_get_channel` / `teams_create_channel` — manage channels
- `teams_list_channel_messages` / `teams_get_channel_message` / `teams_send_channel_message` — read and post channel messages
- `teams_reply_to_channel_message` / `teams_list_message_replies` — manage channel threads
- `teams_list_chats` / `teams_get_chat` / `teams_list_chat_messages` / `teams_send_chat_message` / `teams_create_chat` — 1:1 and group chats
- `teams_create_meeting` / `teams_get_meeting` / `teams_find_meeting_by_url` / `teams_list_meetings` — online meetings with join URLs
- `teams_list_meeting_transcripts` / `teams_get_meeting_transcript` — Teams meeting transcript metadata and VTT content
- `teams_list_meeting_recordings` / `teams_download_meeting_recording` — Teams meeting recording metadata and downloads
- `teams_list_meeting_ai_insights` / `teams_get_meeting_ai_insight` — Copilot meeting recap/insight metadata and full detail

> Teams meeting transcripts and recordings require explicit opt-in via `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` plus the delegated permissions `OnlineMeetingTranscript.Read.All` and `OnlineMeetingRecording.Read.All`. Teams AI insights require explicit opt-in via `MCP_ENABLE_TEAMS_AI_INSIGHTS`, the delegated permission `OnlineMeetingAiInsight.Read.All`, and Microsoft 365 Copilot licensing.

#### Profile Management (5 tools)

- `list_ms_profiles` — list all configured profiles and which is the default
- `add_ms_profile` — add a new account (name, client_id, tenant_id)
- `remove_ms_profile` — remove a profile and delete its cached tokens
- `authenticate_ms_profile` — trigger interactive OAuth for a profile
- `set_default_ms_profile` — change which profile is used when none is specified

## Installation

### Option A: Claude Desktop Extension (MCPB) — Recommended

```bash
npx @anthropic-ai/mcpb install mcp-microsoft-0.6.0.mcpb
```

The installer prompts for your Azure App Registration details (see [Azure Setup](#azure-setup)):

| Prompt | Description |
|---|---|
| **Azure Client ID** | Application (client) ID from your App Registration |
| **Tenant ID** | `common` for personal + work, `consumers` for personal only, or your org's tenant ID/domain |
| **Credentials Directory** | Optional. Defaults to `~/.microsoft-mcp/` |
| **Enable Teams Tools** | Toggle Teams tools on/off (requires work/school account) |
| **Enable Teams Meeting Artifacts** | Toggle Teams transcript/recording tools on/off (requires work/school account and extra Graph permissions) |
| **Enable Teams AI Insights** | Toggle Teams Copilot AI insight tools on/off (requires work/school account, extra Graph permissions, and Copilot licensing) |
| **Enable SharePoint Tools** | Toggle SharePoint tools on/off (requires work/school account) |

A `default` profile is created automatically from these values.

### Option B: From Source

```bash
git clone https://github.com/guilhermeinacio/mcp-microsoft.git
cd mcp-microsoft
uv sync
export MS365_CLIENT_ID=your-client-id
export MS365_TENANT_ID=common
# Optional overrides:
# export MCP_ENABLE_SHAREPOINT=false
# export MCP_ENABLE_TEAMS=false
# export MCP_ENABLE_TEAMS_MEETING_ARTIFACTS=true
# export MCP_ENABLE_TEAMS_AI_INSIGHTS=true
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
   - `Contacts.ReadWrite`
   - `Files.ReadWrite`
   - `offline_access` *(usually pre-added)*
7. If you want to enable **SharePoint** tools, also add `Sites.ReadWrite.All`.
8. If you want to enable **Teams** tools, also add:
   - `Team.ReadBasic.All`
   - `Channel.ReadBasic.All`
   - `Channel.Create`
   - `ChannelMessage.Read.All`
   - `ChannelMessage.Send`
   - `Chat.ReadWrite`
   - `Chat.Create`
   - `OnlineMeetings.ReadWrite`
9. If you want to enable **Teams meeting transcripts and recordings**, also add:
   - `OnlineMeetingTranscript.Read.All`
   - `OnlineMeetingRecording.Read.All`
10. If you want to enable **Teams Copilot AI insights**, also add:
   - `OnlineMeetingAiInsight.Read.All`
   All users of applications that call this API need a Microsoft 365 Copilot license.
11. For admin-restricted permissions such as `Sites.ReadWrite.All`, click **Grant admin consent**. Your IT administrator must approve this once per tenant.
12. From the **Overview** page, copy the **Application (client) ID** and, if targeting a specific tenant, the **Directory (tenant) ID**.

For a detailed walkthrough with screenshots, see [`docs/azure-setup.md`](docs/azure-setup.md).

## Account Type Compatibility

| Module | Personal (Outlook.com/Live) | Work/School (Azure AD/Entra) |
|---|---|---|
| Mail | Yes | Yes |
| Calendar | Yes | Yes |
| OneDrive | Yes | Yes |
| Contacts | Yes | Yes |
| SharePoint | No | Yes (admin consent required) |
| Teams | No | Yes |

Personal accounts use `tenant_id=consumers`. For access to both personal and work accounts via a single profile, use `tenant_id=common`. Teams transcript/recording and AI-insight APIs remain work-account-only and are additional explicit opt-ins on top of the base Teams tools.

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

Profiles are stored in `~/.microsoft-mcp/profiles.json`. Token caches are stored as `~/.microsoft-mcp/msal_cache_{name}.bin` and encrypted at rest via OS-native APIs (DPAPI on Windows, Keychain on macOS, libsecret on Linux) using [msal-extensions](https://github.com/AzureAD/microsoft-authentication-extensions-for-python). Legacy plaintext caches from older versions are migrated automatically on first run. After the first interactive login, MSAL handles token refresh silently.

> **Security note:** `profiles.json` contains the client/tenant IDs only — it has no secrets but is restricted to mode `0600` on POSIX. The `msal_cache_*.bin` files hold encrypted refresh tokens; do not commit either to version control. `MS365_CLIENT_ID` is not a secret and can be committed.
>
> **Disabling destructive tools:** set `MCP_DISABLE_DELETION_TOOLS=1` to suppress registration of all permanent-delete tools (`delete_email`, `bulk_delete_emails`, `delete_event`, `delete_contact`, `delete_folder`, `delete_drive_item`, `delete_list_item`, `remove_ms_profile`). Recoverable variants such as `trash_email` remain available.

## Configuration

| Variable | Required | Default | Description |
|---|---|---|---|
| `MS365_CLIENT_ID` | Yes (for bootstrap) | — | Azure App Registration client ID for the default profile |
| `MS365_TENANT_ID` | No | `common` | Tenant ID for the default profile |
| `MS365_CREDENTIALS_DIR` | No | `~/.microsoft-mcp/` | Directory for `profiles.json` and token caches |
| `MCP_ENABLE_SHAREPOINT` | No | auto-detect | Set to `true` to force-enable SharePoint tools or `false` to force-disable them |
| `MCP_ENABLE_TEAMS` | No | auto-detect | Set to `true` to force-enable Teams tools or `false` to force-disable them |
| `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` | No | `false` | Set to `true` to register Teams transcript and recording tools and request the related delegated scopes |
| `MCP_ENABLE_TEAMS_AI_INSIGHTS` | No | `false` | Set to `true` to register Teams Copilot AI insight tools and request `OnlineMeetingAiInsight.Read.All` |
| `MCP_DISABLE_DELETION_TOOLS` | No | `false` | Set to `true` to suppress registration of all permanent-delete tools (recoverable trash variants remain) |

`MS365_CLIENT_ID`, `MS365_TENANT_ID`, and `MS365_CREDENTIALS_DIR` are only used to bootstrap the `default` profile on first run. `MCP_ENABLE_TEAMS` and `MCP_ENABLE_SHAREPOINT` participate in the existing auto-detection behavior; `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` and `MCP_ENABLE_TEAMS_AI_INSIGHTS` are explicit opt-ins only.

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

Tool implementations are organized by surface area under `src/mcp_microsoft/tools/`: `mail.py`, `drafts.py`, `folders.py`, `attachments.py`, `calendar.py`, `onedrive.py`, `sharepoint.py`, `contacts.py`, `teams.py`, `profiles.py`. Authentication is handled by [MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-python) with a per-profile serializable token cache. HTTP calls go through a shared async `httpx` client initialized at server startup.

## License

MIT
