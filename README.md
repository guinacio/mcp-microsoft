# mcp-microsoft

Microsoft 365 MCP server — **Mail, Calendar, and OneDrive** — for personal (Outlook.com / Live) and enterprise (Azure AD / Entra ID) accounts via the Microsoft Graph API.

Built with [FastMCP](https://github.com/jlowin/fastmcp), MSAL, and async httpx. Mirrors the architecture and tool surface of [mcp-google-workspace](https://github.com/guinacio/mcp-google-workspace).

> **Renamed from `mcp-outlook`.** See [Migration](#migration-from-mcp-outlook) below.

## Features

- **Mail** — Read, send, search, move, trash, and delete mail; draft creation, editing, and sending; folder management; attachment listing and download
- **Calendar** — List calendars and events; create, update, and delete events; RSVP to invitations; check free/busy availability; find meeting times
- **OneDrive** — List, search, upload, download, move, copy, and delete files and folders
- Works with both personal Microsoft accounts (Outlook.com / Live) and work / school accounts (Azure AD / Entra ID) — single App Registration, no admin consent needed for personal use
- Token cache with automatic silent refresh via MSAL

## Requirements

- Python 3.11+
- `uv` package manager
- An Azure App Registration (see Setup below)

## Setup

### 1. Create an Azure App Registration

1. Go to [portal.azure.com](https://portal.azure.com) > Azure Active Directory > App registrations > **New registration**.
2. **Supported account types**: select *"Accounts in any organizational directory and personal Microsoft accounts"*.
3. **Redirect URI**: Platform = *Mobile and desktop applications*, URI = `http://localhost`.
4. Under **Authentication**, enable **"Allow public client flows"** (required for the interactive loopback OAuth flow — no client secret needed).
5. Under **API permissions**, add the following **delegated** permissions:
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Calendars.ReadWrite`
   - `Files.ReadWrite`
   - `offline_access` (usually pre-added)
6. Click **Grant admin consent** if prompted (only required for enterprise tenants with admin-consent policies).
7. Copy the **Application (client) ID** from the Overview page.

### 2. Configure environment variables

Copy `.env.template` to `.env` and fill in your values:

```
MS365_CLIENT_ID=<your-application-client-id>
MS365_CREDENTIALS_DIR=   # optional — defaults to ~/.sentinel/microsoft-mcp/
```

Legacy env vars `OUTLOOK_CLIENT_ID` and `OUTLOOK_CREDENTIALS_DIR` still work as fallbacks.

### 3. Install

```powershell
uv pip install -e C:\Repositories\mcp-outlook
```

### 4. First run — browser consent

The first time the server starts (or whenever the token cache is missing), it opens a browser window for Microsoft's OAuth consent flow. After granting consent, the token is cached at `MS365_CREDENTIALS_DIR/msal_token_cache.json` and silent refresh handles subsequent runs automatically.

To reset auth (e.g. to change accounts or re-consent expanded scopes), delete the token cache file.

## Run (stdio)

```powershell
mcp-microsoft
# or
python -m mcp_microsoft.server
```

## Sentinel integration (`mcp_config.py`)

```python
"microsoft": {
    "command": sys.executable,
    "args": ["-m", "mcp_microsoft.server"],
    "env": {
        "MS365_CLIENT_ID": "<your-client-id>",
        "MS365_CREDENTIALS_DIR": str(_MS365_MCP_CREDS_DIR),
    },
}
```

## Tools

### Mail (11 tools)

| Tool | Description |
|---|---|
| `list_emails` | List messages from a folder. Params: `folder`, `max_results`, `unread_only`, `page_token`. |
| `read_email` | Fetch a full message by ID including body, headers, and attachment list. Supports `summary_mode`. |
| `search_emails` | Search messages via Graph KQL `$search`. |
| `send_email` | Send a new email. Params: `to`, `cc`, `bcc`, `subject`, `body`, `body_type`. |
| `reply_email` | Reply to a message. Params: `message_id`, `body`, `reply_all`, `body_type`. |
| `forward_email` | Forward a message. Params: `message_id`, `to`, `comment`. |
| `mark_as_read` | Mark a message as read. |
| `mark_as_unread` | Mark a message as unread. |
| `move_email` | Move a message to a folder by well-known name or folder ID. |
| `trash_email` | Move a message to Deleted Items (soft delete, recoverable). |
| `delete_email` | Permanently delete a message (irreversible). |

### Drafts (5 tools)

| Tool | Description |
|---|---|
| `create_draft` | Create a draft message. |
| `get_draft` | Fetch a draft by ID. |
| `list_drafts` | List draft messages. Supports pagination. |
| `update_draft` | Update draft fields. Only provided fields are changed. |
| `send_draft` | Send an existing draft by ID. |

### Folders (3 tools)

| Tool | Description |
|---|---|
| `list_folders` | List all mail folders including well-known and custom. Supports `include_child_folders`. |
| `create_folder` | Create a custom folder. Optionally nested under a parent folder. |
| `delete_folder` | Delete a folder and all its contents (irreversible). |

### Attachments (2 tools)

| Tool | Description |
|---|---|
| `list_attachments` | List attachments on a message: name, size, content type, attachment ID. |
| `download_attachment` | Download an attachment to a local file path or return base64. |

### Calendar (10 tools)

| Tool | Description |
|---|---|
| `list_calendars` | List all calendars in the mailbox. |
| `list_events` | List events from a calendar with optional date filtering. |
| `list_upcoming_events` | List events using calendarView (expands recurring events). Params: `start_datetime`, `end_datetime`. |
| `get_event` | Fetch full event details by ID including attendees, body, and recurrence. |
| `create_event` | Create a new event. Params: `subject`, `start_datetime`, `end_datetime`, `timezone`, `attendees`, `location`, `is_online_meeting`. |
| `update_event` | Update an existing event. Only provided fields are changed. |
| `delete_event` | Delete a calendar event. |
| `rsvp_event` | Accept, decline, or tentatively accept an event invitation. |
| `get_free_busy` | Check free/busy availability for one or more people in a time window. |
| `find_meeting_times` | Find available meeting time suggestions for a set of attendees. |

### OneDrive (8 tools)

| Tool | Description |
|---|---|
| `list_drive_items` | List files and folders. Defaults to OneDrive root. Accepts `folder_id`. |
| `get_drive_item` | Get metadata for a specific file or folder. |
| `search_drive` | Search files and folders by name or content. |
| `create_drive_folder` | Create a new folder. Optionally nested under a parent folder. |
| `upload_file` | Upload a local file. Auto-detects size and uses simple PUT (<4 MB) or resumable upload session. |
| `download_file` | Download a file to a local path. |
| `delete_drive_item` | Delete a file or folder (moved to recycle bin). |
| `move_or_copy_item` | Move or copy an item to a different folder. |

**Total: 39 tools**

## Migration from mcp-outlook

If you were using the previous `mcp-outlook` package:

1. Update your install:
   ```powershell
   uv remove mcp-outlook
   uv pip install -e C:\Repositories\mcp-outlook
   ```

2. In `mcp_config.py` or equivalent, change the server module:
   ```
   mcp_outlook.server → mcp_microsoft.server
   ```

3. Rename env vars in `.env` (optional — old names still work):
   - `OUTLOOK_CLIENT_ID` → `MS365_CLIENT_ID`
   - `OUTLOOK_CREDENTIALS_DIR` → `MS365_CREDENTIALS_DIR`

4. Optionally copy token cache to avoid re-authentication:
   ```
   cp ~/.sentinel/outlook-mcp/msal_token_cache.json ~/.sentinel/microsoft-mcp/msal_token_cache.json
   ```

5. Restart the MCP server. New scopes (Calendar, OneDrive) trigger a one-time browser consent popup.

## Token cache security

`msal_token_cache.json` contains refresh tokens and must not be committed to version control. It is listed in `.gitignore`. The `MS365_CLIENT_ID` is not a secret and can be committed.

## Tests

```powershell
uv run pytest -q
```
