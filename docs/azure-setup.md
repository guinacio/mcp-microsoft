# Azure App Registration Setup

This guide walks you through creating your own Azure App Registration to get the **Client ID** and **Tenant ID** needed to install and use `mcp-microsoft`.

> **Why your own App Registration?**
> Each person should register their own application in Azure. This gives you full control over which permissions are granted, avoids sharing credentials with others, and prevents permission conflicts where one user's admin consent settings affect another user's access.

---

## Prerequisites

- A Microsoft account (personal Outlook.com/Live **or** a work/school Microsoft 365 account)
- Access to [portal.azure.com](https://portal.azure.com) — sign in with the same account you want to use with Claude

---

## Step 1 — Open App Registrations

1. Go to [portal.azure.com](https://portal.azure.com) and sign in.
2. In the top search bar, type **"App registrations"** and click the result under **Services**.
3. Click **+ New registration**.

---

## Step 2 — Register the Application

Fill in the form:

| Field | Value |
|-------|-------|
| **Name** | `mcp-microsoft` (or any name you prefer) |
| **Supported account types** | See the table below |
| **Redirect URI** | Platform: **Mobile and desktop applications** → URI: `http://localhost` |

### Choosing Supported Account Types

| Your situation | Choose |
|----------------|--------|
| Personal Outlook.com / Live account only | *Personal Microsoft accounts only* |
| Work/school Microsoft 365 account only | *Accounts in this organizational directory only* |
| Both personal and work accounts | *Accounts in any organizational directory and personal Microsoft accounts* |

> If you're unsure, choose **"Accounts in any organizational directory and personal Microsoft accounts"** — it works for all account types and you can always restrict it later.

Click **Register**.

---

## Step 3 — Copy Your IDs

After registering, you land on the app's **Overview** page. Copy these two values — you'll need them during installation:

| Value | Where to find it | Example |
|-------|-----------------|---------|
| **Application (client) ID** | Overview page, first field | `a1b2c3d4-1234-5678-abcd-ef0123456789` |
| **Directory (tenant) ID** | Overview page, second field | `9a8b7c6d-...` or `common` |

### What Tenant ID to use

- **Personal account only** → use `consumers`
- **Work/school account only** → use the **Directory (tenant) ID** from the Overview page
- **Both personal and work** → use `common`

---

## Step 4 — Enable Public Client Flows

This allows the server to authenticate using a browser login flow without needing a client secret.

1. In the left sidebar, click **Authentication**.
2. Scroll down to **Advanced settings**.
3. Set **Allow public client flows** to **Yes**.
4. Click **Save**.

---

## Step 5 — Add API Permissions

1. In the left sidebar, click **API permissions**.
2. Click **+ Add a permission** → **Microsoft Graph** → **Delegated permissions**.
3. Search for and add the permissions for the features you want:

### Base permissions (required for everyone)

| Permission | What it enables |
|------------|----------------|
| `Mail.ReadWrite` | Read, move, delete, and organise email |
| `Mail.Send` | Send email |
| `Calendars.ReadWrite` | Read and manage calendar events |
| `Contacts.ReadWrite` | Read and manage contacts |
| `Files.ReadWrite` | Read and manage OneDrive files |
| `offline_access` | Keep you signed in (token refresh) — usually pre-added |

### Teams tools (optional — work/school accounts only)

| Permission | What it enables |
|------------|----------------|
| `Team.ReadBasic.All` | List joined teams |
| `Channel.ReadBasic.All` | List channels |
| `Channel.Create` | Create channels |
| `ChannelMessage.Read.All` | Read channel messages |
| `ChannelMessage.Send` | Post channel messages |
| `Chat.ReadWrite` | Read and send chat messages |
| `Chat.Create` | Create new chats |
| `OnlineMeetings.ReadWrite` | Create and manage Teams meetings |

### SharePoint tools (optional — work/school accounts only)

| Permission | What it enables |
|------------|----------------|
| `Sites.ReadWrite.All` | Browse and edit SharePoint sites and lists |

> **SharePoint note:** `Sites.ReadWrite.All` requires **admin consent** in enterprise tenants. See Step 6.

### Teams meeting transcripts & recordings (optional — explicit opt-in)

| Permission | What it enables |
|------------|----------------|
| `OnlineMeetingTranscript.Read.All` | Read meeting transcript VTT files |
| `OnlineMeetingRecording.Read.All` | Access meeting recording metadata and downloads |

### Teams Copilot AI insights (optional — requires Copilot license)

| Permission | What it enables |
|------------|----------------|
| `OnlineMeetingAiInsight.Read.All` | Read Copilot meeting recaps and insights |

---

## Step 6 — Grant Admin Consent (Work Accounts Only)

If you are on a work/school account and added `Sites.ReadWrite.All` or any other admin-restricted permission, you need a one-time admin consent:

- If **you are the tenant admin**: click **Grant admin consent for [your organisation]** on the API permissions page and confirm.
- If **you are not the admin**: ask your IT administrator to approve the permissions for your app registration. They can do this by navigating to the App Registration in their Azure portal and clicking the same button.

Personal Microsoft accounts (Outlook.com/Live) **do not require admin consent** — all permissions are user-consented on first login.

---

## Step 7 — Install mcp-microsoft

You now have everything you need. During installation, enter:

- **Azure Client ID** → the Application (client) ID from Step 3
- **Tenant ID** → the value from Step 3 based on your account type

### MCPB (Claude Desktop, recommended)

Download `mcp-microsoft-0.6.0.mcpb` from the [latest release](https://github.com/guinacio/mcp-microsoft/releases/latest) and double-click to open in Claude Desktop. The installer will prompt for the values above.

### Manual (from source)

```bash
export MS365_CLIENT_ID=your-application-client-id
export MS365_TENANT_ID=common   # or consumers, or your tenant ID
uv run mcp-microsoft
```

---

## Step 8 — First Sign-in

On the first tool call, Claude will ask you to authenticate:

1. A browser window opens automatically to the Microsoft login page.
2. Sign in with the Microsoft account you want to use.
3. Accept the permission consent screen — this is the list of permissions from Step 5.
4. The browser redirects to `localhost` and closes. You are now authenticated.

After this, the server refreshes tokens silently in the background. You only need to re-authenticate if you revoke access or the refresh token expires (typically 90 days of inactivity).

---

## Troubleshooting

### "AADSTS50011: The redirect URI does not match"
Make sure you added `http://localhost` (not `https://`) under **Mobile and desktop applications** in Authentication, not under Web.

### "AADSTS65001: The user or administrator has not consented"
For work accounts, your IT administrator needs to grant admin consent for the app. Share your **Application (client) ID** with them and ask them to approve it in Azure AD.

### "AADSTS700016: Application not found in the directory"
Your Tenant ID doesn't match the account. If using a personal account, set Tenant ID to `consumers`. If using a work account, use the Directory (tenant) ID from the Overview page.

### SharePoint tools not appearing
SharePoint requires `MCP_ENABLE_SHAREPOINT=true` (or the toggle in the MCPB installer) and a work/school account. It will not work with personal Outlook.com accounts.

### Teams tools not appearing
Teams requires `MCP_ENABLE_TEAMS=true` (or the toggle in the installer) and a work/school account with a specific tenant ID (not `consumers`).

---

## Summary

| What you need | Where to get it |
|---------------|----------------|
| Application (client) ID | Azure Portal → App registrations → your app → Overview |
| Tenant ID | Azure Portal → App registrations → your app → Overview (or use `common` / `consumers`) |
| Admin consent | Only needed for `Sites.ReadWrite.All` on enterprise tenants |

Total setup time: approximately 10 minutes.
