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

## App registration for the remote (http) server

> **This is a separate App Registration from the one above.** Everything in Steps 1–8 creates a **public client** (platform: "Mobile and desktop applications", no client secret) for the single-user **stdio** server (MCPB / Claude Desktop / running from source). The multi-user **remote http server** (`MCP_TRANSPORT=http`, added in 0.8.0) needs a **confidential client** instead — a different platform type, with a client secret, because it performs a server-side On-Behalf-Of (OBO) token exchange on behalf of each connecting user. You cannot reuse the stdio registration for this; create a new one (or add a second platform + secret to a fresh registration dedicated to the remote server — don't bolt this onto the public-client registration your desktop users already depend on).
>
> Background on *why* the server needs this: see the "Remote server — multi-user (Streamable HTTP)" section of the main [README](../README.md).

### 1. Register the application

1. [portal.azure.com](https://portal.azure.com) → **App registrations** → **+ New registration**.
2. **Name**: anything identifiable, e.g. `mcp-microsoft-remote`.
3. **Supported account types**: work/school only — either *this organizational directory only* or *any organizational directory*, depending on who should be able to sign in. Do **not** choose a personal-account option: On-Behalf-Of and custom API scopes are not reliably supported for consumer Microsoft accounts, so the remote server doesn't support them either (personal accounts stay on stdio).
4. **Redirect URI**: platform **Web** (not "Mobile and desktop applications" — that's the stdio registration's platform), URI:

   ```
   {MCP_BASE_URL}/auth/callback
   ```

   where `{MCP_BASE_URL}` is the exact public HTTPS URL from your `.env` (e.g. `https://mcp.example.com/auth/callback`). `/auth/callback` is FastMCP's `AzureProvider` default redirect path.
5. Click **Register**.

### 2. Expose an API

1. Left sidebar → **Expose an API**.
2. **Application ID URI**: click **Add**, accept the default `api://{client_id}` (this is also `AzureProvider`'s default — no config needed on the server side if you keep it), and **Save**.
3. **+ Add a scope**:
   - **Scope name**: `mcp-access` (matches the server's default `MCP_AUTH_REQUIRED_SCOPE`; if you pick a different name, set `MCP_AUTH_REQUIRED_SCOPE` to match).
   - **Who can consent**: Admins and users (or Admins only, if you want to gate access centrally).
   - **Admin/user consent display name & description**: anything descriptive, e.g. "Access mcp-microsoft" / "Allows the MCP client to call mcp-microsoft on your behalf."
   - **State**: Enabled.

### 3. Set the access token version (manifest edit)

1. Left sidebar → **Manifest**.
2. Find `"requestedAccessTokenVersion"` and set it to `2`:
   ```json
   "requestedAccessTokenVersion": 2
   ```
3. **Save**. This is required for Azure to issue v2.0 tokens with the claim shapes (`scp`, `oid`, `preferred_username`, etc.) `AzureProvider` and the server's audit logging expect.

### 4. Create a client secret

1. Left sidebar → **Certificates & secrets** → **+ New client secret**.
2. Description: anything, e.g. `mcp-microsoft-remote-prod`. Pick an expiry (Azure caps this at 24 months).
3. Copy the secret **value** immediately — it's shown once. This becomes `MCP_AUTH_CLIENT_SECRET`.

> **Rotation note:** the secret has a hard expiry — plan to rotate it before then. Create a new secret alongside the old one (both are valid simultaneously), roll `MCP_AUTH_CLIENT_SECRET` to the new value and restart the server, then delete the old secret from the App Registration once you've confirmed the new one is live. The server holds no long-lived cache of the secret itself beyond process memory, so a restart is sufficient — no data migration needed.

### 5. Add delegated Graph permissions

Same permission set as the stdio server (Step 5 above), added to **this** registration instead:

**Base (always required):**

| Permission | What it enables |
|------------|----------------|
| `Mail.ReadWrite` | Read, move, delete, and organise email |
| `Mail.Send` | Send email |
| `Calendars.ReadWrite` | Read and manage calendar events |
| `Contacts.ReadWrite` | Read and manage contacts |
| `Files.ReadWrite` | Read and manage OneDrive files |

`offline_access` is added automatically by `AzureProvider` — no need to request it explicitly.

**Optional, only if the corresponding feature flag is enabled** (see `MCP_ENABLE_*` in `.env.template`):

| Flag | Permissions |
|------|-------------|
| `MCP_ENABLE_TEAMS` | `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `Channel.Create`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `Chat.ReadWrite`, `Chat.Create`, `OnlineMeetings.ReadWrite` |
| `MCP_ENABLE_TEAMS_MEETING_ARTIFACTS` (requires Teams also enabled) | `OnlineMeetingTranscript.Read.All`, `OnlineMeetingRecording.Read.All` |
| `MCP_ENABLE_TEAMS_AI_INSIGHTS` (requires Teams also enabled, plus Copilot licensing) | `OnlineMeetingAiInsight.Read.All` |
| `MCP_ENABLE_SHAREPOINT` | `Sites.ReadWrite.All` |

Only request permissions for services you're actually enabling — the server builds its OBO scope request (`https://graph.microsoft.com/.default`) from whatever's granted on this registration, gated by these same env flags, which in http mode must be set explicitly (no auto-detection).

### 6. Grant admin consent

Because this is a confidential-client, work/school-only registration, plan on **tenant-wide admin consent** rather than per-user consent:

- If you're the tenant admin: **API permissions** → **Grant admin consent for [tenant]**.
- Otherwise: send your IT administrator the **Application (client) ID** and ask them to grant it from the App Registration's API permissions page.

`Sites.ReadWrite.All` (SharePoint) in particular will not work without admin consent in most tenants. Doing this up front for the whole permission set avoids each new user hitting an individual consent prompt they may not be able to approve themselves.

### 7. Collect the values for `.env`

| Azure value | Env var |
|---|---|
| Application (client) ID | `MCP_AUTH_CLIENT_ID` |
| Client secret value (Step 4) | `MCP_AUTH_CLIENT_SECRET` |
| Directory (tenant) ID, or `organizations` | `MCP_AUTH_TENANT_ID` |
| Your reverse proxy's public HTTPS URL | `MCP_BASE_URL` |
| The scope name from Step 2 (only if not `mcp-access`) | `MCP_AUTH_REQUIRED_SCOPE` |

Fill these into `.env` (copy from `.env.template`'s "Remote server (http) mode" section) and start the server with `MCP_TRANSPORT=http` — see the README's "Remote server — multi-user (Streamable HTTP)" section for the full quickstart, including the Docker Compose path.

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
