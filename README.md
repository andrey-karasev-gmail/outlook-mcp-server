# Outlook MCP Server

A Model Context Protocol (MCP) server for accessing and managing Outlook.com email accounts through Claude and other MCP clients.

## Features

- **List emails**: Retrieve recent messages from your Outlook inbox
- **Read emails**: Get full email details including body, attachments info
- **Delete emails**: Remove emails permanently
- **Mark emails**: Set emails as read/unread
- **Move emails**: Transfer emails to other folders
- **Attachments**: List and download email attachments
- **Flags**: Set, complete, or clear follow-up flags (with optional due dates)
- **Categories**: List/create categories and apply them to emails
- **Inbox rules**: List, create, and delete server-side Inbox rules
- **Drafts**: Create new draft emails and reply drafts in Drafts (never sent)

## Setup

### Prerequisites

- Node.js 18+ and npm
- An Outlook.com or Microsoft 365 account
- Microsoft Azure Application Registration

### 1. Register Azure Application

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Microsoft Entra ID** (formerly Azure Active Directory) > **App registrations**
3. Click **New registration**
4. Fill in the details:
   - Name: `Outlook MCP Server`
   - Supported account types: `Accounts in any organizational directory and personal Microsoft accounts`
5. Click **Register**

### 2. Add the Redirect URI

The one-time token script (`scripts/get-auth-token.js`) receives the
authorization code on a local callback, so the app registration must
whitelist that exact address — without it, sign-in fails with a
redirect-URI mismatch error (AADSTS50011).

1. In your app registration, go to **Authentication**
2. Click **Add a platform** > **Web**
3. Enter the redirect URI: `http://localhost:3000/callback`
4. Leave the implicit-grant checkboxes **off** and click **Configure**

No other platform (mobile/desktop, SPA) is needed.

### 3. Create Client Secret

1. In your app registration, go to **Certificates & secrets**
2. Click **New client secret**
3. Add a description and select expiration
4. Copy the **Value** (you'll need this as `OUTLOOK_CLIENT_SECRET`)

### 4. Set API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission** > **Microsoft Graph**
3. Select **Delegated permissions**
4. Search and add these permissions:
   - `offline_access` — **required for unattended operation**: without it
     Microsoft issues no refresh token at all, and the server would need a
     manual sign-in for every access token. (It sits in the **OpenID
     permissions** group, together with `openid` and `profile`, which the
     token script also requests.)
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `MailboxSettings.ReadWrite` (for categories and Inbox rules; flags work with `Mail.ReadWrite` alone)
5. Click **Grant admin consent** if you're in an organization. For a
   personal Microsoft account there is no admin consent — you grant these
   permissions yourself on the consent screen during the token step below.

`Mail.Send` is **not** needed: the server has no send tool, and neither
the token script nor the refresh flow requests it.

### 5. Get Your Credentials

From your app registration overview page, copy:
- **Application (client) ID** → `OUTLOOK_CLIENT_ID`
- **Directory (tenant) ID** → `OUTLOOK_TENANT_ID` (or use "common")

### 6. Get Refresh Token (one-time, seeds unattended refresh)

This is the **only manual sign-in in the whole setup**. The script starts
a temporary server on `http://localhost:3000/callback` (the redirect URI
from step 2), opens your browser, captures the authorization code, and
exchanges it for a refresh token. Sign in with the account whose mailbox
the server will access, and accept the permissions from step 4 on the
consent screen.

Put `OUTLOOK_CLIENT_ID` and `OUTLOOK_CLIENT_SECRET` in `.env` first,
then run:

```bash
# Generate auth URL
node scripts/get-auth-token.js

# Or, with no cloning, using the published package:
npx --package=@web4w3/outlook-mcp-server outlook-mcp-server-get-auth-token
```

Copy the printed `OUTLOOK_REFRESH_TOKEN=…` line into `.env`. From that
moment the server refreshes and rotates the token by itself — see
[Token lifecycle](#token-lifecycle). You only repeat this step if the
stored token dies (revoked, or unused for ~90 days) or after adding new
API permissions that the token must carry.

### 7. Install

```bash
npx -y @web4w3/outlook-mcp-server
```

No cloning or build step needed. The package is pre-compiled and available on npm as [`@web4w3/outlook-mcp-server`](https://www.npmjs.com/package/@web4w3/outlook-mcp-server).

### 8. Configure in Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": ["-y", "@web4w3/outlook-mcp-server"],
      "env": {
        "OUTLOOK_CLIENT_ID": "your_client_id",
        "OUTLOOK_CLIENT_SECRET": "your_client_secret",
        "OUTLOOK_TENANT_ID": "common",
        "OUTLOOK_REFRESH_TOKEN": "your_refresh_token"
      }
    }
  }
}
```

### Remote (HTTP) mode

By default the server speaks MCP over stdio for a client on the same machine.
To host it instead (for example on an always-on home server behind a tunnel):

```bash
MCP_TRANSPORT=http MCP_HTTP_PORT=3000 MCP_HTTP_TOKEN=<redacted> npm start
```

The server then listens for Streamable HTTP MCP requests on `/mcp` and
requires the token as `Authorization: Bearer <token>` on every request.
The endpoint is stateless; each request is independent. Keep it behind a
private tunnel or a TLS-terminating proxy — do not expose it to the open
internet without one.

## Token lifecycle

The server authenticates with a delegated refresh token and manages it
carefully, because Microsoft **rotates** refresh tokens: most token
responses carry a new refresh token and the old one is eventually
invalidated. Relying on a static token from `.env` therefore breaks
sooner or later — this is how the server avoids that:

- **Seed**: `OUTLOOK_REFRESH_TOKEN` in `.env` is only the starting
  point, minted once with `scripts/get-auth-token.js`.
- **State file**: whenever a token response contains a new refresh
  token, it is adopted in memory and persisted atomically (temp file +
  rename, permissions `0600`) to a JSON state file:
  `{"refreshToken": "…", "updatedAt": "<ISO timestamp>"}`.
  Location: the `OUTLOOK_TOKEN_STATE_PATH` environment variable when
  set, otherwise `<working directory>/data/refresh-token.json`.
- **Precedence**: at startup the state file wins over the `.env` seed —
  it holds the freshest token. A missing or corrupt state file falls
  back to the seed with a warning on stderr; if neither source has a
  token, startup fails as before. If `.env` has no refresh token but a
  valid state file exists, the server still starts.
- **Access tokens** are cached in memory and reused until 5 minutes
  before they expire; concurrent refreshes share one request.
- **Retries**: the token request is retried up to 3 attempts with
  backoff (~0.5s, then ~2s) on transient failures only — network
  errors, HTTP 429 and HTTP 5xx. OAuth errors such as `invalid_grant`
  are never retried.
- **Recovery**: if the stored refresh token dies (revoked, expired, or
  superseded by a rotation whose write was lost), tool calls fail with
  an `invalid_grant` error saying so. Re-run
  `scripts/get-auth-token.js`, put the new token into
  `OUTLOOK_REFRESH_TOKEN` in `.env`, and delete the state file so the
  fresh seed is used; then restart the server.

Operational notes: keep the state file on durable storage (for a
systemd deployment, the service's working directory), and do not copy
an old state file over a newer one — the older refresh token in it may
already be invalidated.

### Enabling unattended refresh on a server

Checklist for an always-on deployment (v1.0.6 or later — the version
that ships the token manager):

1. Complete the app registration exactly as in steps 1–5 above; the
   `offline_access` permission and the `http://localhost:3000/callback`
   redirect URI are the two pieces unattended refresh depends on.
2. Mint the seed token once (step 6) **after** all permissions are in
   place, and put it in the server's `.env` as `OUTLOOK_REFRESH_TOKEN`.
3. Point `OUTLOOK_TOKEN_STATE_PATH` at durable storage — an absolute
   path is safest under systemd, e.g.
   `OUTLOOK_TOKEN_STATE_PATH=/home/<user>/outlook-mcp-server/data/refresh-token.json`.
   The default (`data/refresh-token.json` in the working directory) is
   fine when the working directory itself is durable. The service user
   needs write access to that directory.
4. Start (or restart) the server. On the first refresh that returns a
   rotated token, the state file appears (mode `0600`) and stderr logs
   `refresh token rotated; new token persisted`.
5. Verify after ~1 hour of uptime: the state file's `updatedAt` is
   fresh. From then on the chain is self-sustaining — access tokens
   refresh about hourly, every rotation is persisted, and restarts
   resume from the state file. A refresh token dies of old age only
   after ~90 days **unused**, which a running server never lets happen.
6. Run **one** server instance per mailbox against a given state file:
   two instances rotating the same token chain race each other, and the
   loser's stored token may be the superseded one.

## Available Tools

### `outlook_list_folders`
List all mail folders in your Outlook mailbox.

**Parameters:** None

**Returns:** Array of folder objects with `id` and `displayName`

---

### `outlook_list_emails`
List recent emails from your mailbox.

**Parameters:**
- `limit` (number, 1-100, default: 20): Number of emails to retrieve
- `folder` (string, optional): Folder name (default: "inbox")

**Returns:** Array of email summaries with subject, sender, date, read status, and preview

---

### `outlook_get_email`
Get full details of a specific email including body, recipients, and attachments info.

**Parameters:**
- `id` (string, required): Email message ID
- `markAsRead` (boolean, default: false): Mark email as read when retrieving

**Returns:** Full email object with complete details

---

### `outlook_delete_email`
Permanently delete an email.

**Parameters:**
- `id` (string, required): Email message ID

**Returns:** Confirmation message

---

### `outlook_mark_email`
Mark an email as read or unread.

**Parameters:**
- `id` (string, required): Email message ID
- `isRead` (boolean, required): True to mark as read, false for unread

**Returns:** Confirmation message

---

### `outlook_move_email`
Move an email to a different folder.

**Parameters:**
- `id` (string, required): Email message ID
- `destinationFolder` (string, required): Target folder ID or name

**Returns:** Confirmation message

**Note:** If `destinationFolder` is a folder name (not an ID), the server will:
1. Look up the folder by name
2. If not found, create a new folder with that name
3. Move the email to the resolved folder ID

---

### `outlook_get_attachments`
Get list of attachments for an email.

**Parameters:**
- `messageId` (string, required): Email message ID

**Returns:** Array of attachment objects with name, size, and MIME type

---

### `outlook_get_attachment_content`
Download attachment content (base64 encoded).

**Parameters:**
- `messageId` (string, required): Email message ID
- `attachmentId` (string, required): Attachment ID

**Returns:** Attachment object with base64 encoded content

---

### `outlook_flag_email`
Set, complete, or clear the follow-up flag on an email.

**Parameters:**
- `id` (string, required): Email message ID
- `status` (string, required): `flagged`, `complete`, or `notFlagged`
- `dueDateTime` (string, optional): ISO date or datetime, only with `flagged`
- `startDateTime` (string, optional): ISO date or datetime, only with `flagged`
- `timeZone` (string, optional): Time zone for the dates (default `UTC`)

**Returns:** Confirmation message

---

### `outlook_list_categories`
List all categories from the master category list.

**Returns:** Array of categories with `id`, `displayName`, and `color`

---

### `outlook_create_category`
Create a new category in the master category list.

**Parameters:**
- `displayName` (string, required): Category name
- `color` (string, optional): `none` or a preset `preset0`–`preset24` (default `preset0`)

**Returns:** The created category

---

### `outlook_categorize_email`
Change the categories of an email.

**Parameters:**
- `id` (string, required): Email message ID
- `set` (string[], optional): Replace the whole category list
- `add` (string[], optional): Add categories to the current ones
- `remove` (string[], optional): Remove categories from the current ones

**Returns:** The resulting list of categories on the email

---

### `outlook_list_rules`
List all Inbox rules (server-side rules that act on incoming mail).

**Returns:** Array of rules with `id`, `displayName`, `sequence`, `isEnabled`, `conditions`, and `actions`

---

### `outlook_create_rule`
Create an Inbox rule. Rules run on the server and apply to incoming mail
whether or not this server is running.

**Parameters:**
- `displayName` (string, required): Rule name
- `isEnabled` (boolean, optional): Default `true`
- `sequence` (number, optional): Evaluation order, lower runs first
- Conditions (all optional; no conditions matches all mail):
  - `fromAddresses` (string[]): Sender email addresses
  - `senderContains` (string[]): Substrings to match in the sender
  - `subjectContains` (string[]): Substrings to match in the subject
  - `hasAttachments` (boolean)
  - `importance` (string): `low`, `normal`, or `high`
- Actions (at least one required):
  - `moveToFolder` (string): Folder ID or display name (resolved like `outlook_move_email`)
  - `markAsRead` (boolean)
  - `markImportance` (string): `low`, `normal`, or `high`
  - `assignCategories` (string[])
  - `delete` (boolean)
  - `stopProcessingRules` (boolean)

**Returns:** The created rule

---

### `outlook_delete_rule`
Delete an Inbox rule by its ID (see `outlook_list_rules`).

**Parameters:**
- `ruleId` (string, required): Rule ID

**Returns:** Confirmation message

---

### `outlook_create_draft`
Create a draft email in the Drafts folder. The message is never sent; the body is plain text so line breaks are preserved.

**Parameters:**
- `subject` (string, required): Draft subject
- `toRecipients` (string[], required): Recipient email addresses
- `ccRecipients` (string[], optional): CC email addresses
- `body` (string, required): Plain-text body

**Returns:** The created draft (`id`, `subject`, `toAddresses`, `conversationId`, `isDraft`)

---

### `outlook_create_reply_draft`
Create a reply draft to an existing email. The draft keeps the thread (`Re:` subject, same conversation) and its body is replaced with the given plain-text reply. The message is never sent.

**Parameters:**
- `id` (string, required): ID of the original message
- `body` (string, required): Plain-text reply body

**Returns:** The created reply draft (`id`, `subject`, `toAddresses`, `conversationId`, `isDraft`)

---

## Integration with Email Analytics Agent

This server is designed to work with the [Email Analytics Agent](https://github.com/andrey-karasev/email-analytics-agent), which provides:
- Intelligent email grouping by sender domain
- Interactive folder assignment
- Rule learning for automatic email organization

## Development

Run in development mode with auto-reload:

```bash
npm run dev
```

Build TypeScript:

```bash
npm run build
```

Start the compiled server:

```bash
npm start
```

## Common Issues

### "Missing required environment variable"
Check that all required `.env` variables are set correctly.

### "Unauthorized" or "Invalid credentials"
- Verify your Client ID, Client Secret, and Refresh Token are correct
- Check that your Azure app registration has the required permissions
- Ensure the refresh token hasn't expired (refresh tokens expire if unused for 90 days)
- If calls fail with `invalid_grant`, the stored token died — see
  [Token lifecycle](#token-lifecycle) for the recovery steps

### "Permission denied"
Make sure the API permissions in your Azure app registration include:
- offline_access (without it no refresh token is issued at all)
- Mail.Read
- Mail.ReadWrite
- MailboxSettings.ReadWrite (categories and Inbox rules only; if this one is
  missing, mail tools keep working and only the category/rule tools fail —
  re-run the refresh-token step after adding the permission so the new token
  carries it)

## License

MIT
