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

## Setup

### Prerequisites

- Node.js 18+ and npm
- An Outlook.com or Microsoft 365 account
- Microsoft Azure Application Registration

### 1. Register Azure Application

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Fill in the details:
   - Name: `Outlook MCP Server`
   - Supported account types: `Accounts in any organizational directory and personal Microsoft accounts`
5. Click **Register**

### 2. Create Client Secret

1. In your app registration, go to **Certificates & secrets**
2. Click **New client secret**
3. Add a description and select expiration
4. Copy the **Value** (you'll need this as `OUTLOOK_CLIENT_SECRET`)

### 3. Set API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission** > **Microsoft Graph**
3. Select **Delegated permissions**
4. Search and add these permissions:
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `MailboxSettings.ReadWrite` (for categories and Inbox rules; flags work with `Mail.ReadWrite` alone)
5. Click **Grant admin consent** (if you're in an organization)

### 4. Get Your Credentials

From your app registration overview page, copy:
- **Application (client) ID** → `OUTLOOK_CLIENT_ID`
- **Directory (tenant) ID** → `OUTLOOK_TENANT_ID` (or use "common")

### 5. Get Refresh Token (for user login)

You'll need to perform a one-time authorization flow. Use this script:

```bash
# Generate auth URL
node scripts/get-auth-token.js

# Or, with no cloning, using the published package:
npx --package=@web4w3/outlook-mcp-server outlook-mcp-server-get-auth-token

# Follow the link to login and copy the authorization code
# Then exchange it for a refresh token
```

Or use the Microsoft Graph explorer at https://developer.microsoft.com/en-us/graph/graph-explorer

### 6. Install

```bash
npx -y @web4w3/outlook-mcp-server
```

No cloning or build step needed. The package is pre-compiled and available on npm as [`@web4w3/outlook-mcp-server`](https://www.npmjs.com/package/@web4w3/outlook-mcp-server).

### 7. Configure in Claude Desktop

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

### "Permission denied"
Make sure the API permissions in your Azure app registration include:
- Mail.Read
- Mail.ReadWrite
- MailboxSettings.ReadWrite (categories and Inbox rules only; if this one is
  missing, mail tools keep working and only the category/rule tools fail —
  re-run the refresh-token step after adding the permission so the new token
  carries it)

## License

MIT
