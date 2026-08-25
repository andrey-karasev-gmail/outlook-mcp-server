#!/usr/bin/env node
import "dotenv/config";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

import { getConfig } from "./config.js";
import { OutlookService } from "./service.js";

const config = getConfig();
const service = new OutlookService(config);

const server = new McpServer({
  name: "outlook-mcp-server",
  version: "1.0.0"
});

// List emails from Outlook inbox
server.tool(
  "outlook_list_emails",
  "List recent emails from Outlook mailbox",
  {
    limit: z.number().int().min(1).max(100).default(20),
    folder: z.string().default("Inbox").optional()
  },
  async ({ limit = 20, folder }) => {
    const emails = await service.listEmails(limit, folder);
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(emails, null, 2)
        }
      ]
    };
  }
);

// Get a single email by ID
server.tool(
  "outlook_get_email",
  "Read a single email by ID with full details",
  {
    id: z.string().min(1),
    markAsRead: z.boolean().default(false)
  },
  async ({ id, markAsRead = false }) => {
    const email = await service.getEmail(id, markAsRead);
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(email, null, 2)
        }
      ]
    };
  }
);

// Delete an email
server.tool(
  "outlook_delete_email",
  "Delete an email by ID",
  {
    id: z.string().min(1)
  },
  async ({ id }) => {
    await service.deleteEmail(id);
    return {
      content: [
        {
          type: "text",
          text: `Email ${id} deleted successfully`
        }
      ]
    };
  }
);

// Mark email as read/unread
server.tool(
  "outlook_mark_email",
  "Mark an email as read or unread",
  {
    id: z.string().min(1),
    isRead: z.boolean()
  },
  async ({ id, isRead }) => {
    await service.markAsRead(id, isRead);
    return {
      content: [
        {
          type: "text",
          text: `Email ${id} marked as ${isRead ? "read" : "unread"}`
        }
      ]
    };
  }
);

// Move email to folder
server.tool(
  "outlook_move_email",
  "Move an email to a specific folder",
  {
    id: z.string().min(1),
    destinationFolder: z.string().min(1)
  },
  async ({ id, destinationFolder }) => {
    await service.moveEmail(id, destinationFolder);
    return {
      content: [
        {
          type: "text",
          text: `Email ${id} moved to ${destinationFolder}`
        }
      ]
    };
  }
);

// List mail folders
server.tool(
  "outlook_list_folders",
  "List all mail folders in the mailbox",
  {},
  async () => {
    const folders = await service.listFolders();
    return {
      content: [{ type: "text", text: JSON.stringify(folders, null, 2) }]
    };
  }
);

// Get email attachments
server.tool(
  "outlook_get_attachments",
  "Get list of attachments for an email",
  {
    messageId: z.string().min(1)
  },
  async ({ messageId }) => {
    const attachments = await service.getAttachments(messageId);
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(attachments, null, 2)
        }
      ]
    };
  }
);

// Get attachment content
server.tool(
  "outlook_get_attachment_content",
  "Get the content of a specific attachment (base64 encoded)",
  {
    messageId: z.string().min(1),
    attachmentId: z.string().min(1)
  },
  async ({ messageId, attachmentId }) => {
    const attachment = await service.getAttachmentContent(messageId, attachmentId);
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(attachment, null, 2)
        }
      ]
    };
  }
);

// Start server
const transport = new StdioServerTransport();
await server.connect(transport);
