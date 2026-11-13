#!/usr/bin/env node
import "dotenv/config";
import http from "node:http";
import { timingSafeEqual } from "node:crypto";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { z } from "zod";

import { getConfig } from "./config.js";
import { OutlookService } from "./service.js";

const config = getConfig();
const service = new OutlookService(config);

function createOutlookServer(): McpServer {
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

  // Set, complete, or clear the follow-up flag on an email
  server.tool(
    "outlook_flag_email",
    "Set, complete, or clear the follow-up flag on an email. dueDateTime/startDateTime are ISO dates or datetimes and apply only when status is 'flagged'",
    {
      id: z.string().min(1),
      status: z.enum(["flagged", "complete", "notFlagged"]),
      dueDateTime: z.string().optional(),
      startDateTime: z.string().optional(),
      timeZone: z.string().optional()
    },
    async ({ id, status, dueDateTime, startDateTime, timeZone }) => {
      await service.flagEmail(id, { status, dueDateTime, startDateTime, timeZone });
      return {
        content: [
          {
            type: "text",
            text: `Email ${id} flag set to ${status}`
          }
        ]
      };
    }
  );

  // List categories from the master category list
  server.tool(
    "outlook_list_categories",
    "List all Outlook categories from the master category list",
    {},
    async () => {
      const categories = await service.listCategories();
      return {
        content: [{ type: "text", text: JSON.stringify(categories, null, 2) }]
      };
    }
  );

  // Create a new category in the master category list
  server.tool(
    "outlook_create_category",
    "Create a new Outlook category in the master category list",
    {
      displayName: z.string().min(1),
      color: z
        .enum([
          "none",
          "preset0", "preset1", "preset2", "preset3", "preset4",
          "preset5", "preset6", "preset7", "preset8", "preset9",
          "preset10", "preset11", "preset12", "preset13", "preset14",
          "preset15", "preset16", "preset17", "preset18", "preset19",
          "preset20", "preset21", "preset22", "preset23", "preset24"
        ])
        .default("preset0")
    },
    async ({ displayName, color }) => {
      const category = await service.createCategory(displayName, color);
      return {
        content: [{ type: "text", text: JSON.stringify(category, null, 2) }]
      };
    }
  );

  // Set, add, or remove categories on an email
  server.tool(
    "outlook_categorize_email",
    "Change the categories of an email: 'set' replaces the whole list, 'add'/'remove' merge with the current categories. Returns the resulting categories",
    {
      id: z.string().min(1),
      set: z.array(z.string()).optional(),
      add: z.array(z.string()).optional(),
      remove: z.array(z.string()).optional()
    },
    async ({ id, set, add, remove }) => {
      const categories = await service.categorizeEmail(id, { set, add, remove });
      return {
        content: [{ type: "text", text: JSON.stringify(categories, null, 2) }]
      };
    }
  );

  // List Inbox rules
  server.tool(
    "outlook_list_rules",
    "List all Inbox rules (server-side rules that act on incoming mail)",
    {},
    async () => {
      const rules = await service.listRules();
      return {
        content: [{ type: "text", text: JSON.stringify(rules, null, 2) }]
      };
    }
  );

  // Create an Inbox rule
  server.tool(
    "outlook_create_rule",
    "Create an Inbox rule. Conditions are optional (an empty condition set matches all mail); at least one action is required. moveToFolder accepts a folder display name or ID",
    {
      displayName: z.string().min(1),
      isEnabled: z.boolean().default(true),
      sequence: z.number().int().min(1).optional(),
      fromAddresses: z.array(z.string()).optional(),
      senderContains: z.array(z.string()).optional(),
      subjectContains: z.array(z.string()).optional(),
      hasAttachments: z.boolean().optional(),
      importance: z.enum(["low", "normal", "high"]).optional(),
      moveToFolder: z.string().optional(),
      markAsRead: z.boolean().optional(),
      markImportance: z.enum(["low", "normal", "high"]).optional(),
      assignCategories: z.array(z.string()).optional(),
      delete: z.boolean().optional(),
      stopProcessingRules: z.boolean().optional()
    },
    async (args) => {
      const rule = await service.createRule({
        displayName: args.displayName,
        isEnabled: args.isEnabled,
        sequence: args.sequence,
        conditions: {
          fromAddresses: args.fromAddresses,
          senderContains: args.senderContains,
          subjectContains: args.subjectContains,
          hasAttachments: args.hasAttachments,
          importance: args.importance
        },
        actions: {
          moveToFolder: args.moveToFolder,
          markAsRead: args.markAsRead,
          markImportance: args.markImportance,
          assignCategories: args.assignCategories,
          delete: args.delete,
          stopProcessingRules: args.stopProcessingRules
        }
      });
      return {
        content: [{ type: "text", text: JSON.stringify(rule, null, 2) }]
      };
    }
  );

  // Delete an Inbox rule
  server.tool(
    "outlook_delete_rule",
    "Delete an Inbox rule by its ID (see outlook_list_rules)",
    {
      ruleId: z.string().min(1)
    },
    async ({ ruleId }) => {
      await service.deleteRule(ruleId);
      return {
        content: [
          {
            type: "text",
            text: `Inbox rule ${ruleId} deleted successfully`
          }
        ]
      };
    }
  );

  return server;
}

async function readJsonBody(req: http.IncomingMessage): Promise<unknown> {
  const chunks: Buffer[] = [];
  for await (const chunk of req) {
    chunks.push(chunk as Buffer);
  }
  if (chunks.length === 0) {
    return undefined;
  }
  return JSON.parse(Buffer.concat(chunks).toString("utf8"));
}

// Start server
if (config.transport === "http") {
  if (!config.httpToken) {
    throw new Error("MCP_HTTP_TOKEN is required when MCP_TRANSPORT=http");
  }
  const expectedAuth = Buffer.from(`Bearer ${config.httpToken}`);

  const httpServer = http.createServer(async (req, res) => {
    try {
      const url = new URL(req.url ?? "/", "http://localhost");
      if (url.pathname !== "/mcp") {
        res.writeHead(404).end();
        return;
      }
      const providedAuth = Buffer.from(req.headers.authorization ?? "");
      if (
        providedAuth.length !== expectedAuth.length ||
        !timingSafeEqual(providedAuth, expectedAuth)
      ) {
        res.writeHead(401, { "WWW-Authenticate": "Bearer" }).end();
        return;
      }

      // Stateless mode: a fresh server+transport pair per request.
      const mcpServer = createOutlookServer();
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: undefined,
        enableJsonResponse: true
      });
      res.on("close", () => {
        void transport.close();
        void mcpServer.close();
      });
      await mcpServer.connect(transport);
      const body = req.method === "POST" ? await readJsonBody(req) : undefined;
      await transport.handleRequest(req, res, body);
    } catch {
      if (!res.headersSent) {
        res
          .writeHead(500, { "Content-Type": "application/json" })
          .end(
            JSON.stringify({
              jsonrpc: "2.0",
              error: { code: -32603, message: "Internal server error" },
              id: null
            })
          );
      }
    }
  });

  httpServer.listen(config.httpPort, () => {
    console.error(
      `outlook-mcp-server (http) listening on port ${config.httpPort}, path /mcp`
    );
  });
} else {
  const transport = new StdioServerTransport();
  await createOutlookServer().connect(transport);
}
