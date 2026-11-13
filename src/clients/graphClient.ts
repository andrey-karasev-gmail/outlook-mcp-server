import { Client } from "@microsoft/microsoft-graph-client";
import {
  TokenManager,
  defaultTokenStatePath,
  isScopeConsentError,
  type TokenRequestFn
} from "./tokenManager.js";
import type {
  EmailDetail,
  EmailSummary,
  AttachmentSummary,
  FlagUpdate,
  MailCategory,
  CategorizeUpdate,
  CreateRuleInput,
  MailRule
} from "../types/email.js";
import type { EmailClient, EmailAttachment } from "./client.js";
import type { OutlookConfig } from "../config.js";

/** Convert an ISO date/datetime into Graph's dateTimeTimeZone shape: the
 * datetime itself must be zone-less, the zone travels in its own field. */
function toDateTimeTimeZone(
  value: string,
  timeZone?: string
): { dateTime: string; timeZone: string } {
  let dateTime = value.trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateTime)) {
    dateTime = `${dateTime}T00:00:00`;
  } else {
    dateTime = dateTime.replace(/(Z|[+-]\d{2}:?\d{2})$/, "");
  }
  return { dateTime, timeZone: timeZone || "UTC" };
}

export class GraphEmailClient implements EmailClient {
  private graphClient: Client;
  private readonly config: OutlookConfig;
  private readonly tokenManager: TokenManager;

  constructor(config: OutlookConfig, tokenRequestFn?: TokenRequestFn) {
    this.config = config;

    // The effective refresh token may come from the token state file
    // (freshest, rotated) even when .env carries no seed; the missing-
    // token error below only fires when neither source has one.
    this.tokenManager = new TokenManager({
      clientId: config.clientId,
      clientSecret: config.clientSecret,
      tenantId: config.tenantId,
      seedRefreshToken: config.refreshToken,
      stateFilePath: config.tokenStatePath ?? defaultTokenStatePath(),
      requestFn: tokenRequestFn
    });

    if (!this.tokenManager.hasRefreshToken()) {
      throw new Error(
        "Missing OUTLOOK_REFRESH_TOKEN in environment. Delegated authentication requires a refresh token."
      );
    }

    this.graphClient = Client.init({
      authProvider: async (done) => {
        try {
          const accessToken = await this.getAccessToken();
          done(null, accessToken);
        } catch (error) {
          done(error, null);
        }
      }
    });
  }

  private scopeMode: "extended" | "base" | null = null;

  private async getAccessToken(): Promise<string> {
    const baseScope = "offline_access mail.read mail.readwrite";
    if (this.scopeMode === "base") {
      return this.tokenManager.getAccessToken(baseScope);
    }
    try {
      const token = await this.tokenManager.getAccessToken(
        `${baseScope} mailboxsettings.readwrite`
      );
      this.scopeMode = "extended";
      return token;
    } catch (error) {
      if (!isScopeConsentError(error)) {
        // Transient failures (network, 429, 5xx) and dead tokens surface
        // as errors; they must not silently downgrade the scope set.
        throw error;
      }
      // The refresh token predates the MailboxSettings consent (or the app
      // registration does not grant it): fall back to the original scope so
      // the mail tools keep working. Category and rule calls will surface a
      // Graph 403 until a token is re-minted with the wider scope.
      this.scopeMode = "base";
      return this.tokenManager.getAccessToken(baseScope);
    }
  }

  async listEmails(limit: number, folder: string = "inbox"): Promise<EmailSummary[]> {
    try {
      const messages = await this.graphClient
        .api(`/me/mailFolders/${folder}/messages`)
        .select(
          "id,subject,from,receivedDateTime,isRead,bodyPreview,webLink"
        )
        .orderby("receivedDateTime desc")
        .top(Math.min(limit, 999))
        .get();

      return messages.value.map((msg: any) => this.mapToEmailSummary(msg));
    } catch (error) {
      throw new Error(`Failed to list emails: ${error}`);
    }
  }

  async getEmail(id: string, markAsRead: boolean = false): Promise<EmailDetail> {
    try {
      const message = await this.graphClient
        .api(`/me/messages/${id}`)
        .select(
          "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,isRead,bodyPreview,body,importance,hasAttachments,webLink,flag"
        )
        .get();

      if (markAsRead && !message.isRead) {
        await this.markAsRead(id, true);
      }

      return this.mapToEmailDetail(message);
    } catch (error) {
      throw new Error(`Failed to get email ${id}: ${error}`);
    }
  }

  async deleteEmail(id: string): Promise<void> {
    try {
      await this.graphClient.api(`/me/messages/${id}`).delete();
    } catch (error) {
      throw new Error(`Failed to delete email ${id}: ${error}`);
    }
  }

  async markAsRead(id: string, isRead: boolean): Promise<void> {
    try {
      await this.graphClient
        .api(`/me/messages/${id}`)
        .patch({ isRead });
    } catch (error) {
      throw new Error(`Failed to mark email ${id} as read: ${error}`);
    }
  }

  async listFolders(): Promise<{ id: string; displayName: string }[]> {
    try {
      const result = await this.graphClient
        .api("/me/mailFolders")
        .select("id,displayName")
        .top(50)
        .get();
      return result.value.map((f: any) => ({ id: f.id, displayName: f.displayName }));
    } catch (error) {
      throw new Error(`Failed to list folders: ${error}`);
    }
  }

  /** Resolve a folder display name to its ID, creating the folder if it does
   * not exist yet. Values that already look like a Graph folder ID pass
   * through unchanged. */
  private async resolveFolderId(destinationFolder: string): Promise<string> {
    if (destinationFolder.match(/^[A-Za-z0-9_=-]{20,}$/)) {
      return destinationFolder;
    }
    const folders = await this.listFolders();
    const match = folders.find(
      (f) => f.displayName.toLowerCase() === destinationFolder.toLowerCase()
    );
    if (match) {
      return match.id;
    }
    const created = await this.graphClient
      .api("/me/mailFolders")
      .post({ displayName: destinationFolder });
    return created.id;
  }

  async moveEmail(id: string, destinationFolder: string): Promise<void> {
    try {
      const folderId = await this.resolveFolderId(destinationFolder);
      await this.graphClient.api(`/me/messages/${id}/move`).post({ destinationId: folderId });
    } catch (error) {
      throw new Error(`Failed to move email ${id}: ${error}`);
    }
  }

  async flagEmail(id: string, flag: FlagUpdate): Promise<void> {
    try {
      const followupFlag: Record<string, unknown> = { flagStatus: flag.status };
      if (flag.status === "flagged") {
        if (flag.startDateTime) {
          followupFlag.startDateTime = toDateTimeTimeZone(flag.startDateTime, flag.timeZone);
        }
        if (flag.dueDateTime) {
          followupFlag.dueDateTime = toDateTimeTimeZone(flag.dueDateTime, flag.timeZone);
        }
      }
      await this.graphClient.api(`/me/messages/${id}`).patch({ flag: followupFlag });
    } catch (error) {
      throw new Error(`Failed to flag email ${id}: ${error}`);
    }
  }

  async listCategories(): Promise<MailCategory[]> {
    try {
      const result = await this.graphClient.api("/me/outlook/masterCategories").get();
      return result.value.map((c: any) => ({
        id: c.id,
        displayName: c.displayName,
        color: c.color
      }));
    } catch (error) {
      throw new Error(`Failed to list categories: ${error}`);
    }
  }

  async createCategory(displayName: string, color: string): Promise<MailCategory> {
    try {
      const created = await this.graphClient
        .api("/me/outlook/masterCategories")
        .post({ displayName, color });
      return { id: created.id, displayName: created.displayName, color: created.color };
    } catch (error) {
      throw new Error(`Failed to create category ${displayName}: ${error}`);
    }
  }

  async categorizeEmail(id: string, update: CategorizeUpdate): Promise<string[]> {
    try {
      let categories: string[];
      if (update.set) {
        categories = [...new Set(update.set)];
      } else {
        const message = await this.graphClient
          .api(`/me/messages/${id}`)
          .select("categories")
          .get();
        categories = [...((message.categories as string[] | undefined) ?? [])];
        for (const name of update.add ?? []) {
          if (!categories.some((c) => c.toLowerCase() === name.toLowerCase())) {
            categories.push(name);
          }
        }
        if (update.remove?.length) {
          const removing = new Set(update.remove.map((n) => n.toLowerCase()));
          categories = categories.filter((c) => !removing.has(c.toLowerCase()));
        }
      }
      await this.graphClient.api(`/me/messages/${id}`).patch({ categories });
      return categories;
    } catch (error) {
      throw new Error(`Failed to categorize email ${id}: ${error}`);
    }
  }

  async listRules(): Promise<MailRule[]> {
    try {
      const result = await this.graphClient.api("/me/mailFolders/inbox/messageRules").get();
      return result.value.map((r: any) => this.mapToMailRule(r));
    } catch (error) {
      throw new Error(`Failed to list inbox rules: ${error}`);
    }
  }

  async createRule(rule: CreateRuleInput): Promise<MailRule> {
    try {
      if (!rule.actions || Object.keys(rule.actions).length === 0) {
        throw new Error("A rule needs at least one action");
      }

      const conditions: Record<string, unknown> = {};
      if (rule.conditions?.fromAddresses?.length) {
        conditions.fromAddresses = rule.conditions.fromAddresses.map((address) => ({
          emailAddress: { address }
        }));
      }
      if (rule.conditions?.senderContains?.length) {
        conditions.senderContains = rule.conditions.senderContains;
      }
      if (rule.conditions?.subjectContains?.length) {
        conditions.subjectContains = rule.conditions.subjectContains;
      }
      if (rule.conditions?.hasAttachments !== undefined) {
        conditions.hasAttachments = rule.conditions.hasAttachments;
      }
      if (rule.conditions?.importance) {
        conditions.importance = rule.conditions.importance;
      }

      const actions: Record<string, unknown> = {};
      if (rule.actions.moveToFolder) {
        actions.moveToFolder = await this.resolveFolderId(rule.actions.moveToFolder);
      }
      if (rule.actions.markAsRead !== undefined) {
        actions.markAsRead = rule.actions.markAsRead;
      }
      if (rule.actions.markImportance) {
        actions.markImportance = rule.actions.markImportance;
      }
      if (rule.actions.assignCategories?.length) {
        actions.assignCategories = rule.actions.assignCategories;
      }
      if (rule.actions.delete !== undefined) {
        actions.delete = rule.actions.delete;
      }
      if (rule.actions.stopProcessingRules !== undefined) {
        actions.stopProcessingRules = rule.actions.stopProcessingRules;
      }
      if (Object.keys(actions).length === 0) {
        throw new Error("A rule needs at least one action");
      }

      const created = await this.graphClient
        .api("/me/mailFolders/inbox/messageRules")
        .post({
          displayName: rule.displayName,
          sequence: rule.sequence ?? 1,
          isEnabled: rule.isEnabled ?? true,
          conditions,
          actions
        });
      return this.mapToMailRule(created);
    } catch (error) {
      throw new Error(`Failed to create inbox rule ${rule.displayName}: ${error}`);
    }
  }

  async deleteRule(ruleId: string): Promise<void> {
    try {
      await this.graphClient.api(`/me/mailFolders/inbox/messageRules/${ruleId}`).delete();
    } catch (error) {
      throw new Error(`Failed to delete inbox rule ${ruleId}: ${error}`);
    }
  }

  async getAttachments(messageId: string): Promise<AttachmentSummary[]> {
    try {
      const attachments = await this.graphClient
        .api(`/me/messages/${messageId}/attachments`)
        .select("id,name,size,@odata.type,mimeType")
        .get();

      return attachments.value
        .filter((att: any) => att["@odata.type"] !== "#microsoft.graph.itemAttachment")
        .map((att: any) => ({
          id: att.id,
          name: att.name,
          size: att.size || 0,
          mimeType: att.mimeType || "application/octet-stream"
        }));
    } catch (error) {
      throw new Error(
        `Failed to get attachments for message ${messageId}: ${error}`
      );
    }
  }

  async getAttachmentContent(
    messageId: string,
    attachmentId: string
  ): Promise<EmailAttachment> {
    try {
      const attachment = await this.graphClient
        .api(`/me/messages/${messageId}/attachments/${attachmentId}`)
        .select("id,name,size,@odata.type,mimeType,contentBytes")
        .get();

      return {
        id: attachment.id,
        name: attachment.name,
        size: attachment.size || 0,
        mimeType: attachment.mimeType || "application/octet-stream",
        contentBytes: attachment.contentBytes
      };
    } catch (error) {
      throw new Error(
        `Failed to get attachment ${attachmentId} content: ${error}`
      );
    }
  }

  private mapToMailRule(rule: any): MailRule {
    return {
      id: rule.id,
      displayName: rule.displayName,
      sequence: rule.sequence,
      isEnabled: rule.isEnabled,
      conditions: rule.conditions,
      actions: rule.actions
    };
  }

  private mapToEmailSummary(msg: any): EmailSummary {
    return {
      id: msg.id,
      subject: msg.subject || "(no subject)",
      from: msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || "Unknown",
      fromAddress: msg.from?.emailAddress?.address,
      date: msg.receivedDateTime,
      isRead: msg.isRead,
      snippet: msg.bodyPreview || "",
      receivedDateTime: msg.receivedDateTime
    };
  }

  private mapToEmailDetail(msg: any): EmailDetail {
    const toAddresses = msg.toRecipients?.map((r: any) => r.emailAddress?.address).filter(Boolean) || [];
    const ccAddresses = msg.ccRecipients?.map((r: any) => r.emailAddress?.address).filter(Boolean) || [];
    const bccAddresses = msg.bccRecipients?.map((r: any) => r.emailAddress?.address).filter(Boolean) || [];

    return {
      id: msg.id,
      subject: msg.subject || "(no subject)",
      from: msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || "Unknown",
      fromAddress: msg.from?.emailAddress?.address,
      to: toAddresses.join(", "),
      toAddresses,
      cc: ccAddresses.length > 0 ? ccAddresses.join(", ") : undefined,
      ccAddresses: ccAddresses.length > 0 ? ccAddresses : undefined,
      bcc: bccAddresses.length > 0 ? bccAddresses.join(", ") : undefined,
      bccAddresses: bccAddresses.length > 0 ? bccAddresses : undefined,
      date: msg.receivedDateTime,
      isRead: msg.isRead,
      bodyPreview: msg.bodyPreview || "",
      body: msg.body?.content || "",
      bodyType: msg.body?.contentType === "html" ? "html" : "text",
      snippet: msg.bodyPreview || "",
      receivedDateTime: msg.receivedDateTime,
      importance: msg.importance?.toLowerCase() || "normal",
      hasAttachments: msg.hasAttachments || false,
      isReminderOn: msg.isReminderOn,
      isFlagged: msg.flag?.flagStatus === "flagged"
    };
  }
}
