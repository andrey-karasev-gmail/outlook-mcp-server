import type { OutlookConfig } from "./config.js";
import type {
  EmailDetail,
  EmailSummary,
  AttachmentSummary,
  FlagUpdate,
  MailCategory,
  CategorizeUpdate,
  CreateRuleInput,
  MailRule
} from "./types/email.js";
import type { EmailClient } from "./clients/client.js";
import { GraphEmailClient } from "./clients/graphClient.js";
import type { TokenRequestFn } from "./clients/tokenManager.js";

export class OutlookService {
  private readonly graphClient: EmailClient;

  constructor(
    private readonly config: OutlookConfig,
    tokenRequestFn?: TokenRequestFn
  ) {
    this.graphClient = new GraphEmailClient(config, tokenRequestFn);
  }

  listEmails(limit: number, folder?: string): Promise<EmailSummary[]> {
    return this.graphClient.listEmails(
      Math.min(limit, this.config.batchSize),
      folder || this.config.mailbox
    );
  }

  getEmail(id: string, markAsRead: boolean = false): Promise<EmailDetail> {
    return this.graphClient.getEmail(id, markAsRead);
  }

  deleteEmail(id: string): Promise<void> {
    return this.graphClient.deleteEmail(id);
  }

  markAsRead(id: string, isRead: boolean): Promise<void> {
    return this.graphClient.markAsRead(id, isRead);
  }

  moveEmail(id: string, destinationFolder: string): Promise<void> {
    return this.graphClient.moveEmail(id, destinationFolder);
  }

  listFolders(): Promise<{ id: string; displayName: string }[]> {
    return this.graphClient.listFolders();
  }

  getAttachments(messageId: string): Promise<AttachmentSummary[]> {
    return this.graphClient.getAttachments(messageId);
  }

  getAttachmentContent(messageId: string, attachmentId: string) {
    return this.graphClient.getAttachmentContent(messageId, attachmentId);
  }

  flagEmail(id: string, flag: FlagUpdate): Promise<void> {
    return this.graphClient.flagEmail(id, flag);
  }

  listCategories(): Promise<MailCategory[]> {
    return this.graphClient.listCategories();
  }

  createCategory(displayName: string, color: string): Promise<MailCategory> {
    return this.graphClient.createCategory(displayName, color);
  }

  categorizeEmail(id: string, update: CategorizeUpdate): Promise<string[]> {
    return this.graphClient.categorizeEmail(id, update);
  }

  listRules(): Promise<MailRule[]> {
    return this.graphClient.listRules();
  }

  createRule(rule: CreateRuleInput): Promise<MailRule> {
    return this.graphClient.createRule(rule);
  }

  deleteRule(ruleId: string): Promise<void> {
    return this.graphClient.deleteRule(ruleId);
  }
}
