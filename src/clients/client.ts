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

export interface EmailClient {
  listEmails(limit: number, folder?: string): Promise<EmailSummary[]>;
  getEmail(id: string, markAsRead?: boolean): Promise<EmailDetail>;
  deleteEmail(id: string): Promise<void>;
  markAsRead(id: string, isRead: boolean): Promise<void>;
  moveEmail(id: string, destinationFolder: string): Promise<void>;
  listFolders(): Promise<{ id: string; displayName: string }[]>;
  getAttachments(id: string): Promise<AttachmentSummary[]>;
  getAttachmentContent(messageId: string, attachmentId: string): Promise<EmailAttachment>;
  flagEmail(id: string, flag: FlagUpdate): Promise<void>;
  listCategories(): Promise<MailCategory[]>;
  createCategory(displayName: string, color: string): Promise<MailCategory>;
  categorizeEmail(id: string, update: CategorizeUpdate): Promise<string[]>;
  listRules(): Promise<MailRule[]>;
  createRule(rule: CreateRuleInput): Promise<MailRule>;
  deleteRule(ruleId: string): Promise<void>;
}

export interface EmailAttachment extends AttachmentSummary {
  contentBytes?: string;
}
