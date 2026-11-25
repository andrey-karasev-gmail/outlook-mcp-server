export interface EmailSummary {
  id: string;
  subject: string;
  from: string;
  fromAddress?: string;
  date: string | null;
  isRead: boolean;
  snippet: string;
  receivedDateTime: string;
}

export interface EmailDetail extends EmailSummary {
  to: string;
  toAddresses?: string[];
  cc?: string;
  ccAddresses?: string[];
  bcc?: string;
  bccAddresses?: string[];
  bodyPreview: string;
  body: string;
  bodyType: "text" | "html";
  isReminderOn?: boolean;
  isFlagged?: boolean;
  importance: "low" | "normal" | "high";
  hasAttachments: boolean;
}

export interface AttachmentSummary {
  id: string;
  name: string;
  size: number;
  mimeType: string;
}

export type FlagStatus = "notFlagged" | "flagged" | "complete";

/** Follow-up flag update for a message. Dates are ISO 8601 (`2026-11-15` or
 * `2026-11-15T09:00:00`); `timeZone` is a Windows/IANA zone name, default UTC.
 * Dates only apply while the status is `flagged`. */
export interface FlagUpdate {
  status: FlagStatus;
  startDateTime?: string;
  dueDateTime?: string;
  timeZone?: string;
}

export interface MailCategory {
  id: string;
  displayName: string;
  color: string;
}

/** Category change for a message: `set` replaces the whole list; `add` and
 * `remove` are merged into the message's current categories instead. */
export interface CategorizeUpdate {
  set?: string[];
  add?: string[];
  remove?: string[];
}

export type ImportanceLevel = "low" | "normal" | "high";

export interface RuleConditionsInput {
  fromAddresses?: string[];
  senderContains?: string[];
  subjectContains?: string[];
  hasAttachments?: boolean;
  importance?: ImportanceLevel;
}

export interface RuleActionsInput {
  /** Folder display name or ID (resolved like outlook_move_email does). */
  moveToFolder?: string;
  markAsRead?: boolean;
  markImportance?: ImportanceLevel;
  assignCategories?: string[];
  delete?: boolean;
  stopProcessingRules?: boolean;
}

export interface CreateRuleInput {
  displayName: string;
  isEnabled?: boolean;
  sequence?: number;
  conditions?: RuleConditionsInput;
  actions: RuleActionsInput;
}

export interface MailRule {
  id: string;
  displayName: string;
  sequence: number;
  isEnabled: boolean;
  conditions?: Record<string, unknown>;
  actions?: Record<string, unknown>;
}

export interface EmailAttachment extends AttachmentSummary {
  contentBytes?: string;
}

/** A file to attach to a draft: `contentBytes` is the file's base64 content.
 * Attachments are created inline with the draft, so the whole request
 * (body + all attachments) must stay under Graph's ~3 MB limit; larger
 * files need an upload session instead. */
export interface DraftAttachmentInput {
  name: string;
  contentType: string;
  contentBytes: string;
}

/** Input for creating a new draft message (never sent). */
export interface CreateDraftInput {
  subject: string;
  toRecipients: string[];
  ccRecipients?: string[];
  body: string;
  attachments?: DraftAttachmentInput[];
}

/** A draft message created in Drafts. */
export interface DraftSummary {
  id: string;
  subject: string;
  toAddresses: string[];
  conversationId?: string;
  isDraft: boolean;
}
