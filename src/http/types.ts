/**
 * Outlook REST v2 resource shapes.
 *
 * These interfaces are the single source of truth for typing response bodies
 * from `https://outlook.office.com/api/v2.0/...`. The backend returns OData
 * with PascalCase property names; we keep the field names verbatim so JSON
 * pass-through round-trips losslessly.
 *
 * Normative source: project-design.md §3.2.
 *
 * NOTE: Additional fields not listed here may be present in upstream responses.
 * Commands treat the decoded objects as pass-through: unknown fields survive
 * into the JSON output.
 */

// ---------------------------------------------------------------------------
// Common sub-types
// ---------------------------------------------------------------------------

export interface EmailAddress {
  Name: string;
  Address: string;
}

export interface Recipient {
  EmailAddress: EmailAddress;
}

export interface Body {
  ContentType: 'HTML' | 'Text';
  Content: string;
}

// ---------------------------------------------------------------------------
// Mail
// ---------------------------------------------------------------------------

/** Shape returned by list-mail. */
export interface MessageSummary {
  Id: string;
  Subject: string;
  From?: Recipient;
  ReceivedDateTime: string;
  HasAttachments: boolean;
  IsRead: boolean;
  WebLink: string;
  /** Present when `$select` includes `ConversationId` (used by get-thread). */
  ConversationId?: string;
}

/**
 * Full message from GET /me/messages/{id}. Fields beyond the declared set may
 * be present and must be preserved by callers.
 */
export interface Message extends MessageSummary {
  Sender?: Recipient;
  ToRecipients: Recipient[];
  CcRecipients: Recipient[];
  BccRecipients: Recipient[];
  ReplyTo: Recipient[];
  Body?: Body;
  BodyPreview?: string;
  Importance?: 'Low' | 'Normal' | 'High';
  ConversationId?: string;
  InternetMessageId?: string;
  SentDateTime?: string;
  /** Added by get-mail via a separate request to /attachments. */
  Attachments?: AttachmentSummary[];
}

// ---------------------------------------------------------------------------
// Attachments — discriminated union by @odata.type
// ---------------------------------------------------------------------------

/**
 * Discriminator values observed on /api/v2.0 endpoints. These are the v2
 * OutlookServices namespace forms — NOT the Graph `#microsoft.graph.*` forms.
 */
export type AttachmentODataType =
  | '#Microsoft.OutlookServices.FileAttachment'
  | '#Microsoft.OutlookServices.ItemAttachment'
  | '#Microsoft.OutlookServices.ReferenceAttachment';

/** Fields common to every attachment subtype. */
export interface AttachmentBase {
  '@odata.type': AttachmentODataType;
  Id: string;
  Name: string;
  ContentType: string | null;
  Size: number;
  IsInline: boolean;
  LastModifiedDateTime: string;
}

export interface FileAttachment extends AttachmentBase {
  '@odata.type': '#Microsoft.OutlookServices.FileAttachment';
  ContentId: string | null;
  ContentLocation: string | null;
  /** base64; may be null on the list endpoint for large items. */
  ContentBytes: string | null;
}

export interface ItemAttachment extends AttachmentBase {
  '@odata.type': '#Microsoft.OutlookServices.ItemAttachment';
  /** Only populated with `$expand=Item`. */
  Item: unknown | null;
}

export interface ReferenceAttachment extends AttachmentBase {
  '@odata.type': '#Microsoft.OutlookServices.ReferenceAttachment';
  SourceUrl: string;
  ProviderType:
    | 'oneDriveBusiness'
    | 'oneDriveConsumer'
    | 'dropbox'
    | 'box'
    | 'google'
    | 'other';
  ThumbnailUrl: string | null;
  PreviewUrl: string | null;
  Permission: 'Edit' | 'View';
  IsFolder: boolean;
}

/** Discriminated union of the three concrete attachment shapes. */
export type Attachment = FileAttachment | ItemAttachment | ReferenceAttachment;

/** Back-compat alias matching project-design.md §3.2. */
export type AttachmentEnvelope = Attachment;

/** Subset returned by get-mail's $select query on /attachments. */
export interface AttachmentSummary {
  Id: string;
  Name: string;
  ContentType: string | null;
  Size: number;
  IsInline: boolean;
}

// ---------------------------------------------------------------------------
// Calendar
// ---------------------------------------------------------------------------

/** Timestamp wrapper used by calendar resources. */
export interface DateTimeWithTimeZone {
  DateTime: string;
  TimeZone: string;
}

/** Calendar event summary (list-calendar). */
export interface EventSummary {
  Id: string;
  Subject: string;
  Start: DateTimeWithTimeZone;
  End: DateTimeWithTimeZone;
  Organizer?: Recipient;
  Location?: { DisplayName: string };
  IsAllDay: boolean;
}

/** Full event (get-event). Additional fields pass through unchanged. */
export interface Event extends EventSummary {
  Body?: Body;
  Attendees?: Array<{
    EmailAddress: EmailAddress;
    Type: 'Required' | 'Optional' | 'Resource';
    Status?: { Response: string; Time: string };
  }>;
  BodyPreview?: string;
  ResponseRequested?: boolean;
  IsOnlineMeeting?: boolean;
  OnlineMeetingUrl?: string | null;
  WebLink?: string;
}

// ---------------------------------------------------------------------------
// Mail folders
// ---------------------------------------------------------------------------

/**
 * Shape returned by `GET /me/MailFolders`, `GET /me/MailFolders/{id}/childfolders`,
 * and `POST /me/MailFolders/{parent}/childfolders`.
 *
 * PascalCase matches the REST v2.0 wire convention used throughout this file.
 * Normative source: project-design.md §10.3.1.
 *
 * NOTE: `Path` is NOT returned by Outlook; it is materialized by the client
 * during a recursive `list-folders` walk and surfaced on the same shape so
 * consumers can render it verbatim. Slash-separated, with `/` and `\` escaped
 * per project-design.md §10.5.
 */
export interface FolderSummary {
  Id: string;
  DisplayName: string;
  ParentFolderId?: string;
  ChildFolderCount?: number;
  UnreadItemCount?: number;
  TotalItemCount?: number;
  /** Populated by Outlook only on well-known folders (e.g. "inbox"). */
  WellKnownName?: string;
  IsHidden?: boolean;
  /**
   * Selected explicitly by the resolver; required for `--first-match`
   * ordering (see project-design.md ADR-14).
   */
  CreatedDateTime?: string;
  /**
   * Not returned by Outlook; materialized by the client during a recursive
   * `list-folders` walk. Slash-separated, `/` and `\` escaped per §10.5.
   */
  Path?: string;
}

/** Request body for `POST /me/MailFolders/{parent}/childfolders`. */
export interface FolderCreateRequest {
  DisplayName: string;
}

/** Request body for `POST /me/messages/{id}/move`. */
export interface MoveMessageRequest {
  DestinationId: string;
}

// ---------------------------------------------------------------------------
// OData envelope
// ---------------------------------------------------------------------------

/**
 * Generic OData list envelope: `{ "value": T[] }`. Commands typically unwrap
 * `.value` before returning to callers.
 *
 * `@odata.count` is present only when the request includes `$count=true` AND
 * the server honors it. Used by list-mail's `--just-count` mode.
 */
export interface ODataListResponse<T> {
  '@odata.context'?: string;
  '@odata.nextLink'?: string;
  '@odata.count'?: number;
  value: T[];
}
