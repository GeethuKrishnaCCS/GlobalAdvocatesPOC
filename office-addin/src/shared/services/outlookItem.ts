import { FILE_TO_FOLDER_PROPERTY_ID } from "../config";
import { graphClient } from "./graphClient";
import {
  base64ToBlob,
  parseFilingTarget,
  type FilingTarget,
  type IFilingAttachment,
  type IMailAttachment,
} from "./mailFiling";

/* global Office, Blob */

/** Promise over an Office.js callback API. */
function officeCall<T>(
  start: (callback: (result: Office.AsyncResult<T>) => void) => void
): Promise<T> {
  return new Promise((resolve, reject) =>
    start((result) =>
      result.status === Office.AsyncResultStatus.Succeeded
        ? resolve(result.value)
        : reject(new Error(result.error?.message ?? "Outlook could not complete the request."))
    )
  );
}

/** The email open now; capture it before async work, a pinned pane can move on to another email meanwhile. */
export function readItem(): Office.MessageRead {
  return Office.context.mailbox.item as Office.MessageRead;
}

function composeItem(): Office.MessageCompose {
  return Office.context.mailbox.item as Office.MessageCompose;
}

/** The open email (read mode) as the Items step needs it. */
export interface IOpenEmail {
  subject: string;
  received: Date;
  attachments: IMailAttachment[];
}

export function getOpenEmail(): IOpenEmail {
  const item = readItem();
  return {
    subject: item.subject ?? "",
    received: item.dateTimeCreated ?? new Date(),
    attachments: (item.attachments ?? []).map((a) => ({
      id: a.id,
      name: a.name,
      size: a.size,
      attachmentType: a.attachmentType,
      isInline: a.isInline,
    })),
  };
}

/** True when this Outlook can hand over the whole email as EML (Mailbox 1.14). */
export function canGetEmail(): boolean {
  return Office.context.requirements.isSetSupported("Mailbox", "1.14");
}

/** The open email as an .eml file. */
export async function getEmailFile(item: Office.MessageRead): Promise<Blob> {
  return base64ToBlob(await officeCall<string>((cb) => item.getAsFileAsync(cb)), "message/rfc822");
}

/** Content of one attachment of the open email, as a file. */
export async function getAttachmentFile(
  item: Office.MessageRead,
  attachment: IFilingAttachment
): Promise<Blob> {
  const content = await officeCall<Office.AttachmentContent>((cb) =>
    item.getAttachmentContentAsync(attachment.id, cb)
  );
  switch (content.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      return base64ToBlob(content.content, "application/octet-stream");
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      return new Blob([content.content], { type: "message/rfc822" });
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      return new Blob([content.content], { type: "text/calendar" });
    default:
      throw new Error(`${attachment.fileName} is shared as a link and cannot be filed.`);
  }
}

interface IGraphMessage {
  categories?: string[];
  singleValueExtendedProperties?: { id: string; value: string }[];
}

/**
 * Where the open email asks to be filed: the folder marker Word Send puts on its draft, else a `Matter: <ref>`
 * category. Read through Graph (Office.js cannot read extended properties it did not write).
 */
export async function getFilingTarget(): Promise<FilingTarget | undefined> {
  const mailbox = Office.context.mailbox;
  const restId = mailbox.convertToRestId(readItem().itemId, Office.MailboxEnums.RestVersion.v2_0);
  const message = (await graphClient
    .api(`/me/messages/${restId}`)
    .select("categories")
    .expand(`singleValueExtendedProperties($filter=id eq '${FILE_TO_FOLDER_PROPERTY_ID}')`)
    .get()) as IGraphMessage;
  return parseFilingTarget(message.singleValueExtendedProperties?.[0]?.value, message.categories);
}

/** Names of the files already attached to the email being composed. */
export async function getAttachedNames(): Promise<string[]> {
  const attachments = await officeCall<Office.AttachmentDetailsCompose[]>((cb) =>
    composeItem().getAttachmentsAsync(cb)
  );
  return attachments.map((a) => a.name);
}

/** Attaches a file (base64) to the email being composed. */
export async function attachToEmail(fileName: string, base64: string): Promise<void> {
  await officeCall<string>((cb) =>
    composeItem().addFileAttachmentFromBase64Async(base64, fileName, cb)
  );
}
