import type { Message } from "@microsoft/microsoft-graph-types";
import { FILE_TO_FOLDER_PROPERTY_ID } from "../config";
import { graphClient } from "./graphClient";

/* global Blob, FileReader, fetch */

/** Graph accepts attachments up to 3 MB in one request; larger ones need an upload session. */
const INLINE_ATTACHMENT_LIMIT = 3 * 1024 * 1024;
/** Upload session slice: under 4 MB per request, multiple of 320 KiB. */
const UPLOAD_SLICE = 9 * 320 * 1024;

/** What the new Outlook draft contains (attachments are added separately). */
export interface IDraftRequest {
  subject: string;
  to: string[];
  cc: string[];
  /** Plain text; Outlook shows it as the message body. */
  message: string;
  categories: string[];
  /** Folder the sent email should be filed to, as "driveId|itemId"; omitted when not wanted. */
  fileToFolder: string | undefined;
}

export interface IDraft {
  id: string;
  webLink: string;
}

const toRecipients = (addresses: string[]): Message["toRecipients"] =>
  addresses.map((address) => ({ emailAddress: { address } }));

function toMessage(request: IDraftRequest): Message {
  return {
    subject: request.subject,
    body: { contentType: "text", content: request.message },
    toRecipients: toRecipients(request.to),
    ccRecipients: toRecipients(request.cc),
    categories: request.categories,
    singleValueExtendedProperties: request.fileToFolder
      ? [{ id: FILE_TO_FOLDER_PROPERTY_ID, value: request.fileToFolder }]
      : undefined,
  };
}

/** Creates a draft in the signed-in user's mailbox (delegated Mail.ReadWrite). Nothing is sent. */
export async function createDraft(request: IDraftRequest): Promise<IDraft> {
  const created = (await graphClient.api("/me/messages").post(toMessage(request))) as Message;
  return { id: created.id!, webLink: created.webLink ?? "" };
}

/** Rewrites an existing draft's fields (used when a retry finishes a draft whose attachments failed). */
export async function updateDraft(id: string, request: IDraftRequest): Promise<void> {
  await graphClient.api(`/me/messages/${id}`).patch(toMessage(request));
}

/** Base64 of a Blob without building one huge binary string (FileReader data URL). */
function toBase64(body: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).replace(/^data:[^,]*,/, ""));
    reader.onerror = () => reject(reader.error ?? new Error("Could not read the attachment."));
    reader.readAsDataURL(body);
  });
}

/** Adds one file attachment to a draft: inline up to 3 MB, upload session above. */
export async function attachFile(messageId: string, name: string, body: Blob): Promise<void> {
  if (body.size <= INLINE_ATTACHMENT_LIMIT) {
    await graphClient.api(`/me/messages/${messageId}/attachments`).post({
      "@odata.type": "#microsoft.graph.fileAttachment",
      name,
      contentBytes: await toBase64(body),
    });
    return;
  }
  const session = (await graphClient
    .api(`/me/messages/${messageId}/attachments/createUploadSession`)
    .post({ AttachmentItem: { attachmentType: "file", name, size: body.size } })) as {
    uploadUrl: string;
  };
  // Plain fetch, as in the SPFx MailDraftService: Outlook ends with an empty 201 that graph-client's
  // LargeFileUploadTask may try to parse as JSON. The upload URL is pre-authorised: no bearer token.
  for (let start = 0; start < body.size; start += UPLOAD_SLICE) {
    const end = Math.min(start + UPLOAD_SLICE, body.size);
    const response = await fetch(session.uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Range": `bytes ${start}-${end - 1}/${body.size}`,
        "Content-Type": "application/octet-stream",
      },
      body: body.slice(start, end),
    });
    if (!response.ok)
      throw new Error(`Attachment upload failed for ${name} (HTTP ${response.status}).`);
  }
}
