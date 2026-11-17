import type { MSGraphClientV3 } from '@microsoft/sp-http-msgraph';
import type { IFilePayload } from '../models/ISend';

/** Graph accepts attachments up to 3 MB in one request; larger ones need an upload session. */
const INLINE_ATTACHMENT_LIMIT: number = 3 * 1024 * 1024;
/** Upload session chunk size (must stay under 4 MB per request). */
const UPLOAD_CHUNK_SIZE: number = 3 * 1024 * 1024;
/**
 * Exchange Online default maximum message size is 35 MB (raised per mailbox up to 150 MB). Attachments that
 * would push the draft past this go in as links instead, so the draft can always be sent.
 */
export const MAX_DRAFT_ATTACHMENT_BYTES: number = 35 * 1024 * 1024;

function toBase64(bytes: ArrayBuffer): string {
  const view: Uint8Array = new Uint8Array(bytes);
  const chunk: number = 0x8000;
  let binary: string = '';
  for (let i = 0; i < view.length; i += chunk) {
    binary += String.fromCharCode.apply(undefined, Array.from(view.subarray(i, i + chunk)));
  }
  return btoa(binary);
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** A document sent as a link (over the size limit) rather than an attachment. */
export interface IDraftLink {
  name: string;
  url: string;
}

/** Outlook draft in the signed-in user's mailbox (Graph, delegated Mail.ReadWrite). Nothing is sent. */
export class MailDraftService {
  constructor(private readonly _graph: MSGraphClientV3) {}

  public async createDraft(subject: string): Promise<{ id: string; webLink: string }> {
    const message: { id: string; webLink: string } = await this._graph.api('/me/messages')
      .post({ subject, body: { contentType: 'HTML', content: '' } });
    return { id: message.id, webLink: message.webLink };
  }

  public async attach(messageId: string, file: IFilePayload): Promise<void> {
    if (file.bytes.byteLength <= INLINE_ATTACHMENT_LIMIT) {
      await this._graph.api(`/me/messages/${messageId}/attachments`).post({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: file.name,
        contentBytes: toBase64(file.bytes)
      });
      return;
    }
    const session: { uploadUrl: string } = await this._graph.api(`/me/messages/${messageId}/attachments/createUploadSession`)
      .post({ AttachmentItem: { attachmentType: 'file', name: file.name, size: file.bytes.byteLength } });
    const total: number = file.bytes.byteLength;
    for (let start = 0; start < total; start += UPLOAD_CHUNK_SIZE) {
      const end: number = Math.min(start + UPLOAD_CHUNK_SIZE, total);
      // The upload URL is pre-authorised; it must not carry the Graph bearer token.
      const response: Response = await fetch(session.uploadUrl, {
        method: 'PUT',
        headers: { 'Content-Range': `bytes ${start}-${end - 1}/${total}`, 'Content-Type': 'application/octet-stream' },
        body: file.bytes.slice(start, end)
      });
      if (!response.ok) {
        throw new Error(`Attachment upload failed for ${file.name} (HTTP ${response.status}).`);
      }
    }
  }

  /** Writes the body: an intro line plus links for documents too large to attach. */
  public async setBody(messageId: string, intro: string, linksHeading: string, links: ReadonlyArray<IDraftLink>): Promise<void> {
    const linkHtml: string = links.length
      ? `<p>${escapeHtml(linksHeading)}</p><ul>${links.map(l => `<li><a href="${escapeHtml(l.url)}">${escapeHtml(l.name)}</a></li>`).join('')}</ul>`
      : '';
    await this._graph.api(`/me/messages/${messageId}`).patch({ body: { contentType: 'HTML', content: `<p>${escapeHtml(intro)}</p>${linkHtml}` } });
  }
}
