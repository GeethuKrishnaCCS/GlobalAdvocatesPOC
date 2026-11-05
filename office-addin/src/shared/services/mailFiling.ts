import { MATTER_CATEGORY_PREFIX } from "../config";
import { MAX_FILE_NAME_LENGTH, cleanBaseName } from "./fileName";

/* global Blob, atob, btoa */

const EML = ".eml";
const NO_SUBJECT = "(no subject)";
/** Outlook's attachment limit per item (Exchange Online default send/receive limit is ~25 MB of encoded content). */
export const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;

/** Attachment as Office.js lists it (read mode `item.attachments`). */
export interface IMailAttachment {
  id: string;
  name: string;
  size: number;
  attachmentType: "file" | "item" | "cloud" | string;
  isInline: boolean;
}

/** One row of the Items step. */
export interface IFilingAttachment {
  id: string;
  fileName: string;
  size: number;
  kind: "file" | "item";
  fileable: boolean;
  /** Ticked by default: real attachments yes, inline images (signatures, logos) no. */
  ticked: boolean;
  reason?: string;
}

/** Where a filed email should go: an exact folder (Word Send marker) or just a matter (category). */
export type FilingTarget = { driveId: string; folderId: string } | { matterRef: string };

function datePrefix(date: Date): string {
  const pad = (n: number): string => String(n).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

/** "<yyyy-MM-dd> <subject>.eml", cleaned for SharePoint and within its name length. Pure. */
export function emlFileName(subject: string | undefined, received: Date): string {
  const prefix = `${datePrefix(received)} `;
  const base = cleanBaseName(subject ?? "") || NO_SUBJECT;
  return prefix + base.slice(0, MAX_FILE_NAME_LENGTH - prefix.length - EML.length).trimEnd() + EML;
}

/** Items step row for an attachment: what it is saved as and whether it can be filed at all. Pure. */
export function classifyAttachment(attachment: IMailAttachment): IFilingAttachment {
  const base = { id: attachment.id, size: attachment.size };
  if (attachment.attachmentType === "cloud") {
    return {
      ...base,
      fileName: attachment.name,
      kind: "file",
      fileable: false,
      ticked: false,
      reason: "Shared as a link, not fileable",
    };
  }
  if (attachment.attachmentType === "item") {
    // Attached emails come back as EML content; their name is the subject, without extension.
    const name = cleanBaseName(attachment.name) || NO_SUBJECT;
    const fileName = name.toLowerCase().endsWith(EML) ? name : name + EML;
    return { ...base, fileName, kind: "item", fileable: true, ticked: true };
  }
  return {
    ...base,
    fileName: attachment.name,
    kind: "file",
    fileable: true,
    ticked: !attachment.isInline,
  };
}

/**
 * Filing target from the Word Send marker (`AdvFileToFolder` = "driveId|itemId") or else a `Matter: <ref>`
 * category; undefined when the email carries neither. Pure.
 */
export function parseFilingTarget(
  folderProperty: string | undefined,
  categories: string[] | undefined
): FilingTarget | undefined {
  const [driveId, folderId, ...rest] = (folderProperty ?? "").split("|");
  if (driveId && folderId && !rest.length) return { driveId, folderId };
  const category = categories?.find((c) => c.startsWith(MATTER_CATEGORY_PREFIX));
  const matterRef = category?.slice(MATTER_CATEGORY_PREFIX.length).trim();
  return matterRef ? { matterRef } : undefined;
}

/** True when the email already has an attachment of this name (Outlook keeps no link to the source file). Pure. */
export function isAlreadyAttached(fileName: string, attachedNames: string[]): boolean {
  const name = fileName.toLowerCase();
  return attachedNames.some((n) => n.toLowerCase() === name);
}

/** Why a document cannot be attached to an email, or undefined when it can. Pure. */
export function attachProblem(fileName: string, size: number): string | undefined {
  return size > MAX_ATTACHMENT_BYTES
    ? `${fileName} is too large for an Outlook attachment (over 25 MB). Share a link instead.`
    : undefined;
}

/** Blob from base64 content (Office.js attachment / EML APIs return base64). */
export function base64ToBlob(base64: string, type: string): Blob {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new Blob([bytes], { type });
}

/** Bytes per String.fromCharCode call; spreading a whole file overflows the call stack. */
const CHUNK = 0x8000;

/** Base64 of downloaded bytes, for Office.js addFileAttachmentFromBase64Async. */
export function bytesToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (let i = 0; i < bytes.length; i += CHUNK)
    binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
  return btoa(binary);
}

/** Folder names below the library root from a Graph parentReference.path ("/drives/x/root:/A/B" → ["A", "B"]). Pure. */
export function folderPathSegments(parentPath: string | undefined): string[] {
  const below = (parentPath ?? "").split("root:")[1] ?? "";
  return below
    .split("/")
    .filter(Boolean)
    .map((s) => {
      try {
        return decodeURIComponent(s);
      } catch {
        return s;
      }
    });
}
