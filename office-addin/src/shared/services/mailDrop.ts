import { MAX_FILE_NAME_LENGTH, cleanBaseName } from "./fileName";

/* global Office, Blob, DataTransfer, EventTarget, Element, document */

/** A message (.eml / .msg) or attachment dragged from Outlook into the pane. */
export interface IDroppedFile {
  name: string;
  body: Blob;
}

/** Attribute on every element emails can be dropped onto; its value identifies the target. */
export const DROP_KEY_ATTRIBUTE = "data-drop-key";
const NO_SUBJECT = "(no subject)";

export const matterDropKey = (siteId: string): string => `matter:${siteId}`;
export const folderDropKey = (folderId: string): string => `folder:${folderId}`;

/** Outlook's file name for a dragged item, cleaned for SharePoint and within its name length. Pure. */
export function droppedFileName(name: string): string {
  const trimmed = name.trim();
  // Not splitFileName: an email without subject arrives as ".eml", which must keep its extension.
  const dot = trimmed.lastIndexOf(".");
  const ext = dot < 0 ? "" : trimmed.slice(dot).replace(/[^.\w-]/g, "");
  const base = cleanBaseName(dot < 0 ? trimmed : trimmed.slice(0, dot)) || NO_SUBJECT;
  return base.slice(0, MAX_FILE_NAME_LENGTH - ext.length).trimEnd() + ext;
}

/** True for a filed email (.eml from Outlook on the web / new Outlook / Mac, .msg from classic Outlook). Pure. */
export function isEmailFile(name: string): boolean {
  return /\.(eml|msg)$/i.test(name);
}

/** Files of an HTML drop (classic Outlook on Windows, Mac). */
export function filesFromDataTransfer(dataTransfer: DataTransfer | null): IDroppedFile[] {
  return Array.from(dataTransfer?.files ?? []).map((f) => ({
    name: droppedFileName(f.name),
    body: f,
  }));
}

/** Files of an Office.js drop (Outlook on the web, new Outlook on Windows). */
export function filesFromOfficeDrop(files: { name: string; fileContent: Blob }[]): IDroppedFile[] {
  return files.map((f) => ({ name: droppedFileName(f.name), body: f.fileContent }));
}

/** Drop target key of the element (or its nearest ancestor) an HTML drag event hit. */
export function dropKeyOf(target: EventTarget | null): string | undefined {
  const element = target instanceof Element ? target.closest(`[${DROP_KEY_ATTRIBUTE}]`) : null;
  return element?.getAttribute(DROP_KEY_ATTRIBUTE) ?? undefined;
}

/** Drop target key under a point of the pane (Office.js reports the pointer position, not an element). */
export function dropKeyAt(x: number, y: number): string | undefined {
  return dropKeyOf(document.elementFromPoint(x, y));
}

export type OfficeDragData = Office.DragAndDropEventArgs["dragAndDropEventData"];

/**
 * Subscribes to Outlook's own drag-and-drop event (Outlook on the web, new Outlook on Windows: items dragged there
 * never reach the pane as HTML drag events). No-op where Outlook lacks it (classic, Mac). Returns the unsubscribe.
 */
export function onOfficeDrag(handler: (data: OfficeDragData) => void): () => void {
  const mailbox = Office.context.mailbox;
  const type = Office.EventType.DragAndDropEvent;
  if (!mailbox?.addHandlerAsync || !type) return () => undefined;
  mailbox.addHandlerAsync(type, (args: Office.DragAndDropEventArgs) =>
    handler(args.dragAndDropEventData)
  );
  return () => mailbox.removeHandlerAsync(type);
}
