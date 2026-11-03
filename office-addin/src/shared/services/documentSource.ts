import { ResponseType } from "@microsoft/microsoft-graph-client";
import type { DriveItem } from "@microsoft/microsoft-graph-types";
import { graphClient } from "./graphClient";
import { splitFileName } from "./fileName";

/* global Office, btoa, console, TextEncoder, Blob */

/** Office's maximum getFileAsync slice. */
const SLICE_SIZE = 4 * 1024 * 1024;

/**
 * Extensions each host saves in (first = default). getFileAsync(Compressed) returns Open XML, so a legacy .doc/.xls/.ppt
 * becomes the default type; macro-enabled files keep their extension. The user renames the file, not its type.
 */
const HOST_EXTENSIONS: Partial<Record<Office.HostType, string[]>> = {
  [Office.HostType.Word]: [".docx", ".docm"],
  [Office.HostType.Excel]: [".xlsx", ".xlsm"],
  [Office.HostType.PowerPoint]: [".pptx", ".pptm"],
};

export type DocumentFormat = "native" | "pdf";

/** The document open in the host. */
export interface ICurrentDocument {
  /** Where it lives: a SharePoint/OneDrive URL on the web, a local path on desktop, "" when never saved. */
  url: string;
  baseName: string;
  extension: string;
  /** Parent folder name for display ("" when unknown). */
  folderName: string;
}

/** Decodes a URL segment; desktop paths are not URL-encoded, so "50% off" must survive as is. */
function decodeSegment(segment: string): string {
  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
}

/** Name and type of the open document. */
export function getCurrentDocument(host: Office.HostType): ICurrentDocument {
  const url = Office.context.document.url ?? "";
  const segments = url.split(/[?#]/)[0].split(/[\\/]/).filter(Boolean).map(decodeSegment);
  const { baseName, extension } = splitFileName(segments.pop() ?? "");
  const allowed = HOST_EXTENSIONS[host];
  return {
    url,
    baseName: baseName || "Document",
    extension: allowed
      ? allowed.includes(extension.toLowerCase())
        ? extension.toLowerCase()
        : allowed[0]
      : extension,
    folderName: segments.length > 1 ? segments[segments.length - 1] : "",
  };
}

/** Whole file through Office.js, slice by slice; the file handle is always closed. */
function readWithOffice(fileType: Office.FileType): Promise<Blob> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(fileType, { sliceSize: SLICE_SIZE }, (opened) => {
      if (opened.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(opened.error.message));
        return;
      }
      const file = opened.value;
      // Typed copies of each slice; the Blob is built from them without one more full-size copy.
      const slices: Uint8Array<ArrayBuffer>[] = [];
      const finish = (error?: Error): void =>
        file.closeAsync(() => (error ? reject(error) : resolve(new Blob(slices))));
      const readSlice = (index: number): void =>
        file.getSliceAsync(index, (slice) => {
          if (slice.status === Office.AsyncResultStatus.Failed) {
            finish(new Error(slice.error.message));
            return;
          }
          slices.push(new Uint8Array(slice.value.data as number[]));
          if (index + 1 < file.sliceCount) readSlice(index + 1);
          else finish();
        });
      if (file.sliceCount === 0) finish();
      else readSlice(0);
    });
  });
}

/** Graph "shares" id for a SharePoint/OneDrive URL (base64url, "u!" prefix). */
function shareId(url: string): string {
  const base64 = btoa(String.fromCharCode(...new TextEncoder().encode(url)));
  return `u!${base64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-")}`;
}

/** Saved copy from SharePoint/OneDrive (Office on the web autosaves, so it is current); PDF converted by Graph. */
async function readWithGraph(url: string, format: DocumentFormat): Promise<Blob> {
  if (!/^https:\/\//i.test(url))
    throw new Error("Save the document to OneDrive or SharePoint first.");
  const item = (await graphClient
    .api(`/shares/${shareId(url)}/driveItem`)
    .select("id,parentReference")
    .get()) as DriveItem;
  const request = graphClient
    .api(`/drives/${item.parentReference!.driveId}/items/${item.id}/content`)
    .responseType(ResponseType.ARRAYBUFFER);
  if (format === "pdf") request.query({ format: "pdf" });
  return new Blob([(await request.get()) as ArrayBuffer]);
}

/**
 * Bytes of the open document, as is or as PDF, through Office.js (unsaved changes included). Only Office on the web
 * falls back to Graph when the host cannot export, because there the saved copy is current; on desktop the saved copy
 * may lack edits, so the Office error is shown instead.
 */
export async function readDocument(doc: ICurrentDocument, format: DocumentFormat): Promise<Blob> {
  try {
    return await readWithOffice(
      format === "pdf" ? Office.FileType.Pdf : Office.FileType.Compressed
    );
  } catch (officeError) {
    if (Office.context.platform !== Office.PlatformType.OfficeOnline) throw officeError;
    console.log(
      `[documentSource] getFileAsync(${format}) failed, using Graph: ${(officeError as Error).message}`
    );
    return readWithGraph(doc.url, format);
  }
}
