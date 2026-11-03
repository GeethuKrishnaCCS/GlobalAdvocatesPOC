import type { DriveItem, FieldValueSet } from "@microsoft/microsoft-graph-types";
import { graphClient } from "./graphClient";
import { uploadToFolder } from "./driveUpload";

/* global Blob */

/** Document properties set on the saved file (matter library columns, by internal name). */
export interface IDocumentProperties {
  docType: string;
  direction: string;
  description: string;
}

/** A file uploaded to the matter library; `version` is set once its properties are applied. */
export interface ISavedFile {
  itemId: string;
  name: string;
  webUrl: string;
  version: string | undefined;
}

const HTTP_LOCKED = 423;

function isLocked(e: unknown): boolean {
  const error = e as { statusCode?: number; code?: string; message?: string };
  // Simple uploads report the status; upload sessions only carry the Graph error code / text.
  return (
    error.statusCode === HTTP_LOCKED ||
    error.code === "resourceLocked" ||
    /locked/i.test(error.message ?? "")
  );
}

/** Library columns written; Description is SharePoint's built-in column. An empty choice clears the column. */
function toFields(properties: IDocumentProperties): Record<string, string | null> {
  return {
    DocType: properties.docType || null,
    AdvDirection: properties.direction || null,
    _ExtendedDescription: properties.description || null,
  };
}

/** Uploads one file into a library folder ("root" for the library root); same name → new version of that file. */
export async function uploadFile(
  driveId: string,
  folderId: string,
  fileName: string,
  body: Blob
): Promise<ISavedFile> {
  let item: DriveItem;
  try {
    item = await uploadToFolder(driveId, folderId, fileName, body);
  } catch (e) {
    if (isLocked(e)) {
      throw new Error(
        `"${fileName}" is open for editing in that folder, so it cannot be replaced. Choose another name or folder.`
      );
    }
    throw e;
  }
  return {
    itemId: item.id!,
    name: item.name ?? fileName,
    webUrl: item.webUrl ?? "",
    version: undefined,
  };
}

/**
 * Sets the document properties and returns the resulting version label.
 * shortcut: this PATCH adds a second version per save (Graph cannot set fields in the upload); move to SharePoint
 * ValidateUpdateListItem(bNewDocumentUpdate) if version noise matters to the firm.
 */
export async function applyProperties(
  driveId: string,
  file: ISavedFile,
  properties: IDocumentProperties
): Promise<ISavedFile> {
  const fieldsPath = `/drives/${driveId}/items/${file.itemId}/listItem/fields`;
  await graphClient.api(fieldsPath).patch(toFields(properties));
  const fields = (await graphClient
    .api(fieldsPath)
    .select("_UIVersionString")
    .get()) as FieldValueSet & {
    _UIVersionString?: string;
  };
  return { ...file, version: fields._UIVersionString ?? "" };
}
