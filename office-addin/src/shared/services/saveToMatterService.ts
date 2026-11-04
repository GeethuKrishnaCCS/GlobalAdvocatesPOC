import type { DriveItem, FieldValueSet } from "@microsoft/microsoft-graph-types";
import { graphClient } from "./graphClient";
import { uploadToFolder } from "./driveUpload";

/* global Blob, fetch */

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

interface IFieldResult {
  FieldName: string;
  HasException: boolean;
  ErrorMessage: string | null;
}

/**
 * Sets the fields on the version the upload just created (bNewDocumentUpdate), so a first save stays 1.0 and every
 * update is exactly +1. SharePoint REST, because Graph has no way to update fields without adding a version.
 */
async function setFieldsWithoutNewVersion(
  driveId: string,
  itemId: string,
  properties: IDocumentProperties,
  sharePointToken: string
): Promise<void> {
  const item = (await graphClient
    .api(`/drives/${driveId}/items/${itemId}`)
    .select("sharepointIds")
    .get()) as DriveItem;
  const ids = item.sharepointIds!;
  const response = await fetch(
    `${ids.siteUrl}/_api/web/lists(guid'${ids.listId}')/items(${ids.listItemId})/ValidateUpdateListItem`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${sharePointToken}`,
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
      },
      body: JSON.stringify({
        formValues: Object.entries(toFields(properties)).map(([FieldName, value]) => ({
          FieldName,
          FieldValue: value ?? "",
        })),
        bNewDocumentUpdate: true,
      }),
    }
  );
  if (!response.ok)
    throw new Error(`Setting the document properties failed (HTTP ${response.status}).`);
  const failed = ((await response.json()) as { value: IFieldResult[] }).value.filter(
    (f) => f.HasException
  );
  if (failed.length) {
    throw new Error(
      `Setting ${failed.map((f) => `${f.FieldName}: ${f.ErrorMessage}`).join("; ")} failed.`
    );
  }
}

/**
 * Sets the document properties on the version the upload created and returns its label: one version per save
 * (1.0 on first save, +1 per update). There is deliberately no Graph PATCH fallback; it would add a second version.
 */
export async function applyProperties(
  driveId: string,
  file: ISavedFile,
  properties: IDocumentProperties,
  sharePointToken: string
): Promise<ISavedFile> {
  const fieldsPath = `/drives/${driveId}/items/${file.itemId}/listItem/fields`;
  await setFieldsWithoutNewVersion(driveId, file.itemId, properties, sharePointToken);
  const fields = (await graphClient
    .api(fieldsPath)
    .select("_UIVersionString")
    .get()) as FieldValueSet & {
    _UIVersionString?: string;
  };
  return { ...file, version: fields._UIVersionString ?? "" };
}
