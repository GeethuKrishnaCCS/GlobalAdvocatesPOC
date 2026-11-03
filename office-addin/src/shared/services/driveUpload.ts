import { FileUpload, LargeFileUploadTask } from "@microsoft/microsoft-graph-client";
import type { DriveItem } from "@microsoft/microsoft-graph-types";
import { graphClient } from "./graphClient";

/** Graph's limit for a single PUT /content; larger files need an upload session. */
const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024;
/** Upload session slice: Graph requires multiples of 320 KiB. */
const UPLOAD_SLICE = 10 * 320 * 1024;
/** Same name in the folder → new version of that file (library versioning on), never a renamed copy. */
const CONFLICT_BEHAVIOR = "replace";

/** DriveItem fields the pane uses after an upload. */
const ITEM_SELECT = "id,name,webUrl,size,parentReference";

/**
 * Uploads `content` as `fileName` into a drive folder (`folderId` "root" for the library root), replacing a file of
 * the same name as a new version. `fileName` must already be clean (see fileName.ts).
 */
export async function uploadToFolder(
  driveId: string,
  folderId: string,
  fileName: string,
  content: ArrayBuffer
): Promise<DriveItem> {
  const itemPath = `/drives/${driveId}/items/${folderId}:/${encodeURIComponent(fileName)}:`;
  if (content.byteLength <= SIMPLE_UPLOAD_LIMIT) {
    return (
      (await graphClient
        .api(`${itemPath}/content`)
        // graph-client labels put() bodies as JSON unless told otherwise.
        .header("Content-Type", "application/octet-stream")
        .query({ "@microsoft.graph.conflictBehavior": CONFLICT_BEHAVIOR, $select: ITEM_SELECT })
        .put(content)) as DriveItem
    );
  }
  const session = await LargeFileUploadTask.createUploadSession(
    graphClient,
    `${itemPath}/createUploadSession`,
    {
      item: { "@microsoft.graph.conflictBehavior": CONFLICT_BEHAVIOR },
    }
  );
  // The SDK sends slices to the pre-authenticated uploadUrl without the Graph bearer token.
  const task = new LargeFileUploadTask(
    graphClient,
    new FileUpload(content, fileName, content.byteLength),
    session,
    {
      rangeSize: UPLOAD_SLICE,
    }
  );
  const result = await task.upload();
  return result.responseBody as DriveItem;
}
