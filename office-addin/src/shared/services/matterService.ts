import { ResponseType, type GraphRequest } from "@microsoft/microsoft-graph-client";
import type { Drive, DriveItem, List, Site } from "@microsoft/microsoft-graph-types";
import { MATTER_LIBRARY_NAME, MATTER_SITE_SEARCH } from "../config";
import { graphClient } from "./graphClient";
import { folderPathSegments } from "./mailFiling";

/* global console */

/** A matter = one SharePoint site. */
export interface IMatter {
  siteId: string;
  /** Matter reference shown to users (site URL alias until the AdvMatterRef property is provisioned). */
  ref: string;
  title: string;
  webUrl: string;
}

/** A folder in a matter library. */
export interface IMatterFolder {
  id: string;
  name: string;
  webUrl: string;
}

interface IGraphCollection<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

/** Library folders SharePoint creates for itself; never a save target. */
const SYSTEM_FOLDERS = new Set(["Forms"]);
const PAGE_SIZE = 999;
/** Pages read at most per listing (999 entries each); stops a runaway loop on huge tenants or folders. */
const MAX_PAGES = 20;

/** All pages of a Graph collection, following @odata.nextLink up to MAX_PAGES. */
async function getAllPages<T>(request: GraphRequest, what: string): Promise<T[]> {
  const items: T[] = [];
  let page = (await request.get()) as IGraphCollection<T>;
  for (let pages = 1; ; pages++) {
    items.push(...page.value);
    const next = page["@odata.nextLink"];
    if (!next) return items;
    if (pages >= MAX_PAGES) {
      console.warn(
        `[matterService] ${what}: stopped after ${MAX_PAGES} pages; the list is incomplete.`
      );
      return items;
    }
    page = (await graphClient.api(next).get()) as IGraphCollection<T>;
  }
}

/**
 * Matter sites the signed-in user can access, sorted by reference. Site search matches whole words only, so a prefix
 * ("ADVTestMatter" vs "ADVTestMatter01") finds nothing; list all accessible sites (search=*) and filter by URL alias.
 * shortcut: client-side filter over all sites; switch to a matter registry list or hub site when the firm has
 * thousands of sites.
 */
export async function listMatters(): Promise<IMatter[]> {
  const prefix = MATTER_SITE_SEARCH.toLowerCase();
  const sites = await getAllPages<Site>(
    graphClient
      .api("/sites")
      .query({ search: "*" })
      .select("id,name,displayName,webUrl")
      .top(PAGE_SIZE),
    "sites"
  );
  return sites
    .filter((s) => (s.name ?? "").toLowerCase().startsWith(prefix))
    .map((s) => ({
      siteId: s.id!,
      ref: s.name!,
      title: s.displayName ?? s.name!,
      webUrl: s.webUrl ?? "",
    }))
    .sort((a, b) => a.ref.localeCompare(b.ref));
}

/**
 * Drive id of the matter library. Resolved by list id (`/lists/{id}/drive`), with the list matched on its URL name,
 * so localized library titles ("Dokumenter") do not break it.
 */
export async function getMatterDriveId(siteId: string): Promise<string> {
  const lists = await getAllPages<List>(
    graphClient.api(`/sites/${siteId}/lists`).select("id,name").top(PAGE_SIZE),
    "lists"
  );
  const library = lists.find((l) => l.name?.toLowerCase() === MATTER_LIBRARY_NAME.toLowerCase());
  if (!library?.id) throw new Error(`This matter has no "${MATTER_LIBRARY_NAME}" library.`);
  const drive = (await graphClient
    .api(`/sites/${siteId}/lists/${library.id}/drive`)
    .select("id")
    .get()) as Drive;
  return drive.id!;
}

/** Sub-folders of a library folder ("root" for the library root), sorted by name. */
export async function listFolders(driveId: string, folderId: string): Promise<IMatterFolder[]> {
  // Files count towards each page too, so a big folder needs several pages to show all its sub-folders.
  const children = await getAllPages<DriveItem>(
    graphClient
      .api(`/drives/${driveId}/items/${folderId}/children`)
      .select("id,name,webUrl,folder")
      .top(PAGE_SIZE),
    "folders"
  );
  return children
    .filter((c) => c.folder && !SYSTEM_FOLDERS.has(c.name ?? ""))
    .map((c) => ({ id: c.id!, name: c.name!, webUrl: c.webUrl ?? "" }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

const HTTP_NOT_FOUND = 404;

/** Item id of a folder by path from the library root ("Correspondence/Matter"), or undefined when it does not exist. */
export async function findFolder(driveId: string, path: string): Promise<string | undefined> {
  const encoded = path.split("/").map(encodeURIComponent).join("/");
  try {
    const item = (await graphClient
      .api(`/drives/${driveId}/root:/${encoded}`)
      .select("id,folder")
      .get()) as DriveItem;
    return item.folder ? item.id : undefined;
  } catch (e) {
    if ((e as { statusCode?: number }).statusCode === HTTP_NOT_FOUND) return undefined;
    throw e;
  }
}

/** A file in a matter library. `docId` is the SharePoint Document ID (when the site has the feature on). */
export interface IMatterFile {
  id: string;
  name: string;
  webUrl: string;
  size: number;
  docId: string | undefined;
}

type DriveItemWithDocId = DriveItem & { listItem?: { fields?: { _dlc_DocId?: string } } };

function toFile(item: DriveItemWithDocId): IMatterFile {
  return {
    id: item.id!,
    name: item.name!,
    webUrl: item.webUrl ?? "",
    size: item.size ?? 0,
    docId: item.listItem?.fields?._dlc_DocId || undefined,
  };
}

/** Sub-folders and files of a library folder ("root" for the library root), each sorted by name. */
export async function listFolderContents(
  driveId: string,
  folderId: string
): Promise<{ folders: IMatterFolder[]; files: IMatterFile[] }> {
  const children = await getAllPages<DriveItemWithDocId>(
    graphClient
      .api(`/drives/${driveId}/items/${folderId}/children`)
      .select("id,name,webUrl,size,folder,file")
      .expand("listItem($select=id;$expand=fields($select=_dlc_DocId))")
      .top(PAGE_SIZE),
    "folder contents"
  );
  const byName = (a: { name: string }, b: { name: string }): number => a.name.localeCompare(b.name);
  return {
    folders: children
      .filter((c) => c.folder && !SYSTEM_FOLDERS.has(c.name ?? ""))
      .map((c) => ({ id: c.id!, name: c.name!, webUrl: c.webUrl ?? "" }))
      .sort(byName),
    files: children
      .filter((c) => c.file)
      .map(toFile)
      .sort(byName),
  };
}

/** Files anywhere in the library matching `query` (name or content); Doc IDs are not returned by search. */
export async function searchFiles(driveId: string, query: string): Promise<IMatterFile[]> {
  const q = query.replace(/'/g, "''");
  const result = (await graphClient
    .api(`/drives/${driveId}/root/search(q='${encodeURIComponent(q)}')`)
    .select("id,name,webUrl,size,file")
    .top(PAGE_SIZE)
    .get()) as IGraphCollection<DriveItem>;
  return result.value.filter((i) => i.file).map(toFile);
}

/** File content. */
export async function downloadFile(driveId: string, itemId: string): Promise<ArrayBuffer> {
  return (await graphClient
    .api(`/drives/${driveId}/items/${itemId}/content`)
    .responseType(ResponseType.ARRAYBUFFER)
    .get()) as ArrayBuffer;
}

/** The matter (of `matters`) whose site holds this drive, or undefined. */
export async function findMatterForDrive(
  driveId: string,
  matters: IMatter[]
): Promise<IMatter | undefined> {
  const root = (await graphClient
    .api(`/drives/${driveId}/root`)
    .select("sharepointIds")
    .get()) as DriveItem;
  const siteGuid = root.sharepointIds?.siteId?.toLowerCase();
  // Graph site ids are "<host>,<site guid>,<web guid>".
  return siteGuid
    ? matters.find((m) => m.siteId.toLowerCase().split(",")[1] === siteGuid)
    : undefined;
}

/** Folders from the library root down to (and including) `folderId`, for the breadcrumb. */
export async function getFolderPath(driveId: string, folderId: string): Promise<IMatterFolder[]> {
  const item = (await graphClient
    .api(`/drives/${driveId}/items/${folderId}`)
    .select("id,name,webUrl,folder,parentReference")
    .get()) as DriveItem;
  if (!item.folder) throw new Error("The filing folder no longer exists.");
  const names = folderPathSegments(item.parentReference?.path);
  const path: IMatterFolder[] = [];
  for (let depth = 1; depth <= names.length; depth++) {
    const id = await findFolder(driveId, names.slice(0, depth).join("/"));
    if (!id) throw new Error("The filing folder was moved.");
    path.push({ id, name: names[depth - 1], webUrl: "" });
  }
  return [...path, { id: item.id!, name: item.name!, webUrl: item.webUrl ?? "" }];
}
