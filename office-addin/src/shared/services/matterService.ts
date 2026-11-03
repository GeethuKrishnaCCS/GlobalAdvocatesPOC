import type { Drive, DriveItem, List, Site } from "@microsoft/microsoft-graph-types";
import { MATTER_LIBRARY_NAME, MATTER_SITE_SEARCH } from "../config";
import { graphClient } from "./graphClient";

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
}

/** Library folders SharePoint creates for itself; never a save target. */
const SYSTEM_FOLDERS = new Set(["Forms"]);
const PAGE_SIZE = 999;

/** Matter sites the signed-in user can access, sorted by reference. */
export async function listMatters(): Promise<IMatter[]> {
  const result = (await graphClient
    .api("/sites")
    .query({ search: MATTER_SITE_SEARCH })
    .select("id,name,displayName,webUrl")
    .get()) as IGraphCollection<Site>;
  // Site search is fuzzy; keep only sites whose URL alias carries the matter prefix.
  const prefix = MATTER_SITE_SEARCH.toLowerCase();
  return result.value
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
  const lists = (await graphClient
    .api(`/sites/${siteId}/lists`)
    .select("id,name")
    .top(PAGE_SIZE)
    .get()) as IGraphCollection<List>;
  const library = lists.value.find(
    (l) => l.name?.toLowerCase() === MATTER_LIBRARY_NAME.toLowerCase()
  );
  if (!library?.id) throw new Error(`This matter has no "${MATTER_LIBRARY_NAME}" library.`);
  const drive = (await graphClient
    .api(`/sites/${siteId}/lists/${library.id}/drive`)
    .select("id")
    .get()) as Drive;
  return drive.id!;
}

/** Sub-folders of a library folder ("root" for the library root), sorted by name. */
export async function listFolders(driveId: string, folderId: string): Promise<IMatterFolder[]> {
  const children = (await graphClient
    .api(`/drives/${driveId}/items/${folderId}/children`)
    .select("id,name,webUrl,folder")
    .top(PAGE_SIZE)
    .get()) as IGraphCollection<DriveItem>;
  return children.value
    .filter((c) => c.folder && !SYSTEM_FOLDERS.has(c.name ?? ""))
    .map((c) => ({ id: c.id!, name: c.name!, webUrl: c.webUrl ?? "" }))
    .sort((a, b) => a.name.localeCompare(b.name));
}
