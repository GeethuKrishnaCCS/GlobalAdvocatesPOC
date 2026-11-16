import { getApiToken } from "../auth/naa";
import { API_BASE_URL, MATTER_LIBRARY_NAME } from "../config";
import type { IMatter } from "./matterService";

/* global fetch, Headers */

/** What the backend reports for one matter it keeps mail folders for. */
export interface IMatterFolderStatus {
  matterRef: string;
  filedCount: number;
  /** ISO time of the last sweep, or null before the first one. */
  lastSweepUtc: string | null;
  lastError: string | null;
}

const HTTP_NO_CONTENT = 204;

/** Calls the backend with the user's token. A failed call throws an Error carrying the backend's `{ error }` text. */
async function apiRequest<T>(
  method: string,
  path: string,
  allowPopup: boolean,
  body?: unknown
): Promise<T | undefined> {
  const headers = new Headers({ Authorization: `Bearer ${await getApiToken(allowPopup)}` });
  if (body !== undefined) headers.set("Content-Type", "application/json");
  const response = await fetch(`${API_BASE_URL}${path}`, {
    method,
    headers,
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  if (!response.ok) {
    const failure = (await response.json().catch(() => undefined)) as
      { error?: string } | undefined;
    throw new Error(
      failure?.error ?? `The mailbox-folder service failed (HTTP ${response.status}).`
    );
  }
  return response.status === HTTP_NO_CONTENT ? undefined : ((await response.json()) as T);
}

const folderPath = (ref: string): string => `/matter-folders/${encodeURIComponent(ref)}`;

/**
 * Has the backend create (or refresh) the matter's mail folders in the user's mailbox. Idempotent.
 * `allowPopup` only from a click (sign-in dialog fallback).
 */
export async function registerMatterFolders(matter: IMatter, allowPopup = false): Promise<void> {
  await apiRequest("PUT", folderPath(matter.ref), allowPopup, {
    webUrl: matter.webUrl,
    documentsList: MATTER_LIBRARY_NAME,
  });
}

/** Stops the backend filing for this matter; the mail folders stay in the mailbox. */
export async function unregisterMatterFolders(ref: string, allowPopup = false): Promise<void> {
  await apiRequest("DELETE", folderPath(ref), allowPopup);
}

/** Filing status of every matter registered for the user. */
export async function getMatterFolderStatus(): Promise<IMatterFolderStatus[]> {
  const result = await apiRequest<{ matters: IMatterFolderStatus[] }>(
    "GET",
    "/matter-folders",
    false
  );
  return result?.matters ?? [];
}
