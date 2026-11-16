import type { IMatterFolderStatus } from "./matterFolderService";

/** The "Mailbox folders" line under My matters. */
export interface IFolderStatusLine {
  text: string;
  isError: boolean;
}

/** Only the matters the user has pinned (the backend may still list one just unpinned). Pure. */
export function statusForPins(
  status: IMatterFolderStatus[],
  pinned: ReadonlySet<string>
): IMatterFolderStatus[] {
  return status.filter((s) => pinned.has(s.matterRef));
}

/** "10:42" (local, 24 h). Pure. */
export function clockTime(date: Date): string {
  return `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
}

/**
 * The status line: nothing without pins; the registration/status-poll error, else the first matter's `lastError`, as
 * an error; otherwise a summary of the pinned matters. `checkedAt` is when the status was last read. Pure.
 */
export function folderStatusLine(
  status: IMatterFolderStatus[] | undefined,
  pinned: ReadonlySet<string>,
  checkedAt: Date | undefined,
  error: string | undefined
): IFolderStatusLine | undefined {
  if (pinned.size === 0) return undefined;
  if (error) return { text: `Mailbox folders: ${error}`, isError: true };
  const mine = statusForPins(status ?? [], pinned);
  const failed = mine.find((s) => s.lastError);
  if (failed)
    return { text: `Mailbox folders: ${failed.matterRef}: ${failed.lastError}`, isError: true };
  if (!status) return { text: "Mailbox folders: checking…", isError: false };
  const filed = mine.reduce((sum, s) => sum + s.filedCount, 0);
  const checked = checkedAt ? `, last checked ${clockTime(checkedAt)}` : "";
  return {
    text: `Mailbox folders: ${mine.length} ${mine.length === 1 ? "matter" : "matters"}, ${filed} ${filed === 1 ? "email" : "emails"} filed${checked}`,
    isError: false,
  };
}
