import * as React from "react";
import { FILING_FOLDER_PATH } from "../config";
import { findFolder, getMatterDriveId, type IMatter } from "../services/matterService";
import {
  dropKeyAt,
  dropKeyOf,
  filesFromDataTransfer,
  filesFromOfficeDrop,
  onOfficeDrag,
  type IDroppedFile,
  type OfficeDragData,
} from "../services/mailDrop";
import { uploadFile } from "../services/saveToMatterService";
import { friendlyAuthError } from "../auth/authErrors";
import { ROOT_FOLDER } from "./useMatterFolders";

/* global document, DragEvent, window */

/** Where a drop goes: a matter (its filing folder) or one library folder. */
export type DropTarget =
  | { kind: "matter"; matter: IMatter }
  | { kind: "folder"; driveId: string; folderId: string; label: string };

/** The library folder a drop target resolved to. */
export interface IDropDestination {
  driveId: string;
  folderId: string;
  label: string;
}

export type DropStatus = "waiting" | "filing" | "filed" | "failed";

/** One dropped file and how its filing went. */
export interface IDropItem {
  id: number;
  name: string;
  /** Where it goes, as shown to the user ("ADVTestMatter01 › Correspondence › Matter"). */
  target: string;
  status: DropStatus;
  error?: string;
  webUrl?: string;
}

export interface IDropToFile {
  state: {
    items: IDropItem[];
    /** Drop target under the pointer while dragging. */
    hoverKey: string | undefined;
    /** Why a drop was ignored. */
    notice: string | undefined;
  };
  actions: {
    retry: (id: number) => Promise<void>;
    clearFinished: () => void;
  };
  computed: { isFiling: boolean; filedCount: number; failedCount: number };
}

interface IJob {
  file: IDroppedFile;
  target: DropTarget;
}

const NOT_A_TARGET = "Drop emails onto a matter or a folder.";
const NOTHING_DROPPED =
  "Nothing to file: drag emails or attachments (not links or pictures inside an email).";
/** Outlook reports no "drag left the pane"; the highlight goes when no dragover followed for this long. */
const HOVER_TIMEOUT_MS = 1500;

/** Matter → its filing folder (library root when the matter has none). */
async function resolveDestination(target: DropTarget): Promise<IDropDestination> {
  if (target.kind === "folder") return target;
  const driveId = await getMatterDriveId(target.matter.siteId);
  const filingId = await findFolder(driveId, FILING_FOLDER_PATH);
  return {
    driveId,
    folderId: filingId ?? ROOT_FOLDER,
    label: filingId
      ? [target.matter.ref, ...FILING_FOLDER_PATH.split("/")].join(" › ")
      : target.matter.ref,
  };
}

/**
 * Files emails dragged from Outlook onto elements carrying a drop key (see mailDrop.ts): HTML drag-and-drop in classic
 * Outlook and on Mac, Outlook's own drag event on the web and in new Outlook. Each dropped file is uploaded on its own,
 * so one failure never loses the others; failed ones can be retried.
 * @param resolveTarget drop key → target, or undefined when the key is no longer on screen.
 * @param onFiled called once per drop (or retry) that filed something, with the folder it went to.
 */
export function useDropToFile(
  resolveTarget: (key: string) => DropTarget | undefined,
  onFiled?: (destination: IDropDestination) => void
): IDropToFile {
  const [items, setItems] = React.useState<IDropItem[]>([]);
  const [hoverKey, setHoverKey] = React.useState<string | undefined>(undefined);
  const [notice, setNotice] = React.useState<string | undefined>(undefined);
  const jobs = React.useRef(new Map<number, IJob>());
  const nextId = React.useRef(1);
  // The document / Outlook listeners are registered once; they read the newest callbacks through these.
  const resolveRef = React.useRef(resolveTarget);
  resolveRef.current = resolveTarget;
  const onFiledRef = React.useRef(onFiled);
  onFiledRef.current = onFiled;

  const patch = React.useCallback((id: number, changes: Partial<IDropItem>) => {
    setItems((current) => current.map((i) => (i.id === id ? { ...i, ...changes } : i)));
  }, []);

  /** Uploads the queued jobs one after the other (a burst of parallel uploads gets throttled by SharePoint). */
  const run = React.useCallback(
    async (ids: number[]) => {
      let destination: IDropDestination | undefined;
      let filed = 0;
      for (const id of ids) {
        const job = jobs.current.get(id);
        if (!job) continue;
        patch(id, { status: "filing", error: undefined });
        try {
          destination = destination ?? (await resolveDestination(job.target));
          patch(id, { target: destination.label });
          const saved = await uploadFile(
            destination.driveId,
            destination.folderId,
            job.file.name,
            job.file.body,
            "rename"
          );
          jobs.current.delete(id);
          filed++;
          patch(id, { status: "filed", name: saved.name, webUrl: saved.webUrl });
        } catch (e) {
          patch(id, { status: "failed", error: friendlyAuthError(e) });
        }
      }
      if (filed && destination) onFiledRef.current?.(destination);
    },
    [patch]
  );

  const drop = React.useCallback(
    (key: string | undefined, files: IDroppedFile[]) => {
      setHoverKey(undefined);
      const target = key ? resolveRef.current(key) : undefined;
      if (!target) return setNotice(NOT_A_TARGET);
      if (!files.length) return setNotice(NOTHING_DROPPED);
      setNotice(undefined);
      const label = target.kind === "matter" ? target.matter.ref : target.label;
      const added: IDropItem[] = files.map((file) => {
        const id = nextId.current++;
        jobs.current.set(id, { file, target });
        return { id, name: file.name, target: label, status: "waiting" };
      });
      setItems((current) => [...added, ...current]);
      void run(added.map((i) => i.id));
    },
    [run]
  );

  React.useEffect(() => {
    let hoverTimer: number | undefined;
    const onDragOver = (e: DragEvent): void => {
      // Without this the browser refuses the drop (and would open a dropped file instead of the pane).
      e.preventDefault();
      const key = dropKeyOf(e.target);
      if (e.dataTransfer) e.dataTransfer.dropEffect = key ? "copy" : "none";
      setHoverKey(key);
    };
    const onDragLeave = (e: DragEvent): void => {
      if (!e.relatedTarget) setHoverKey(undefined);
    };
    const onDrop = (e: DragEvent): void => {
      e.preventDefault();
      drop(dropKeyOf(e.target), filesFromDataTransfer(e.dataTransfer));
    };
    const onOffice = (data: OfficeDragData): void => {
      const key = dropKeyAt(data.taskPaneX, data.taskPaneY);
      window.clearTimeout(hoverTimer);
      if (data.type === "drop") return drop(key, filesFromOfficeDrop(data.dataTransfer.files));
      setHoverKey(key);
      hoverTimer = window.setTimeout(() => setHoverKey(undefined), HOVER_TIMEOUT_MS);
    };
    document.addEventListener("dragover", onDragOver);
    document.addEventListener("dragleave", onDragLeave);
    document.addEventListener("drop", onDrop);
    const offOffice = onOfficeDrag(onOffice);
    return () => {
      window.clearTimeout(hoverTimer);
      document.removeEventListener("dragover", onDragOver);
      document.removeEventListener("dragleave", onDragLeave);
      document.removeEventListener("drop", onDrop);
      offOffice();
    };
  }, [drop]);

  const retry = React.useCallback((id: number) => run([id]), [run]);

  const clearFinished = React.useCallback(() => {
    setItems((current) => current.filter((i) => i.status !== "filed"));
  }, []);

  const computed = React.useMemo(
    () => ({
      isFiling: items.some((i) => i.status === "filing" || i.status === "waiting"),
      filedCount: items.filter((i) => i.status === "filed").length,
      failedCount: items.filter((i) => i.status === "failed").length,
    }),
    [items]
  );

  return { state: { items, hoverKey, notice }, actions: { retry, clearFinished }, computed };
}
