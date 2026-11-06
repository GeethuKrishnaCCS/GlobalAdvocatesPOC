import * as React from "react";
import { FILING_FOLDER_PATH } from "../config";
import {
  findFolder,
  findMatterForDrive,
  getFolderPath,
  getMatterDriveId,
  type IMatter,
} from "../services/matterService";
import { classifyAttachment, emlFileName, type IFilingAttachment } from "../services/mailFiling";
import {
  canGetEmail,
  getAttachmentFile,
  getEmailFile,
  getFilingTarget,
  getOpenEmail,
  readItem,
} from "../services/outlookItem";
import { uploadFile, type ISavedFile } from "../services/saveToMatterService";
import { friendlyAuthError } from "../auth/authErrors";
import { useMatterFolders, type IMatterFolders, type IMatterLocation } from "./useMatterFolders";

export type FileStep = "matter" | "folder" | "items" | "done";

/** Items step row: the email itself or one of its attachments. */
export interface IFilingRow extends Omit<IFilingAttachment, "kind"> {
  kind: "email" | IFilingAttachment["kind"];
}

export const EMAIL_ROW_ID = "email";

export interface IFileToMatter {
  state: {
    step: FileStep;
    browser: IMatterFolders;
    rows: IFilingRow[];
    ticked: ReadonlySet<string>;
    isPrefilling: boolean;
    isFiling: boolean;
    error: string | undefined;
    filed: ISavedFile[];
  };
  actions: {
    selectMatter: (matter: IMatter) => Promise<void>;
    confirmFolder: () => void;
    back: () => void;
    toggle: (rowId: string, on: boolean) => void;
    submit: () => Promise<void>;
  };
  computed: { canSubmit: boolean; tickedCount: number };
}

function buildRows(): IFilingRow[] {
  const email = getOpenEmail();
  const emailRow: IFilingRow = {
    id: EMAIL_ROW_ID,
    fileName: emlFileName(email.subject, email.received),
    size: 0,
    kind: "email",
    fileable: canGetEmail(),
    ticked: canGetEmail(),
    reason: canGetEmail() ? undefined : "Update Outlook to file the email itself",
  };
  return [emailRow, ...email.attachments.map(classifyAttachment)];
}

/** Resolves where the email asks to be filed (Word Send marker or matter category); undefined when it cannot. */
async function resolvePrefill(
  matters: IMatter[]
): Promise<{ location: IMatterLocation; exact: boolean } | undefined> {
  const target = await getFilingTarget();
  if (!target) return undefined;
  if ("driveId" in target) {
    const matter = await findMatterForDrive(target.driveId, matters);
    if (!matter) return undefined;
    const folderPath = await getFolderPath(target.driveId, target.folderId);
    return { location: { matter, driveId: target.driveId, folderPath }, exact: true };
  }
  const matter = matters.find((m) => m.ref.toLowerCase() === target.matterRef.toLowerCase());
  if (!matter) return undefined;
  const driveId = await getMatterDriveId(matter.siteId);
  const folderId = await findFolder(driveId, FILING_FOLDER_PATH);
  const folderPath = folderId ? await getFolderPath(driveId, folderId) : [];
  return { location: { matter, driveId, folderPath }, exact: false };
}

/** File to Matter (Outlook read): Matter → Folder → Items → the email (.eml) and ticked attachments uploaded. */
export function useFileToMatter(): IFileToMatter {
  const browser = useMatterFolders();
  const {
    loadMatters,
    selectMatter: openMatter,
    openLocation,
    clearMatter,
    setError: setBrowserError,
  } = browser.actions;
  const [step, setStep] = React.useState<FileStep>("matter");
  const rows = React.useMemo(buildRows, []);
  const [ticked, setTicked] = React.useState<ReadonlySet<string>>(
    () => new Set(rows.filter((r) => r.fileable && r.ticked).map((r) => r.id))
  );
  const [isPrefilling, setIsPrefilling] = React.useState(true);
  const [isFiling, setIsFiling] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [filed, setFiled] = React.useState<ISavedFile[]>([]);
  /** Uploads done per target folder + row, so Submit after a partial failure only files what is missing. */
  const done = React.useRef(new Map<string, ISavedFile>());

  React.useEffect(() => {
    let active = true;
    void (async () => {
      const matters = await loadMatters();
      try {
        // No matters: keep loadMatters' error so the Matter step offers Retry.
        const prefill = matters && (await resolvePrefill(matters));
        if (active && prefill) {
          if (await openLocation(prefill.location)) setStep(prefill.exact ? "items" : "folder");
          // A marker pointing at a folder the user cannot reach: start at the Matter step as if there were none.
          else setBrowserError(undefined);
        }
      } catch {
        // Marker unreadable (Graph) or folder moved: same, start at the Matter step.
      } finally {
        if (active) setIsPrefilling(false);
      }
    })();
    return () => {
      active = false;
    };
  }, [loadMatters, openLocation, setBrowserError]);

  const selectMatter = React.useCallback(
    async (matter: IMatter) => {
      if (await openMatter(matter)) setStep("folder");
    },
    [openMatter]
  );

  const confirmFolder = React.useCallback(() => setStep("items"), []);

  const back = React.useCallback(() => {
    setError(undefined);
    if (step === "items") setStep("folder");
    else {
      clearMatter();
      setStep("matter");
    }
  }, [clearMatter, step]);

  const toggle = React.useCallback((rowId: string, on: boolean) => {
    setTicked((current) => {
      const next = new Set(current);
      if (on) next.add(rowId);
      else next.delete(rowId);
      return next;
    });
  }, []);

  const location = browser.state.location;
  const folderId = browser.computed.folderId;

  const submit = React.useCallback(async () => {
    if (!location) return;
    const selected = rows.filter((r) => ticked.has(r.id));
    const item = readItem();
    setIsFiling(true);
    setError(undefined);
    const results: ISavedFile[] = [];
    try {
      for (const row of selected) {
        const key = `${location.driveId}|${folderId}|${row.id}`;
        let saved = done.current.get(key);
        if (!saved) {
          const body =
            row.kind === "email"
              ? await getEmailFile(item)
              : await getAttachmentFile(item, { ...row, kind: row.kind });
          saved = await uploadFile(location.driveId, folderId, row.fileName, body);
          done.current.set(key, saved);
        }
        results.push(saved);
      }
      done.current.clear();
      setFiled(results);
      setStep("done");
    } catch (e) {
      const message = friendlyAuthError(e);
      setError(
        results.length
          ? `${results.length} of ${selected.length} filed; the rest failed: ${message} Select Submit to finish.`
          : message
      );
    } finally {
      setIsFiling(false);
    }
  }, [folderId, location, rows, ticked]);

  const tickedCount = ticked.size;
  return {
    state: { step, browser, rows, ticked, isPrefilling, isFiling, error, filed },
    actions: { selectMatter, confirmFolder, back, toggle, submit },
    computed: { canSubmit: !!location && tickedCount > 0 && !isFiling, tickedCount },
  };
}
