import * as React from "react";
import {
  getMatterDriveId,
  listFolders,
  listMatters,
  type IMatter,
  type IMatterFolder,
} from "../services/matterService";
import {
  getCurrentDocument,
  readDocument,
  type ICurrentDocument,
} from "../services/documentSource";
import {
  applyProperties,
  uploadFile,
  type IDocumentProperties,
  type ISavedFile,
} from "../services/saveToMatterService";
import { fileNameProblem, toBaseName } from "../services/fileName";

/* global Office */

export type SaveStep = "matter" | "folder" | "details" | "done";

/** What a save has done so far, so a retry after a partial failure finishes it instead of uploading again. */
interface ISaveProgress {
  /** Drive, folder and name the progress belongs to; a different target starts over. */
  key: string;
  document?: ISavedFile;
  pdf?: ISavedFile;
}

export interface ISaveDetails extends IDocumentProperties {
  baseName: string;
  pdfCopy: boolean;
}

export interface ISaveResult {
  document: ISavedFile;
  pdf: ISavedFile | undefined;
}

export interface ISaveToMatter {
  state: {
    step: SaveStep;
    document: ICurrentDocument;
    matters: IMatter[] | undefined;
    matter: IMatter | undefined;
    /** Folders from the library root down to the open folder ([] = library root). */
    folderPath: IMatterFolder[];
    folders: IMatterFolder[] | undefined;
    details: ISaveDetails;
    isLoading: boolean;
    isSaving: boolean;
    error: string | undefined;
    result: ISaveResult | undefined;
  };
  actions: {
    loadMatters: () => Promise<void>;
    selectMatter: (matter: IMatter) => Promise<void>;
    openFolder: (folder: IMatterFolder) => Promise<void>;
    openPathLevel: (depth: number) => Promise<void>;
    confirmFolder: () => void;
    updateDetails: (patch: Partial<ISaveDetails>) => void;
    back: () => void;
    submit: () => Promise<void>;
    saveAnother: () => void;
  };
  computed: {
    fileName: string;
    fileNameError: string | undefined;
    locationLabel: string;
    canSubmit: boolean;
  };
}

const ROOT = "root";

/** Save to Matter wizard: Matter → Folder → Details → upload (+ optional PDF copy). */
export function useSaveToMatter(host: Office.HostType): ISaveToMatter {
  const document = React.useMemo(() => getCurrentDocument(host), [host]);
  const [step, setStep] = React.useState<SaveStep>("matter");
  const [matters, setMatters] = React.useState<IMatter[] | undefined>(undefined);
  const [matter, setMatter] = React.useState<IMatter | undefined>(undefined);
  const [driveId, setDriveId] = React.useState<string | undefined>(undefined);
  const [folderPath, setFolderPath] = React.useState<IMatterFolder[]>([]);
  const [folders, setFolders] = React.useState<IMatterFolder[] | undefined>(undefined);
  const [details, setDetails] = React.useState<ISaveDetails>({
    baseName: document.baseName,
    docType: "",
    direction: "",
    description: "",
    pdfCopy: false,
  });
  const [isLoading, setIsLoading] = React.useState(false);
  const [isSaving, setIsSaving] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [result, setResult] = React.useState<ISaveResult | undefined>(undefined);

  /** Runs a load step with one spinner/error slot; returns false when it failed. */
  const run = React.useCallback(async (work: () => Promise<void>): Promise<boolean> => {
    setIsLoading(true);
    setError(undefined);
    try {
      await work();
      return true;
    } catch (e) {
      setError((e as Error).message);
      return false;
    } finally {
      setIsLoading(false);
    }
  }, []);

  const loadMatters = React.useCallback(async () => {
    await run(async () => setMatters(await listMatters()));
  }, [run]);

  const selectMatter = React.useCallback(
    async (selected: IMatter) => {
      const ok = await run(async () => {
        const id = await getMatterDriveId(selected.siteId);
        setFolders(await listFolders(id, ROOT));
        setDriveId(id);
      });
      if (!ok) return;
      setMatter(selected);
      setFolderPath([]);
      setStep("folder");
    },
    [run]
  );

  const openFolder = React.useCallback(
    async (folder: IMatterFolder) => {
      if (!driveId) return;
      if (await run(async () => setFolders(await listFolders(driveId, folder.id)))) {
        setFolderPath((path) => [...path, folder]);
      }
    },
    [driveId, run]
  );

  const openPathLevel = React.useCallback(
    async (depth: number) => {
      if (!driveId) return;
      const target = depth === 0 ? ROOT : folderPath[depth - 1].id;
      if (await run(async () => setFolders(await listFolders(driveId, target)))) {
        setFolderPath((path) => path.slice(0, depth));
      }
    },
    [driveId, folderPath, run]
  );

  const confirmFolder = React.useCallback(() => {
    setError(undefined);
    setStep("details");
  }, []);

  const updateDetails = React.useCallback((patch: Partial<ISaveDetails>) => {
    setDetails((current) => ({ ...current, ...patch }));
  }, []);

  const back = React.useCallback(() => {
    setError(undefined);
    setStep((current) => (current === "details" ? "folder" : "matter"));
  }, []);

  const baseName = React.useMemo(
    () => toBaseName(details.baseName, document.extension),
    [details.baseName, document]
  );
  const fileName = baseName + document.extension;
  const fileNameError = React.useMemo(
    () => fileNameProblem(details.baseName, document.extension),
    [details.baseName, document]
  );

  const progressRef = React.useRef<ISaveProgress | undefined>(undefined);

  const submit = React.useCallback(async () => {
    if (!driveId || fileNameError) return;
    setIsSaving(true);
    setError(undefined);
    const folderId = folderPath.length ? folderPath[folderPath.length - 1].id : ROOT;
    const key = `${driveId}|${folderId}|${baseName}`;
    const progress: ISaveProgress =
      progressRef.current?.key === key ? progressRef.current : { key };
    progressRef.current = progress;
    try {
      progress.document ??= await uploadFile(
        driveId,
        folderId,
        fileName,
        await readDocument(document, "native")
      );
      if (!progress.document.version)
        progress.document = await applyProperties(driveId, progress.document, details);
      if (details.pdfCopy) {
        progress.pdf ??= await uploadFile(
          driveId,
          folderId,
          `${baseName}.pdf`,
          await readDocument(document, "pdf")
        );
        if (!progress.pdf.version)
          progress.pdf = await applyProperties(driveId, progress.pdf, details);
      }
      setResult({ document: progress.document, pdf: details.pdfCopy ? progress.pdf : undefined });
      progressRef.current = undefined;
      setStep("done");
    } catch (e) {
      const message = (e as Error).message;
      // Entered details stay; Submit finishes only what is missing (no second upload of a saved file).
      setError(
        progress.document
          ? `${progress.document.name} was saved, but the save did not finish: ${message} Select Submit to finish.`
          : message
      );
    } finally {
      setIsSaving(false);
    }
  }, [baseName, details, document, driveId, fileName, fileNameError, folderPath]);

  const saveAnother = React.useCallback(() => {
    setResult(undefined);
    setStep("details");
  }, []);

  const locationLabel = React.useMemo(
    () => [matter?.ref, ...folderPath.map((f) => f.name)].filter(Boolean).join(" › "),
    [matter, folderPath]
  );
  const canSubmit = !!driveId && !fileNameError && !isSaving;

  React.useEffect(() => {
    void loadMatters();
  }, [loadMatters]);

  return {
    state: {
      step,
      document,
      matters,
      matter,
      folderPath,
      folders,
      details,
      isLoading,
      isSaving,
      error,
      result,
    },
    actions: {
      loadMatters,
      selectMatter,
      openFolder,
      openPathLevel,
      confirmFolder,
      updateDetails,
      back,
      submit,
      saveAnother,
    },
    computed: { fileName, fileNameError, locationLabel, canSubmit },
  };
}
