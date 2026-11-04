import * as React from "react";
import { FILING_FOLDER_PATH, MATTER_CATEGORY_PREFIX } from "../config";
import { findFolder, getMatterDriveId, listMatters, type IMatter } from "../services/matterService";
import { getCurrentDocument, readDocument, type DocumentFormat } from "../services/documentSource";
import {
  attachFile,
  createDraft,
  updateDraft,
  type IDraft,
  type IDraftRequest,
} from "../services/mailDraftService";
import { fileNameProblem, toBaseName } from "../services/fileName";
import { taggedSubject } from "../services/mailText";
import { openWhenReady } from "../openInBrowser";

/* global Office, Blob */

export type AttachFormat = "native" | "pdf" | "both";

/** One file that will be attached. */
export interface IAttachment {
  format: DocumentFormat;
  /** Base name the user edits; the extension is locked. */
  baseName: string;
  extension: string;
  fileName: string;
  problem: string | undefined;
}

/** Editable fields of the form. */
export interface ISendForm {
  format: AttachFormat;
  nativeName: string;
  pdfName: string;
  tagMatter: boolean;
  fileSentEmail: boolean;
}

export interface ISendResult {
  draft: IDraft;
  attachmentNames: string[];
  /** False when the browser blocked the new window; the result screen's link opens it instead. */
  opened: boolean;
}

/** Filing target of the selected matter: "driveId|itemId", null when the matter has no such folder. */
type FilingTarget = string | null | undefined;

export interface ISendAsAttachment {
  state: {
    form: ISendForm;
    matters: IMatter[] | undefined;
    matter: IMatter | undefined;
    filingTarget: FilingTarget;
    isLoading: boolean;
    isCreating: boolean;
    retryPending: boolean;
    error: string | undefined;
    result: ISendResult | undefined;
  };
  actions: {
    update: (patch: Partial<ISendForm>) => void;
    selectMatter: (matter: IMatter) => Promise<void>;
    loadMatters: () => Promise<void>;
    create: () => Promise<void>;
    startOver: () => void;
  };
  computed: {
    fileLabel: string;
    fileLocation: string;
    attachments: IAttachment[];
    subjectToSend: string;
    canCreate: boolean;
    fileEmail: boolean;
    summary: string;
  };
}

/** A draft whose attachments did not all go in; a retry updates it and adds the rest. */
interface IDraftProgress {
  draft: IDraft;
  attached: Set<string>;
}

/** Send as Attachment: an Outlook draft with the open document attached as is (Word/Excel/PowerPoint), PDF or both,
 * opened straight in Outlook; recipients and message are written there. */
export function useSendAsAttachment(host: Office.HostType): ISendAsAttachment {
  const document = React.useMemo(() => getCurrentDocument(host), [host]);
  const [form, setForm] = React.useState<ISendForm>({
    format: "native",
    nativeName: document.baseName,
    pdfName: document.baseName,
    tagMatter: true,
    fileSentEmail: false,
  });
  const [matters, setMatters] = React.useState<IMatter[] | undefined>(undefined);
  const [matter, setMatter] = React.useState<IMatter | undefined>(undefined);
  const [filingTarget, setFilingTarget] = React.useState<FilingTarget>(undefined);
  const [isLoading, setIsLoading] = React.useState(false);
  const [isCreating, setIsCreating] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [result, setResult] = React.useState<ISendResult | undefined>(undefined);
  const progressRef = React.useRef<IDraftProgress | undefined>(undefined);
  /** A draft exists but lacks attachments: what is attached, and the filing marker, are locked until it is finished. */
  const [retryPending, setRetryPending] = React.useState(false);

  const update = React.useCallback(
    (patch: Partial<ISendForm>) => setForm((f) => ({ ...f, ...patch })),
    []
  );

  /** Latest matter selection; an older folder lookup that finishes later must not win (wrong matter's folder). */
  const selectionRef = React.useRef(0);

  /** Selects a matter and resolves its Correspondence › Matter folder for the "file sent email" option. */
  const selectMatter = React.useCallback(async (selected: IMatter) => {
    const selection = ++selectionRef.current;
    setMatter(selected);
    setFilingTarget(undefined);
    setIsLoading(true);
    setError(undefined);
    let target: FilingTarget = null;
    try {
      const driveId = await getMatterDriveId(selected.siteId);
      const folderId = await findFolder(driveId, FILING_FOLDER_PATH);
      target = folderId ? `${driveId}|${folderId}` : null;
    } catch (e) {
      if (selection === selectionRef.current) setError((e as Error).message);
    }
    if (selection !== selectionRef.current) return;
    setFilingTarget(target);
    setIsLoading(false);
  }, []);

  /** Loads the matters and pre-selects the one the open document lives in (by site URL, so a provisioned matter
   * reference that differs from the site alias still matches). */
  const loadMatters = React.useCallback(async () => {
    setIsLoading(true);
    setError(undefined);
    let list: IMatter[];
    try {
      list = await listMatters();
    } catch (e) {
      setError((e as Error).message);
      setIsLoading(false);
      return;
    }
    setMatters(list);
    setIsLoading(false);
    const url = document.url.toLowerCase();
    const own = list.find((m) => m.webUrl && url.startsWith(`${m.webUrl.toLowerCase()}/`));
    if (own) await selectMatter(own);
  }, [document, selectMatter]);

  React.useEffect(() => {
    void loadMatters();
  }, [loadMatters]);

  const attachments = React.useMemo((): IAttachment[] => {
    const build = (format: DocumentFormat, name: string, extension: string): IAttachment => ({
      format,
      baseName: name,
      extension,
      fileName: toBaseName(name, extension) + extension,
      problem: fileNameProblem(name, extension),
    });
    const list: IAttachment[] = [];
    if (form.format !== "pdf") list.push(build("native", form.nativeName, document.extension));
    if (form.format !== "native") list.push(build("pdf", form.pdfName, ".pdf"));
    return list;
  }, [document, form.format, form.nativeName, form.pdfName]);

  const matterRef = form.tagMatter ? matter?.ref : undefined;
  const subjectToSend = taggedSubject(document.baseName, matterRef);
  const canCreate =
    !isCreating && attachments.every((a) => !a.problem) && (!form.tagMatter || !!matter);
  /** Filing applies only with a tagged matter that has the folder; a ticked box for another matter is ignored. */
  const fileEmail = form.tagMatter && form.fileSentEmail && !!filingTarget;

  const create = React.useCallback(async () => {
    if (!canCreate) return;
    // Before any await: the window must open inside the click.
    const open = openWhenReady();
    setIsCreating(true);
    setError(undefined);
    const request: IDraftRequest = {
      subject: subjectToSend,
      categories: matterRef ? [`${MATTER_CATEGORY_PREFIX}${matterRef}`] : [],
      fileToFolder: fileEmail ? filingTarget! : undefined,
    };
    try {
      // Read the document first, so a failure leaves no empty draft. One export at a time: Office limits open files.
      const files: { name: string; body: Blob }[] = [];
      for (const a of attachments.filter((x) => !progressRef.current?.attached.has(x.fileName))) {
        files.push({ name: a.fileName, body: await readDocument(document, a.format) });
      }
      if (progressRef.current) await updateDraft(progressRef.current.draft.id, request);
      else progressRef.current = { draft: await createDraft(request), attached: new Set() };
      const progress = progressRef.current;
      for (const file of files) {
        await attachFile(progress.draft.id, file.name, file.body);
        progress.attached.add(file.name);
      }
      const opened = open(progress.draft.webLink || undefined);
      setResult({
        draft: progress.draft,
        attachmentNames: attachments.map((a) => a.fileName),
        opened,
      });
      progressRef.current = undefined;
      setRetryPending(false);
    } catch (e) {
      open(undefined);
      const message = (e as Error).message;
      setRetryPending(!!progressRef.current);
      setError(
        progressRef.current
          ? `The draft was created, but not every attachment was added: ${message} Select Open in Outlook to finish.`
          : message
      );
    } finally {
      setIsCreating(false);
    }
  }, [attachments, canCreate, document, fileEmail, filingTarget, matterRef, subjectToSend]);

  const startOver = React.useCallback(() => {
    setResult(undefined);
    setRetryPending(false);
  }, []);

  const summary = `${attachments.length} attachment${attachments.length === 1 ? "" : "s"} · ${subjectToSend}`;

  return {
    state: {
      form,
      matters,
      matter,
      filingTarget,
      isLoading,
      isCreating,
      retryPending,
      error,
      result,
    },
    actions: { update, selectMatter, loadMatters, create, startOver },
    computed: {
      fileLabel: document.baseName + document.extension,
      fileLocation: document.url
        ? ["Open document", document.folderName].filter(Boolean).join(" · ")
        : "Not saved yet",
      attachments,
      subjectToSend,
      canCreate,
      fileEmail,
      summary,
    },
  };
}
