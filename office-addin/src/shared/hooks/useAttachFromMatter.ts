import * as React from "react";
import {
  downloadFile,
  searchFiles,
  type IMatter,
  type IMatterFile,
} from "../services/matterService";
import { attachProblem, bytesToBase64, isAlreadyAttached } from "../services/mailFiling";
import { attachToEmail, getAttachedNames } from "../services/outlookItem";
import { friendlyAuthError } from "../auth/authErrors";
import { useMatterFolders, type IMatterFolders } from "./useMatterFolders";

export type AttachStep = "matter" | "documents";

export interface IAttachFromMatter {
  state: {
    step: AttachStep;
    browser: IMatterFolders;
    query: string;
    /** Search hits; undefined while browsing folders. */
    results: IMatterFile[] | undefined;
    ticked: ReadonlySet<string>;
    isSearching: boolean;
    isAttaching: boolean;
    error: string | undefined;
    /** Names attached by the last "Attach to email". */
    attachedNow: string[];
  };
  actions: {
    selectMatter: (matter: IMatter) => Promise<void>;
    back: () => void;
    setQuery: (query: string) => void;
    search: () => Promise<void>;
    toggle: (file: IMatterFile, on: boolean) => void;
    attach: () => Promise<void>;
  };
  computed: {
    shownFiles: IMatterFile[];
    isAttached: (file: IMatterFile) => boolean;
    canAttach: boolean;
  };
}

/** Attach from Matter (Outlook compose): Matter → Documents (browse or search) → ticked files attached to the email. */
export function useAttachFromMatter(): IAttachFromMatter {
  const browser = useMatterFolders();
  const { loadMatters, selectMatter: openMatter, clearMatter } = browser.actions;
  const [step, setStep] = React.useState<AttachStep>("matter");
  const [query, setQuery] = React.useState("");
  const [results, setResults] = React.useState<IMatterFile[] | undefined>(undefined);
  const [ticked, setTicked] = React.useState<ReadonlySet<string>>(new Set());
  const [isSearching, setIsSearching] = React.useState(false);
  const [isAttaching, setIsAttaching] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [attachedNames, setAttachedNames] = React.useState<string[]>([]);
  const [attachedNow, setAttachedNow] = React.useState<string[]>([]);
  /** Every file shown so far, so ticks survive moving between folders and search. */
  const known = React.useRef(new Map<string, IMatterFile>());

  const refreshAttached = React.useCallback(async () => {
    try {
      setAttachedNames(await getAttachedNames());
    } catch {
      // Not knowing what is attached only loses the duplicate marking; attaching still works.
    }
  }, []);

  React.useEffect(() => {
    void loadMatters();
    void refreshAttached();
  }, [loadMatters, refreshAttached]);

  const selectMatter = React.useCallback(
    async (matter: IMatter) => {
      setTicked(new Set());
      setResults(undefined);
      setQuery("");
      if (await openMatter(matter)) setStep("documents");
    },
    [openMatter]
  );

  const back = React.useCallback(() => {
    setError(undefined);
    setTicked(new Set());
    setResults(undefined);
    clearMatter();
    setStep("matter");
  }, [clearMatter]);

  const driveId = browser.state.location?.driveId;

  const search = React.useCallback(async () => {
    const q = query.trim();
    if (!q) {
      setResults(undefined);
      return;
    }
    if (!driveId) return;
    setIsSearching(true);
    setError(undefined);
    try {
      setResults(await searchFiles(driveId, q));
    } catch (e) {
      setError(friendlyAuthError(e));
    } finally {
      setIsSearching(false);
    }
  }, [driveId, query]);

  const updateQuery = React.useCallback((value: string) => {
    setQuery(value);
    if (!value.trim()) setResults(undefined);
  }, []);

  const toggle = React.useCallback((file: IMatterFile, on: boolean) => {
    known.current.set(file.id, file);
    setTicked((current) => {
      const next = new Set(current);
      if (on) next.add(file.id);
      else next.delete(file.id);
      return next;
    });
  }, []);

  const attach = React.useCallback(async () => {
    if (!driveId) return;
    setIsAttaching(true);
    setError(undefined);
    const done: string[] = [];
    const problems: string[] = [];
    for (const id of ticked) {
      const file = known.current.get(id);
      if (!file) continue;
      const tooBig = attachProblem(file.name, file.size);
      if (tooBig) {
        problems.push(tooBig);
        continue;
      }
      try {
        await attachToEmail(file.name, bytesToBase64(await downloadFile(driveId, file.id)));
        done.push(file.name);
      } catch (e) {
        problems.push(`${file.name}: ${friendlyAuthError(e)}`);
      }
    }
    // Attached ones leave the selection; failed ones stay ticked for another try.
    setTicked(
      (current) =>
        new Set([...current].filter((id) => !done.includes(known.current.get(id)?.name ?? "")))
    );
    setAttachedNow(done);
    if (problems.length) setError(problems.join(" "));
    await refreshAttached();
    setIsAttaching(false);
  }, [driveId, refreshAttached, ticked]);

  const isAttached = React.useCallback(
    (file: IMatterFile) => isAlreadyAttached(file.name, attachedNames),
    [attachedNames]
  );
  const shownFiles = results ?? browser.state.files ?? [];

  return {
    state: { step, browser, query, results, ticked, isSearching, isAttaching, error, attachedNow },
    actions: { selectMatter, back, setQuery: updateQuery, search, toggle, attach },
    computed: { shownFiles, isAttached, canAttach: ticked.size > 0 && !isAttaching && !!driveId },
  };
}
