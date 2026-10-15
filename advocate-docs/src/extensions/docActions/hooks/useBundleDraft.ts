import * as React from 'react';
import type { ISelectedDocument } from '../models/ISelectedDocument';
import type { IBundle, IBundleHeader } from '../models/IBundle';
import type { BundleService } from '../services/BundleService';

export interface IBundleDraft {
  state: {
    documents: ISelectedDocument[];
    highlightedId: string | undefined;
    title: string;
    header: IBundleHeader | undefined;
    isSaving: boolean;
    saveError: string | undefined;
  };
  actions: {
    highlight: (uniqueId: string) => void;
    moveUp: (uniqueId: string) => void;
    moveDown: (uniqueId: string) => void;
    remove: (uniqueId: string) => void;
    /** Appends documents not already in the bundle; returns how many were added. */
    addMany: (docs: ReadonlyArray<ISelectedDocument>) => number;
    setTitle: (title: string) => void;
    save: () => Promise<boolean>;
  };
  computed: { highlighted: ISelectedDocument | undefined; count: number; isDirty: boolean; canSave: boolean };
}

function move(list: ISelectedDocument[], uniqueId: string, offset: number): ISelectedDocument[] {
  const from: number = list.findIndex(d => d.uniqueId === uniqueId);
  const to: number = from + offset;
  if (from < 0 || to < 0 || to >= list.length) {
    return list;
  }
  const copy: ISelectedDocument[] = [...list];
  const [item] = copy.splice(from, 1);
  copy.splice(to, 0, item);
  return copy;
}

/** Order + title fingerprint, to tell whether the draft differs from what was last saved. */
function snapshot(title: string, docs: ReadonlyArray<ISelectedDocument>): string {
  return `${title}\n${docs.map(d => d.uniqueId).join('|')}`;
}

/**
 * Bundle being assembled or edited: ordered members, name, the saved header (once saved) and
 * whether there are unsaved changes. Starts from a selection (new bundle) or a saved bundle.
 */
export function useBundleDraft(
  service: BundleService,
  initialDocuments: ReadonlyArray<ISelectedDocument>,
  initialBundle?: IBundle
): IBundleDraft {
  const startDocs: ISelectedDocument[] = initialBundle ? initialBundle.documents : [...initialDocuments];
  const [documents, setDocuments] = React.useState<ISelectedDocument[]>(startDocs);
  const [highlightedId, setHighlightedId] = React.useState<string | undefined>(startDocs[0]?.uniqueId);
  const [title, setTitle] = React.useState<string>(initialBundle?.header.title ?? '');
  const [header, setHeader] = React.useState<IBundleHeader | undefined>(initialBundle?.header);
  const [savedSnapshot, setSavedSnapshot] = React.useState<string | undefined>(
    initialBundle ? snapshot(initialBundle.header.title, initialBundle.documents) : undefined);
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [saveError, setSaveError] = React.useState<string | undefined>(undefined);

  const highlight = React.useCallback((uniqueId: string): void => setHighlightedId(uniqueId), []);
  const moveUp = React.useCallback((uniqueId: string): void => setDocuments(list => move(list, uniqueId, -1)), []);
  const moveDown = React.useCallback((uniqueId: string): void => setDocuments(list => move(list, uniqueId, 1)), []);
  const remove = React.useCallback((uniqueId: string): void => {
    setDocuments(list => {
      const next: ISelectedDocument[] = list.filter(d => d.uniqueId !== uniqueId);
      setHighlightedId(current => (current === uniqueId ? next[0]?.uniqueId : current));
      return next;
    });
  }, []);

  const addMany = React.useCallback((docs: ReadonlyArray<ISelectedDocument>): number => {
    const existing: Set<string> = new Set(documents.map(d => d.uniqueId));
    const fresh: ISelectedDocument[] = docs.filter(d => !existing.has(d.uniqueId));
    if (fresh.length) {
      setDocuments(list => [...list, ...fresh]);
      setHighlightedId(current => current ?? fresh[0].uniqueId);
    }
    return fresh.length;
  }, [documents]);

  const save = React.useCallback(async (): Promise<boolean> => {
    setIsSaving(true);
    setSaveError(undefined);
    try {
      const base: IBundleHeader = header ?? { bundleId: '', title, status: 'Draft', memberCount: 0 };
      const saved: IBundleHeader = await service.saveBundle({ ...base, title: title.trim() }, documents);
      setHeader(saved);
      setSavedSnapshot(snapshot(title.trim(), documents));
      return true;
    } catch (e) {
      setSaveError((e as Error).message);
      return false;
    } finally {
      setIsSaving(false);
    }
  }, [service, header, title, documents]);

  const computed = React.useMemo(() => {
    const current: string = snapshot(title.trim(), documents);
    // A new bundle with nothing in it isn't "dirty"; anything else unsaved is.
    const isDirty: boolean = savedSnapshot === undefined ? documents.length > 0 || title.trim().length > 0 : current !== savedSnapshot;
    return {
      highlighted: documents.find(d => d.uniqueId === highlightedId),
      count: documents.length,
      isDirty,
      canSave: title.trim().length > 0 && documents.length > 0 && isDirty
    };
  }, [documents, highlightedId, title, savedSnapshot]);

  return {
    state: { documents, highlightedId, title, header, isSaving, saveError },
    actions: { highlight, moveUp, moveDown, remove, addMany, setTitle, save },
    computed
  };
}
