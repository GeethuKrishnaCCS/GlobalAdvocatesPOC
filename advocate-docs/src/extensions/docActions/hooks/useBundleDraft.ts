import * as React from 'react';
import type { ISelectedDocument } from '../models/ISelectedDocument';

export interface IBundleDraft {
  state: { documents: ISelectedDocument[]; highlightedId: number | undefined };
  actions: {
    highlight: (id: number) => void;
    moveUp: (id: number) => void;
    moveDown: (id: number) => void;
    remove: (id: number) => void;
    /** Appends documents not already in the bundle; returns how many were added. */
    addMany: (docs: ReadonlyArray<ISelectedDocument>) => number;
  };
  computed: { highlighted: ISelectedDocument | undefined; count: number };
}

function move(list: ISelectedDocument[], id: number, offset: number): ISelectedDocument[] {
  const from: number = list.findIndex(d => d.id === id);
  const to: number = from + offset;
  if (from < 0 || to < 0 || to >= list.length) {
    return list;
  }
  const copy: ISelectedDocument[] = [...list];
  const [item] = copy.splice(from, 1);
  copy.splice(to, 0, item);
  return copy;
}

/** Ordered bundle membership being assembled in the bundle builder, plus the document shown in the preview pane. */
export function useBundleDraft(initial: ReadonlyArray<ISelectedDocument>): IBundleDraft {
  const [documents, setDocuments] = React.useState<ISelectedDocument[]>(() => [...initial]);
  const [highlightedId, setHighlightedId] = React.useState<number | undefined>(initial[0]?.id);

  const highlight = React.useCallback((id: number): void => setHighlightedId(id), []);
  const moveUp = React.useCallback((id: number): void => setDocuments(list => move(list, id, -1)), []);
  const moveDown = React.useCallback((id: number): void => setDocuments(list => move(list, id, 1)), []);
  const remove = React.useCallback((id: number): void => {
    setDocuments(list => {
      const next: ISelectedDocument[] = list.filter(d => d.id !== id);
      setHighlightedId(current => (current === id ? next[0]?.id : current));
      return next;
    });
  }, []);

  const addMany = React.useCallback((docs: ReadonlyArray<ISelectedDocument>): number => {
    const existing: Set<string> = new Set(documents.map(d => d.uniqueId.toLowerCase()));
    const fresh: ISelectedDocument[] = docs.filter(d => !existing.has(d.uniqueId.toLowerCase()));
    if (fresh.length) {
      setDocuments(list => [...list, ...fresh]);
      setHighlightedId(current => current ?? fresh[0].id);
    }
    return fresh.length;
  }, [documents]);

  const computed = React.useMemo(() => ({
    highlighted: documents.find(d => d.id === highlightedId),
    count: documents.length
  }), [documents, highlightedId]);

  return { state: { documents, highlightedId }, actions: { highlight, moveUp, moveDown, remove, addMany }, computed };
}
