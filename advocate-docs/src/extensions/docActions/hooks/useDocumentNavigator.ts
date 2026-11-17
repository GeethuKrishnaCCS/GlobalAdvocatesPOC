import * as React from 'react';
import type { ISelectedDocument } from '../models/ISelectedDocument';

export interface IDocumentNavigator {
  state: { index: number };
  actions: { next: () => void; previous: () => void; goTo: (index: number) => void };
  computed: { current: ISelectedDocument | undefined; hasNext: boolean; hasPrevious: boolean; position: number; total: number };
}

/** Steps through a fixed list of documents, one at a time (Preview dialog). */
export function useDocumentNavigator(documents: ReadonlyArray<ISelectedDocument>): IDocumentNavigator {
  const [index, setIndex] = React.useState<number>(0);
  const total: number = documents.length;

  const goTo = React.useCallback((target: number): void => {
    setIndex(Math.max(0, Math.min(total - 1, target)));
  }, [total]);
  const next = React.useCallback((): void => goTo(index + 1), [goTo, index]);
  const previous = React.useCallback((): void => goTo(index - 1), [goTo, index]);

  const computed = React.useMemo(() => ({
    current: documents[index],
    hasNext: index < total - 1,
    hasPrevious: index > 0,
    position: index + 1,
    total
  }), [documents, index, total]);

  return { state: { index }, actions: { next, previous, goTo }, computed };
}
