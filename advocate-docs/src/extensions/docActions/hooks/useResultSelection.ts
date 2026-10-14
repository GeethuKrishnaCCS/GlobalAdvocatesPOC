import * as React from 'react';
import type { ISelectedDocument } from '../models/ISelectedDocument';

export interface IResultSelection {
  state: { selectedIds: ReadonlySet<string> };
  actions: {
    toggle: (uniqueId: string, selected: boolean) => void;
    selectAll: (selected: boolean) => void;
  };
  computed: {
    /** Results not yet in the bundle — the only ones that can be picked. */
    selectable: ISelectedDocument[];
    selected: ISelectedDocument[];
    allSelected: boolean;
    someSelected: boolean;
  };
}

/** Per-document picking within a finder result; resets whenever the result changes. */
export function useResultSelection(
  results: ReadonlyArray<ISelectedDocument> | undefined,
  existingIds: ReadonlySet<string>
): IResultSelection {
  const [selectedIds, setSelectedIds] = React.useState<ReadonlySet<string>>(new Set());

  React.useEffect(() => { setSelectedIds(new Set()); }, [results]);

  const selectable: ISelectedDocument[] = React.useMemo(
    () => (results ?? []).filter(d => !existingIds.has(d.uniqueId)),
    [results, existingIds]
  );

  const toggle = React.useCallback((uniqueId: string, selected: boolean): void => {
    setSelectedIds(current => {
      const next: Set<string> = new Set(current);
      if (selected) { next.add(uniqueId); } else { next.delete(uniqueId); }
      return next;
    });
  }, []);

  const selectAll = React.useCallback((selected: boolean): void => {
    setSelectedIds(selected ? new Set(selectable.map(d => d.uniqueId)) : new Set());
  }, [selectable]);

  const computed = React.useMemo(() => {
    const selected: ISelectedDocument[] = selectable.filter(d => selectedIds.has(d.uniqueId));
    return {
      selectable,
      selected,
      allSelected: selectable.length > 0 && selected.length === selectable.length,
      someSelected: selected.length > 0 && selected.length < selectable.length
    };
  }, [selectable, selectedIds]);

  return { state: { selectedIds }, actions: { toggle, selectAll }, computed };
}
