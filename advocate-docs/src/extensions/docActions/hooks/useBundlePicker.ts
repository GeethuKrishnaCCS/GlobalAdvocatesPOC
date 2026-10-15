import * as React from 'react';
import type { BundleService } from '../services/BundleService';
import type { IAddToBundleResult, IBundle, IBundleHeader } from '../models/IBundle';
import type { ISelectedDocument } from '../models/ISelectedDocument';

export interface IBundlePicker {
  state: {
    bundles: IBundleHeader[] | undefined;
    selectedId: string | undefined;
    isBusy: boolean;
    error: string | undefined;
    addResult: IAddToBundleResult | undefined;
  };
  actions: {
    select: (bundleId: string) => void;
    open: () => Promise<IBundle | undefined>;
    add: (docs: ReadonlyArray<ISelectedDocument>) => Promise<void>;
  };
  computed: { selected: IBundleHeader | undefined };
}

/** Loads saved bundles and either opens one or appends documents to one. */
export function useBundlePicker(service: BundleService): IBundlePicker {
  const [bundles, setBundles] = React.useState<IBundleHeader[] | undefined>(undefined);
  const [selectedId, setSelectedId] = React.useState<string | undefined>(undefined);
  const [isBusy, setIsBusy] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [addResult, setAddResult] = React.useState<IAddToBundleResult | undefined>(undefined);

  React.useEffect(() => {
    let cancelled: boolean = false;
    service.listBundles()
      .then(list => { if (!cancelled) { setBundles(list); } })
      .catch((e: Error) => { if (!cancelled) { setError(e.message); setBundles([]); } });
    return () => { cancelled = true; };
  }, [service]);

  const selected: IBundleHeader | undefined = React.useMemo(
    () => bundles?.find(b => b.bundleId === selectedId), [bundles, selectedId]);

  const run = React.useCallback(async <T>(work: () => Promise<T>): Promise<T | undefined> => {
    setIsBusy(true);
    setError(undefined);
    try {
      return await work();
    } catch (e) {
      setError((e as Error).message);
      return undefined;
    } finally {
      setIsBusy(false);
    }
  }, []);

  const open = React.useCallback(
    (): Promise<IBundle | undefined> => (selected ? run(() => service.getBundle(selected)) : Promise.resolve(undefined)),
    [run, selected, service]);

  const add = React.useCallback(async (docs: ReadonlyArray<ISelectedDocument>): Promise<void> => {
    if (!selected) {
      return;
    }
    const result: IAddToBundleResult | undefined = await run(() => service.addToBundle(selected, docs));
    if (result) {
      setAddResult(result);
    }
  }, [run, selected, service]);

  return {
    state: { bundles, selectedId, isBusy, error, addResult },
    actions: { select: setSelectedId, open, add },
    computed: { selected }
  };
}
