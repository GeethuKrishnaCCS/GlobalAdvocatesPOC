import * as React from 'react';
import type { BundleApiService, IBuildStatus } from '../services/BundleApiService';

/** Poll interval while a build runs. */
const POLL_MS: number = 3000;
const TERMINAL: ReadonlyArray<string> = ['Completed', 'Failed', 'Terminated'];

export interface IBundleBuild {
  state: { status: IBuildStatus | undefined; isRunning: boolean; error: string | undefined };
  actions: { start: (bundleId: string) => Promise<void> };
}

/** Starts a server-side bundle build and polls it until it finishes. Polling stops when the component unmounts. */
export function useBundleBuild(api: BundleApiService): IBundleBuild {
  const [status, setStatus] = React.useState<IBuildStatus | undefined>(undefined);
  const [isRunning, setIsRunning] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const mounted = React.useRef<boolean>(true);
  React.useEffect(() => () => { mounted.current = false; }, []);

  const start = React.useCallback(async (bundleId: string): Promise<void> => {
    setIsRunning(true);
    setError(undefined);
    setStatus(undefined);
    try {
      const buildId: string = await api.startBuild(bundleId);
      let current: IBuildStatus;
      do {
        await new Promise(resolve => setTimeout(resolve, POLL_MS));
        if (!mounted.current) {
          return;
        }
        current = await api.getStatus(buildId);
        setStatus(current);
      } while (TERMINAL.indexOf(current.status) < 0);
      if (current.status !== 'Completed') {
        setError(current.error ?? current.status);
      }
    } catch (e) {
      setError((e as Error).message);
    } finally {
      if (mounted.current) {
        setIsRunning(false);
      }
    }
  }, [api]);

  return { state: { status, isRunning, error }, actions: { start } };
}
