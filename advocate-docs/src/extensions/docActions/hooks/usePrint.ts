import * as React from 'react';
import type { BundleApiService } from '../services/BundleApiService';

/**
 * Fetches the print PDF for the selected documents and opens the browser's print dialog on it
 * (hidden iframe on a blob URL, same origin, so contentWindow.print() is allowed).
 */
export function usePrint(api: BundleApiService, itemIds: number[]): {
  state: { pdfUrl: string | undefined; error: string | undefined; isLoading: boolean };
  actions: { print: () => void; onFrameLoad: () => void };
  frameRef: React.RefObject<HTMLIFrameElement>;
} {
  const [pdfUrl, setPdfUrl] = React.useState<string | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const frameRef = React.useRef<HTMLIFrameElement>(null);

  React.useEffect(() => {
    let url: string | undefined;
    let cancelled = false;
    api.print(itemIds)
      .then(blob => {
        url = URL.createObjectURL(blob);
        if (!cancelled) { setPdfUrl(url); }
      })
      .catch((e: Error) => { if (!cancelled) { setError(e.message); } });
    return () => {
      cancelled = true;
      if (url) { URL.revokeObjectURL(url); }
    };
  }, [api, itemIds]);

  const print = React.useCallback((): void => frameRef.current?.contentWindow?.print(), []);

  return {
    state: { pdfUrl, error, isLoading: !pdfUrl && !error },
    actions: { print, onFrameLoad: print },
    frameRef
  };
}
