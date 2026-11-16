import * as React from "react";
import { friendlyAuthError } from "../auth/authErrors";
import { getMatterFolderStatus, type IMatterFolderStatus } from "../services/matterFolderService";

/* global document, window */

/** How often the backend's filing status is re-read while the pane is open. */
export const STATUS_POLL_MS = 60_000;

export interface IMatterFolderStatusState {
  /** Undefined until the first successful read. */
  status: IMatterFolderStatus[] | undefined;
  checkedAt: Date | undefined;
  error: string | undefined;
}

/**
 * Reads the mailbox-folder status now and every minute while `enabled` and the pane is visible; stops on unmount.
 * A change of `refreshKey` (e.g. a matter got registered) reads again straight away.
 */
export function useMatterFolderStatus(
  enabled: boolean,
  refreshKey: unknown
): { state: IMatterFolderStatusState } {
  const [state, setState] = React.useState<IMatterFolderStatusState>({
    status: undefined,
    checkedAt: undefined,
    error: undefined,
  });

  React.useEffect(() => {
    if (!enabled) return undefined;
    let active = true;
    const poll = (): void => {
      if (document.hidden) return;
      getMatterFolderStatus().then(
        (status) => active && setState({ status, checkedAt: new Date(), error: undefined }),
        (e) => active && setState((prev) => ({ ...prev, error: friendlyAuthError(e) }))
      );
    };
    poll();
    const timer = window.setInterval(poll, STATUS_POLL_MS);
    return () => {
      active = false;
      window.clearInterval(timer);
    };
  }, [enabled, refreshKey]);

  return { state };
}
