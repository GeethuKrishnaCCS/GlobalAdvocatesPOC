import * as React from "react";
import { friendlyAuthError } from "../auth/authErrors";
import { registerMatterFolders, unregisterMatterFolders } from "../services/matterFolderService";
import { loadPins, savePins } from "../services/matterPins";
import type { IMatter } from "../services/matterService";

export interface IMatterPins {
  state: {
    /** References of the pinned matters. */
    pinned: ReadonlySet<string>;
    error: string | undefined;
    /** References whose mailbox folders the backend confirmed this session. */
    withMailFolder: ReadonlySet<string>;
    /** Why the mailbox folders could not be set up (never blocks pinning); undefined when fine. */
    folderError: string | undefined;
  };
  actions: {
    toggle: (matter: IMatter) => void;
    /** Registers every pinned matter of `matters`, once per pane load (refreshes sub-folders, migrates old pins). */
    syncPinned: (matters: IMatter[]) => void;
    /** Re-runs the registration (from a click, so the sign-in dialog may open). */
    retry: () => void;
  };
}

/**
 * The user's pinned matters (Outlook panes): shown first in the matter list, stored in the mailbox. Each pinned matter
 * is also registered with the mailbox-folder backend; a backend failure is reported in `folderError`, never undoes a pin.
 */
export function useMatterPins(): IMatterPins {
  const [pinned, setPinned] = React.useState<ReadonlySet<string>>(() => new Set(loadPins()));
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [withMailFolder, setWithMailFolder] = React.useState<ReadonlySet<string>>(() => new Set());
  const [folderError, setFolderError] = React.useState<string | undefined>(undefined);
  const synced = React.useRef(false);
  const knownMatters = React.useRef<IMatter[]>([]);

  const register = React.useCallback(async (matter: IMatter, allowPopup: boolean) => {
    try {
      await registerMatterFolders(matter, allowPopup);
      setWithMailFolder((prev) => new Set(prev).add(matter.ref));
    } catch (e) {
      setFolderError(`${matter.ref}: ${friendlyAuthError(e)}`);
    }
  }, []);

  const unregister = React.useCallback(async (ref: string) => {
    setWithMailFolder((prev) => {
      const next = new Set(prev);
      next.delete(ref);
      return next;
    });
    try {
      await unregisterMatterFolders(ref);
    } catch (e) {
      setFolderError(`${ref}: ${friendlyAuthError(e)}`);
    }
  }, []);

  /** One after another: each creates mail folders in the mailbox. */
  const registerPinned = React.useCallback(
    async (allowPopup: boolean) => {
      setFolderError(undefined);
      for (const matter of knownMatters.current.filter((m) => pinned.has(m.ref)))
        await register(matter, allowPopup);
    },
    [pinned, register]
  );

  const toggle = React.useCallback(
    (matter: IMatter) => {
      const next = new Set(pinned);
      const adding = !next.delete(matter.ref);
      if (adding) next.add(matter.ref);
      setPinned(next);
      setError(undefined);
      setFolderError(undefined);
      savePins([...next]).then(
        () => (adding ? register(matter, false) : unregister(matter.ref)),
        (e: Error) => {
          // Not stored: show what is really saved, so the pin does not silently vanish on the next open.
          setPinned(pinned);
          setError(`Your pinned matters could not be saved: ${e.message}`);
        }
      );
    },
    [pinned, register, unregister]
  );

  const syncPinned = React.useCallback(
    (matters: IMatter[]) => {
      knownMatters.current = matters;
      if (synced.current) return;
      synced.current = true;
      void registerPinned(false);
    },
    [registerPinned]
  );

  const retry = React.useCallback(() => void registerPinned(true), [registerPinned]);

  return {
    state: { pinned, error, withMailFolder, folderError },
    actions: { toggle, syncPinned, retry },
  };
}
