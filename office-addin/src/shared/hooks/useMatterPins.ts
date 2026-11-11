import * as React from "react";
import { loadPins, savePins } from "../services/matterPins";

export interface IMatterPins {
  state: {
    /** References of the pinned matters. */
    pinned: ReadonlySet<string>;
    error: string | undefined;
  };
  actions: { toggle: (ref: string) => void };
}

/** The user's pinned matters (Outlook panes): shown first in the matter list, stored in the mailbox. */
export function useMatterPins(): IMatterPins {
  const [pinned, setPinned] = React.useState<ReadonlySet<string>>(() => new Set(loadPins()));
  const [error, setError] = React.useState<string | undefined>(undefined);

  const toggle = React.useCallback(
    (ref: string) => {
      const next = new Set(pinned);
      if (!next.delete(ref)) next.add(ref);
      setPinned(next);
      setError(undefined);
      savePins([...next]).catch((e: Error) => {
        // Not stored: show what is really saved, so the pin does not silently vanish on the next open.
        setPinned(pinned);
        setError(`Your pinned matters could not be saved: ${e.message}`);
      });
    },
    [pinned]
  );

  return { state: { pinned, error }, actions: { toggle } };
}
