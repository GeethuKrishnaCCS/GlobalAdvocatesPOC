import * as React from "react";

/* global Office */

export interface IMailItem {
  /** Changes whenever Outlook shows another item in a pinned pane; use as a React key to start over. */
  itemKey: number;
  /** False when the pinned pane has no item (nothing selected, or several). */
  hasItem: boolean;
}

/** Tracks the selected email for a pinned pane (Outlook keeps it open and fires ItemChanged on selection). */
export function useMailItem(): IMailItem {
  const [itemKey, setItemKey] = React.useState(0);
  const [hasItem, setHasItem] = React.useState(() => !!Office.context.mailbox?.item);

  React.useEffect(() => {
    const onChanged = (): void => {
      setHasItem(!!Office.context.mailbox.item);
      setItemKey((key) => key + 1);
    };
    Office.context.mailbox?.addHandlerAsync(Office.EventType.ItemChanged, onChanged);
    return () => {
      Office.context.mailbox?.removeHandlerAsync(Office.EventType.ItemChanged);
    };
  }, []);

  return { itemKey, hasItem };
}
