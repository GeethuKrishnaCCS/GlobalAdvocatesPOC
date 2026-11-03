import type * as React from "react";

/* global Office */

/**
 * Click handler for links to SharePoint / Outlook: desktop task panes would open target=_blank inside the add-in's
 * own webview, so use the system browser there. Office on the web has no openBrowserWindow; the link's normal
 * target=_blank navigation runs instead.
 */
export function openInBrowser(event: React.MouseEvent, url: string): void {
  if (typeof Office.context.ui?.openBrowserWindow !== "function") return;
  event.preventDefault();
  Office.context.ui.openBrowserWindow(url);
}
