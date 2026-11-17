import type * as React from "react";

/* global Office, window */

/**
 * For a URL only known after async work (e.g. a draft being created): call this synchronously in the click handler,
 * then pass the URL (or undefined to give up) to the function it returns, which reports whether it was opened.
 * Office on the web opens a blank window now, inside the click, so popup blockers allow it; desktop task panes use
 * the system browser later, which needs no click.
 */
export function openWhenReady(): (url: string | undefined) => boolean {
  if (typeof Office.context.ui?.openBrowserWindow === "function") {
    return (url) => {
      if (url) Office.context.ui.openBrowserWindow(url);
      return !!url;
    };
  }
  const opened = window.open("", "_blank");
  if (opened) opened.opener = null; // Outlook needs no handle back to the pane
  return (url) => {
    if (!opened || opened.closed) return false;
    if (url) opened.location.href = url;
    else opened.close();
    return !!url;
  };
}

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

/**
 * Opens the Outlook app (new Outlook registers ms-outlook:). It can only bring Outlook to the front: opening a given
 * draft is "not supported yet" in new Outlook (spike 2026-11-09), so the pane points the user to Drafts. Office on
 * Windows asks once whether the add-in may open this kind of link (admins can pre-allow it, see README).
 */
export const OUTLOOK_APP_URL = "ms-outlook://";

/** True for Office desktop platforms (Windows "PC", Mac); Office on the web is "OfficeOnline". Pure. */
export function isDesktopPlatform(platform: Office.PlatformType | string | undefined): boolean {
  return platform === "PC" || platform === "Mac";
}

/** True in Office desktop apps, where links open outside the pane. */
export function isDesktopHost(): boolean {
  return isDesktopPlatform(Office.context.platform);
}
