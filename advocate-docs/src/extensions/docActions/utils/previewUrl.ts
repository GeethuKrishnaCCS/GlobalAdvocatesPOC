/**
 * SharePoint's file viewer in embed mode. Same origin as the library, so it can be framed without
 * Graph permissions; shows the full document (all pages).
 * Spike 2026-10-14 on ccsdev01: renders docx, pdf, txt and eml. xlsx stays blank (Doc.aspx embedview /
 * interactivepreview blank too, and embedview also blanks docx) — xlsx needs the converted-PDF fallback.
 * Visio is the exception (2026-10-16): embed.aspx is blank, Doc.aspx embedview renders it (Visio for the web).
 */
export function getEmbedPreviewUrl(webAbsoluteUrl: string, uniqueId: string, fileType: string): string {
  return isVisio(fileType)
    ? `${webAbsoluteUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(`{${uniqueId}}`)}&action=embedview`
    : `${webAbsoluteUrl}/_layouts/15/embed.aspx?UniqueId=${encodeURIComponent(uniqueId)}`;
}

/** Visio drawings: no PDF conversion in Microsoft 365, but Visio for the web views and prints them. */
export function isVisio(fileType: string): boolean {
  return fileType === 'vsdx' || fileType === 'vsd';
}
