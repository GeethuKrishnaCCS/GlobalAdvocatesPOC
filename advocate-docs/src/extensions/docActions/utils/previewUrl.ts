/**
 * SharePoint's file viewer in embed mode. Same origin as the library, so it can be framed without
 * Graph permissions; shows the full document (all pages).
 * Spike 2026-10-14 on ccsdev01: renders docx, pdf, txt and eml. xlsx stays blank (Doc.aspx embedview /
 * interactivepreview blank too, and embedview also blanks docx) — xlsx needs the converted-PDF fallback.
 */
export function getEmbedPreviewUrl(webAbsoluteUrl: string, uniqueId: string): string {
  return `${webAbsoluteUrl}/_layouts/15/embed.aspx?UniqueId=${encodeURIComponent(uniqueId)}`;
}
