/** A document picked from the list view selection, reduced to what the panels need. */
export interface ISelectedDocument {
  /** List item id. */
  id: number;
  /** File unique id (GUID, no braces) — drives the preview URL. */
  uniqueId: string;
  /** File name with extension. */
  name: string;
  /** Lower-case extension without the dot, e.g. "docx". */
  fileType: string;
  /** Server-relative file URL. */
  serverRelativeUrl: string;
}
