/** Characters SharePoint / OneDrive reject in file and folder names. */
// eslint-disable-next-line no-control-regex -- control characters are illegal in SharePoint names too
const ILLEGAL_CHARS = /["*:<>?/\\|\u0000-\u001f]/g;
/** Names SharePoint reserves (case-insensitive), with or without an extension. */
const RESERVED_NAMES = /^(con|prn|aux|nul|com\d|lpt\d|\.lock|desktop\.ini)$/i;
/** Patterns SharePoint rejects anywhere in a name (`_vti_`) or as its start (`~$`, Office lock files). */
const RESERVED_PARTS = /_vti_|^~\$/i;
/** SharePoint's limit for one file name, extension included. */
export const MAX_FILE_NAME_LENGTH = 250;

/** Base name (no extension) cleaned for SharePoint: illegal characters → "-", spaces/dots trimmed. Pure. */
export function cleanBaseName(baseName: string): string {
  return baseName
    .replace(ILLEGAL_CHARS, "-")
    .replace(/\s+/g, " ")
    .replace(/^[\s.]+|[\s.]+$/g, "");
}

/** Why a base name cannot be saved with this extension, or undefined when it can. Pure. */
export function fileNameProblem(baseName: string, extension: string): string | undefined {
  const cleaned = cleanBaseName(baseName);
  if (!cleaned) return "Enter a file name.";
  if (RESERVED_NAMES.test(cleaned) || RESERVED_PARTS.test(cleaned))
    return `"${cleaned}" is a reserved name. Choose another.`;
  if (cleaned.length + extension.length > MAX_FILE_NAME_LENGTH) return "The file name is too long.";
  return undefined;
}

/** Splits "Letter.v2.docx" into { baseName: "Letter.v2", extension: ".docx" }. Pure. */
export function splitFileName(fileName: string): { baseName: string; extension: string } {
  const dot = fileName.lastIndexOf(".");
  return dot > 0
    ? { baseName: fileName.slice(0, dot), extension: fileName.slice(dot) }
    : { baseName: fileName, extension: "" };
}
