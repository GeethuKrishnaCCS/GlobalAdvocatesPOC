/** Entra app the add-in signs in with (ADVTest-DocsApi, dev tenant ccsdev01 — see repo .project.json). */
export const ENTRA_CLIENT_ID = "daf35d1e-0756-4216-a61d-827d5eca434f";
export const ENTRA_AUTHORITY =
  "https://login.microsoftonline.com/20e5d750-9ba4-44bf-a219-f361c9bfe523";

/** Delegated Graph scopes (admin-consented on ADVTest-DocsApi). */
export const GRAPH_SCOPES = [
  "https://graph.microsoft.com/Files.ReadWrite.All",
  "https://graph.microsoft.com/Sites.Read.All",
  "https://graph.microsoft.com/Mail.ReadWrite",
];

/** Task pane views, selected by the `view` query parameter of the manifest page URL. */
export enum PaneView {
  SaveToMatter = "save",
  SendAsAttachment = "send",
}

/** Matter sites are the sites whose URL alias starts with this (test marker on dev; firm's matter site prefix in prod). */
export const MATTER_SITE_SEARCH = "ADVTestMatter";
/** Matter document library, by list URL name (language-independent, unlike the display name). */
export const MATTER_LIBRARY_NAME = "ADVTestDocuments";

/** Document type choices (values already used in the matter library's DocType column). */
export const DOCUMENT_TYPES = [
  "Affidavit",
  "Contract",
  "Correspondence",
  "Evidence",
  "Order",
  "Pleading",
];
/** AdvDirection column choices. */
export const DIRECTIONS = ["Inward", "Outward"];

/** Matter folder that filed correspondence (sent emails) goes to, relative to the matter library root. */
export const FILING_FOLDER_PATH = "Correspondence/Matter";
/**
 * Draft extended property naming the folder the sent email should be filed to ("driveId|itemId"); Outlook
 * "File to Matter" reads it to pre-fill matter and folder.
 */
export const FILE_TO_FOLDER_PROPERTY_ID =
  "String {8997e24e-72ed-4959-b04f-1912c4f5bd9d} Name AdvFileToFolder";
/** Outlook category that tags an email with its matter ("Matter: ADVTestMatter01"). */
export const MATTER_CATEGORY_PREFIX = "Matter: ";
