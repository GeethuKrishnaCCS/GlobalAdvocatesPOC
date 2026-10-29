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
