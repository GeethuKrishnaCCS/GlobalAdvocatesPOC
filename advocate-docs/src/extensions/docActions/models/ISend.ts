/** A library document as a Graph drive item. */
export interface IDriveItemRef {
  driveId: string;
  itemId: string;
  name: string;
  size: number;
}

/** One entry of a document's version history. */
export interface IDocumentVersion {
  id: string;
  label: string;
  modified: string;
  modifiedBy: string;
  size: number;
  isCurrent: boolean;
}

/** File content ready to attach. */
export interface IFilePayload {
  name: string;
  bytes: ArrayBuffer;
  versionLabel: string | undefined;
}

/** What the Send commands ask for. */
export interface ISendRequest {
  /** Previous-version commands pick a version of a single document; current-version commands send 1+ documents. */
  mode: 'current' | 'version';
  asPdf: boolean;
}

export type SendStepStatus = 'waiting' | 'working' | 'attached' | 'linked' | 'failed';

/** Progress of one document in the draft. */
export interface ISendStep {
  name: string;
  status: SendStepStatus;
  detail?: string;
}

/** Created draft. */
export interface IDraftResult {
  webLink: string;
  attached: number;
  linked: number;
  failed: number;
}
