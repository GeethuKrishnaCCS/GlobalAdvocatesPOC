import type { ISelectedDocument } from './ISelectedDocument';

export type BundleStatus = 'Draft' | 'Built';

/** One saved bundle (row in the bundle registry list). */
export interface IBundleHeader {
  /** Registry list item id; undefined until first saved. */
  itemId?: number;
  bundleId: string;
  title: string;
  status: BundleStatus;
  memberCount: number;
  latestPdfUrl?: string;
  modified?: string;
}

/** A saved bundle with its ordered members. */
export interface IBundle {
  header: IBundleHeader;
  documents: ISelectedDocument[];
}

/** Outcome of appending documents to a saved bundle. */
export interface IAddToBundleResult {
  added: number;
  skipped: number;
}
