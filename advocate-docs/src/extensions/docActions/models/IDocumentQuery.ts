import type { ISelectedDocument } from './ISelectedDocument';

/** A folder in the documents library, for the folder browser. */
export interface ILibraryFolder {
  name: string;
  serverRelativeUrl: string;
  itemCount: number;
}

/** Metadata filter; empty fields are ignored. DocType and DocDate are indexed on the library. */
export interface IDocumentFilter {
  docType: string;
  dateFrom: string;
  dateTo: string;
  fileTypes: string[];
}

/** Outcome of any "add documents" query. */
export interface IDocumentQueryResult {
  documents: ISelectedDocument[];
  /** True when the result hit the safety cap and more documents exist. */
  truncated: boolean;
}
