import type { RowAccessor } from '@microsoft/sp-listview-extensibility';
import type { ISelectedDocument } from '../models/ISelectedDocument';

/** Internal names of the fields the mapper reads (list view rows and RenderListDataAsStream rows). */
export const DocumentField = {
  Id: 'ID',
  UniqueId: 'UniqueId',
  FileRef: 'FileRef',
  FileLeafRef: 'FileLeafRef',
  FileType: 'File_x0020_Type',
  ObjectType: 'FSObjType'
} as const;

const FOLDER_OBJECT_TYPE: string = '1';

type ValueReader = (internalName: string) => unknown;

/** Maps any field reader to a document; undefined for folders or rows without a file id. */
export function toDocument(read: ValueReader): ISelectedDocument | undefined {
  if (String(read(DocumentField.ObjectType)) === FOLDER_OBJECT_TYPE) {
    return undefined;
  }
  const uniqueId: string = String(read(DocumentField.UniqueId) ?? '').replace(/[{}]/g, '').toLowerCase();
  const serverRelativeUrl: string = String(read(DocumentField.FileRef) ?? '');
  if (!uniqueId || !serverRelativeUrl) {
    return undefined;
  }
  const name: string = String(read(DocumentField.FileLeafRef) || serverRelativeUrl.split('/').pop());
  const fileType: string = String(read(DocumentField.FileType) || name.split('.').pop() || '').toLowerCase();
  return { id: Number(read(DocumentField.Id)), uniqueId, name, fileType, serverRelativeUrl };
}

/** Maps the list view selection, dropping folders and unmappable rows. */
export function toSelectedDocuments(rows: ReadonlyArray<RowAccessor>): ISelectedDocument[] {
  return rows
    .map(row => toDocument(name => row.getValueByName(name)))
    .filter((d): d is ISelectedDocument => d !== undefined);
}
