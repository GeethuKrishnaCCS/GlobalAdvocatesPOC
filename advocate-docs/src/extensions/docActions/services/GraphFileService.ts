import type { MSGraphClientV3 } from '@microsoft/sp-http-msgraph';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import type { ISelectedDocument } from '../models/ISelectedDocument';
import type { IDocumentVersion, IDriveItemRef, IFilePayload } from '../models/ISend';

/** Types already PDF — "as PDF" sends them unchanged. */
const PDF_TYPE: string = 'pdf';

interface IGraphVersion {
  id: string;
  lastModifiedDateTime: string;
  size: number;
  lastModifiedBy?: { user?: { displayName?: string } };
}

function toPdfName(name: string): string {
  const dot: number = name.lastIndexOf('.');
  return `${dot > 0 ? name.substring(0, dot) : name}.pdf`;
}

/**
 * Document access through Microsoft Graph (delegated, as the signed-in user): resolve a library item to
 * its drive item, list versions, download any version natively or converted to PDF.
 */
export class GraphFileService {
  constructor(
    private readonly _graph: MSGraphClientV3,
    /** Graph site id: "{hostname},{siteCollectionId},{webId}". */
    private readonly _siteId: string,
    private readonly _listId: string,
    /**
     * Id of the hidden library that holds short-lived copies of old versions being converted to PDF
     * (resolved only when a conversion is needed; each copy is deleted after use).
     */
    private readonly _getConversionListId: () => Promise<string>
  ) {}

  public async resolve(doc: ISelectedDocument): Promise<IDriveItemRef> {
    const item: { id: string; name: string; size: number; parentReference: { driveId: string } } = await this._graph
      .api(`/sites/${this._siteId}/lists/${this._listId}/items/${doc.id}/driveItem`)
      .select('id,name,size,parentReference')
      .get();
    return { driveId: item.parentReference.driveId, itemId: item.id, name: item.name, size: item.size };
  }

  /** All versions, newest first; the first entry is the current version. */
  public async listVersions(ref: IDriveItemRef): Promise<IDocumentVersion[]> {
    const body: { value: IGraphVersion[] } = await this._graph.api(`/drives/${ref.driveId}/items/${ref.itemId}/versions`).get();
    return body.value.map((v, i) => ({
      id: v.id,
      label: v.id,
      modified: v.lastModifiedDateTime,
      modifiedBy: v.lastModifiedBy?.user?.displayName ?? '',
      size: v.size,
      isCurrent: i === 0
    }));
  }

  /** Current version, as stored or as PDF. */
  public async getCurrent(ref: IDriveItemRef, asPdf: boolean): Promise<IFilePayload> {
    const convert: boolean = asPdf && !this._isPdf(ref.name);
    const path: string = `/drives/${ref.driveId}/items/${ref.itemId}/content${convert ? '?format=pdf' : ''}`;
    const bytes: ArrayBuffer = await this._download(path);
    return { name: convert ? toPdfName(ref.name) : ref.name, bytes, versionLabel: undefined };
  }

  /**
   * A specific (older) version. Graph converts only current files, so for PDF the version is copied to a
   * temporary file, converted, and the copy removed (to the recycle bin) whatever the outcome.
   */
  public async getVersion(ref: IDriveItemRef, version: IDocumentVersion, asPdf: boolean): Promise<IFilePayload> {
    if (version.isCurrent) {
      return { ...(await this.getCurrent(ref, asPdf)), versionLabel: version.label };
    }
    const bytes: ArrayBuffer = await this._download(`/drives/${ref.driveId}/items/${ref.itemId}/versions/${version.id}/content`);
    const versionedName: string = this._withVersion(ref.name, version.label);
    if (!asPdf || this._isPdf(ref.name)) {
      return { name: versionedName, bytes, versionLabel: version.label };
    }
    const temp: IDriveItemRef = await this._uploadTemp(`${Date.now()}_${ref.name}`, bytes);
    try {
      const pdf: ArrayBuffer = await this._download(`/drives/${temp.driveId}/items/${temp.itemId}/content?format=pdf`);
      return { name: toPdfName(versionedName), bytes: pdf, versionLabel: version.label };
    } finally {
      await this._graph.api(`/drives/${temp.driveId}/items/${temp.itemId}`).delete().catch(() => undefined);
    }
  }

  private _isPdf(name: string): boolean {
    return (name.split('.').pop() ?? '').toLowerCase() === PDF_TYPE;
  }

  /** "Contract.docx" + "2.0" → "Contract (v2.0).docx", so recipients can tell versions apart. */
  private _withVersion(name: string, label: string): string {
    const dot: number = name.lastIndexOf('.');
    return dot > 0 ? `${name.substring(0, dot)} (v${label})${name.substring(dot)}` : `${name} (v${label})`;
  }

  private async _download(path: string): Promise<ArrayBuffer> {
    return this._graph.api(path).responseType(ResponseType.ARRAYBUFFER).get();
  }

  private async _uploadTemp(fileName: string, bytes: ArrayBuffer): Promise<IDriveItemRef> {
    const listId: string = await this._getConversionListId();
    const path: string = `/sites/${this._siteId}/lists/${listId}/drive/root:/${encodeURIComponent(fileName)}:/content`;
    // A Blob, not an ArrayBuffer: the Graph client's ArrayBuffer path relies on Node's Buffer, absent in browsers.
    const body: Blob = new Blob([bytes], { type: 'application/octet-stream' });
    const item: { id: string; name: string; size: number; parentReference: { driveId: string } } = await this._graph
      .api(path).header('Content-Type', 'application/octet-stream').put(body);
    return { driveId: item.parentReference.driveId, itemId: item.id, name: item.name, size: item.size };
  }
}
