import { Guid } from '@microsoft/sp-core-library';
import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';
import type { IAddToBundleResult, IBundle, IBundleHeader } from '../models/IBundle';
import type { ISelectedDocument } from '../models/ISelectedDocument';

/** Parallel REST writes when syncing members (SharePoint throttles bursts). */
const WRITE_CONCURRENCY: number = 6;
const PAGE_SIZE: number = 2000;

const JSON_HEADERS: Record<string, string> = {
  accept: 'application/json;odata=nometadata',
  'content-type': 'application/json;odata=nometadata'
};

interface IMemberRow {
  Id: number;
  SourceUniqueId: string;
  SortOrder: number;
  IndexTitle: string;
}

/* eslint-disable @rushstack/no-new-null -- SharePoint REST returns null for empty fields */
interface IRegistryRow {
  Id: number;
  Title: string;
  BundleId: string;
  BundleStatus: 'Draft' | 'Built' | null;
  MemberCount: number | null;
  LatestPdfUrl: string | null;
  Modified: string;
}
/* eslint-enable @rushstack/no-new-null */

function odataString(value: string): string {
  return value.replace(/'/g, "''");
}

/** Runs async work over items with a fixed number in flight. */
async function forEachLimited<T>(items: ReadonlyArray<T>, limit: number, work: (item: T) => Promise<void>): Promise<void> {
  let next: number = 0;
  const workers: Promise<void>[] = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) {
      const item: T = items[next++];
      await work(item);
    }
  });
  await Promise.all(workers);
}

function toDocumentFromMember(row: IMemberRow): ISelectedDocument {
  const name: string = row.IndexTitle || row.SourceUniqueId;
  return {
    id: 0,
    uniqueId: row.SourceUniqueId.toLowerCase(),
    name,
    fileType: (name.split('.').pop() ?? '').toLowerCase(),
    serverRelativeUrl: ''
  };
}

/**
 * Saved bundles on the matter site: a header row per bundle in the registry list and one row per
 * member (ordered by SortOrder) in the bundle items list, linked by BundleId.
 */
export class BundleService {
  constructor(
    private readonly _http: SPHttpClient,
    private readonly _webAbsoluteUrl: string,
    private readonly _registryListUrl: string,
    private readonly _itemsListUrl: string
  ) {}

  /** All saved bundles, most recently changed first. */
  public async listBundles(): Promise<IBundleHeader[]> {
    const rows: IRegistryRow[] = await this._getAll<IRegistryRow>(this._registryListUrl,
      '$select=Id,Title,BundleId,BundleStatus,MemberCount,LatestPdfUrl,Modified&$orderby=Modified desc');
    return rows.map(r => ({
      itemId: r.Id,
      bundleId: r.BundleId,
      title: r.Title,
      status: r.BundleStatus ?? 'Draft',
      memberCount: r.MemberCount ?? 0,
      latestPdfUrl: r.LatestPdfUrl ?? undefined,
      modified: r.Modified
    }));
  }

  /** A bundle with its members in order. */
  public async getBundle(header: IBundleHeader): Promise<IBundle> {
    const rows: IMemberRow[] = await this._getMembers(header.bundleId);
    return { header, documents: rows.map(toDocumentFromMember) };
  }

  /**
   * Creates or updates a bundle so its members match `documents` exactly (order included):
   * new members are added, removed ones recycled, moved ones re-ordered.
   */
  public async saveBundle(header: IBundleHeader, documents: ReadonlyArray<ISelectedDocument>): Promise<IBundleHeader> {
    const saved: IBundleHeader = await this._saveHeader({ ...header, memberCount: documents.length });
    const existing: IMemberRow[] = header.itemId ? await this._getMembers(saved.bundleId) : [];
    const byUniqueId: Map<string, IMemberRow> = new Map(existing.map(r => [r.SourceUniqueId.toLowerCase(), r]));
    const wanted: Set<string> = new Set(documents.map(d => d.uniqueId));

    const toRemove: IMemberRow[] = existing.filter(r => !wanted.has(r.SourceUniqueId.toLowerCase()));
    await forEachLimited(toRemove, WRITE_CONCURRENCY, r => this._recycleItem(this._itemsListUrl, r.Id));

    const positions: Array<{ doc: ISelectedDocument; order: number }> = documents.map((doc, i) => ({ doc, order: i + 1 }));
    await forEachLimited(positions, WRITE_CONCURRENCY, async ({ doc, order }) => {
      const row: IMemberRow | undefined = byUniqueId.get(doc.uniqueId);
      if (!row) {
        await this._createMember(saved.bundleId, doc, order);
      } else if (row.SortOrder !== order) {
        await this._updateItem(this._itemsListUrl, row.Id, { SortOrder: order });
      }
    });
    return saved;
  }

  /** Appends documents not already in the bundle, after its current last member. */
  public async addToBundle(header: IBundleHeader, documents: ReadonlyArray<ISelectedDocument>): Promise<IAddToBundleResult> {
    const existing: IMemberRow[] = await this._getMembers(header.bundleId);
    const present: Set<string> = new Set(existing.map(r => r.SourceUniqueId.toLowerCase()));
    const fresh: ISelectedDocument[] = documents.filter(d => !present.has(d.uniqueId));
    const start: number = existing.reduce((max, r) => Math.max(max, r.SortOrder ?? 0), 0);
    await forEachLimited(fresh.map((doc, i) => ({ doc, order: start + i + 1 })), WRITE_CONCURRENCY,
      ({ doc, order }) => this._createMember(header.bundleId, doc, order));
    await this._saveHeader({ ...header, memberCount: existing.length + fresh.length });
    return { added: fresh.length, skipped: documents.length - fresh.length };
  }

  private async _saveHeader(header: IBundleHeader): Promise<IBundleHeader> {
    const fields: Record<string, unknown> = { Title: header.title, MemberCount: header.memberCount };
    if (header.itemId) {
      await this._updateItem(this._registryListUrl, header.itemId, fields);
      return header;
    }
    const bundleId: string = header.bundleId || Guid.newGuid().toString();
    const created: { Id: number } = await this._createItem(this._registryListUrl, { ...fields, BundleId: bundleId, BundleStatus: 'Draft' });
    return { ...header, itemId: created.Id, bundleId, status: 'Draft' };
  }

  private _getMembers(bundleId: string): Promise<IMemberRow[]> {
    return this._getAll<IMemberRow>(this._itemsListUrl,
      `$select=Id,SourceUniqueId,SortOrder,IndexTitle&$filter=BundleId eq '${odataString(bundleId)}'&$orderby=SortOrder`);
  }

  private async _createMember(bundleId: string, doc: ISelectedDocument, order: number): Promise<void> {
    await this._createItem(this._itemsListUrl, {
      Title: doc.name,
      BundleId: bundleId,
      SourceUniqueId: doc.uniqueId,
      SortOrder: order,
      IndexTitle: doc.name
    });
  }

  private _listApi(listUrl: string, path: string = ''): string {
    return `${this._webAbsoluteUrl}/_api/web/GetList(@l)/items${path}?@l='${encodeURIComponent(odataString(listUrl))}'`;
  }

  private async _getAll<T>(listUrl: string, query: string): Promise<T[]> {
    const rows: T[] = [];
    let url: string | undefined = `${this._listApi(listUrl)}&${query}&$top=${PAGE_SIZE}`;
    while (url) {
      const response: SPHttpClientResponse = await this._http.get(url, SPHttpClient.configurations.v1, { headers: JSON_HEADERS });
      await this._ensureOk(response, 'read bundles');
      const body: { value: T[]; 'odata.nextLink'?: string } = await response.json();
      rows.push(...body.value);
      url = body['odata.nextLink'];
    }
    return rows;
  }

  private async _createItem(listUrl: string, fields: Record<string, unknown>): Promise<{ Id: number }> {
    const response: SPHttpClientResponse = await this._http.post(this._listApi(listUrl), SPHttpClient.configurations.v1,
      { headers: JSON_HEADERS, body: JSON.stringify(fields) });
    await this._ensureOk(response, 'save bundle');
    return response.json();
  }

  private async _updateItem(listUrl: string, id: number, fields: Record<string, unknown>): Promise<void> {
    const response: SPHttpClientResponse = await this._http.post(this._listApi(listUrl, `(${id})`), SPHttpClient.configurations.v1, {
      headers: { ...JSON_HEADERS, 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
      body: JSON.stringify(fields)
    });
    await this._ensureOk(response, 'update bundle');
  }

  /** Members are recycled (recoverable from the site recycle bin), never hard-deleted. */
  private async _recycleItem(listUrl: string, id: number): Promise<void> {
    const response: SPHttpClientResponse = await this._http.post(this._listApi(listUrl, `(${id})/recycle()`),
      SPHttpClient.configurations.v1, { headers: JSON_HEADERS });
    await this._ensureOk(response, 'remove bundle member');
  }

  private async _ensureOk(response: SPHttpClientResponse, action: string): Promise<void> {
    if (response.ok) {
      return;
    }
    let detail: string = response.statusText;
    try {
      const body: { 'odata.error'?: { message?: { value?: string } } } = await response.json();
      detail = body['odata.error']?.message?.value ?? detail;
    } catch {
      // Non-JSON error body; keep the status text.
    }
    throw new Error(`Couldn't ${action} (HTTP ${response.status}): ${detail}`);
  }
}
