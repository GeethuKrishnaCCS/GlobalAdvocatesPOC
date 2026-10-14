import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';
import type { ISelectedDocument } from '../models/ISelectedDocument';
import type { IDocumentFilter, IDocumentQueryResult, ILibraryFolder } from '../models/IDocumentQuery';
import { DocumentField, toDocument } from '../utils/documentMapper';

/** Rows per RenderListDataAsStream / search page. */
const PAGE_SIZE: number = 500;
/** Safety cap per query so a broad filter can't load an entire matter into the browser. */
export const MAX_RESULTS: number = 5000;
/** Library system folder that never holds documents. */
const FORMS_FOLDER: string = 'Forms';

const JSON_HEADERS: Record<string, string> = {
  accept: 'application/json;odata=nometadata',
  'content-type': 'application/json;odata=nometadata'
};

interface IListDataResponse {
  Row: Array<Record<string, unknown>>;
  NextHref?: string;
}

// eslint-disable-next-line @rushstack/no-new-null -- SharePoint search returns null cell values
interface ISearchCell { Key: string; Value: string | null }
interface ISearchResponse {
  PrimaryQueryResult?: { RelevantResults?: { TotalRows: number; Table: { Rows: Array<{ Cells: ISearchCell[] }> } } };
}

function xml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/'/g, '&apos;').replace(/"/g, '&quot;');
}

function viewFieldsXml(): string {
  return Object.keys(DocumentField)
    .map(k => `<FieldRef Name='${DocumentField[k as keyof typeof DocumentField]}'/>`)
    .join('');
}

/** Combines CAML conditions with nested <And> (CAML takes exactly two operands per <And>). */
function andAll(conditions: string[]): string {
  return conditions.reduce((acc, c) => (acc ? `<And>${acc}${c}</And>` : c), '');
}

/**
 * Reads documents from the matter's documents library by folder, metadata filter or search.
 * List queries go through RenderListDataAsStream (paged, ordered by ID) so they stay within the
 * list view threshold on large libraries when filtered on indexed columns.
 */
export class DocumentQueryService {
  constructor(
    private readonly _http: SPHttpClient,
    private readonly _webAbsoluteUrl: string,
    private readonly _libraryServerRelativeUrl: string
  ) {}

  public get libraryServerRelativeUrl(): string {
    return this._libraryServerRelativeUrl;
  }

  /** Direct subfolders of a folder (library root when omitted), excluding the system Forms folder. */
  public async getSubfolders(folderServerRelativeUrl: string = this._libraryServerRelativeUrl): Promise<ILibraryFolder[]> {
    const url: string = `${this._webAbsoluteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl=@f)/Folders` +
      `?@f='${encodeURIComponent(folderServerRelativeUrl.replace(/'/g, "''"))}'&$select=Name,ServerRelativeUrl,ItemCount&$orderby=Name`;
    const response: SPHttpClientResponse = await this._http.get(url, SPHttpClient.configurations.v1, { headers: JSON_HEADERS });
    await this._ensureOk(response, 'load folders');
    const body: { value: Array<{ Name: string; ServerRelativeUrl: string; ItemCount: number }> } = await response.json();
    return body.value
      .filter(f => !(folderServerRelativeUrl === this._libraryServerRelativeUrl && f.Name === FORMS_FOLDER))
      .map(f => ({ name: f.Name, serverRelativeUrl: f.ServerRelativeUrl, itemCount: f.ItemCount }));
  }

  /** All documents in a folder; with includeSubfolders, the whole folder tree. */
  public getFolderDocuments(folderServerRelativeUrl: string, includeSubfolders: boolean): Promise<IDocumentQueryResult> {
    const scope: string = includeSubfolders ? " Scope='Recursive'" : '';
    const viewXml: string = `<View${scope}><Query><OrderBy><FieldRef Name='ID'/></OrderBy></Query>` +
      `<ViewFields>${viewFieldsXml()}</ViewFields><RowLimit Paged='TRUE'>${PAGE_SIZE}</RowLimit></View>`;
    return this._renderListData(viewXml, folderServerRelativeUrl);
  }

  /** Documents anywhere in the library matching the metadata filter. */
  public getFilteredDocuments(filter: IDocumentFilter): Promise<IDocumentQueryResult> {
    // Indexed columns first: on large libraries the first condition must narrow below the view threshold.
    const conditions: string[] = [];
    if (filter.docType.trim()) {
      conditions.push(`<Eq><FieldRef Name='DocType'/><Value Type='Text'>${xml(filter.docType.trim())}</Value></Eq>`);
    }
    if (filter.dateFrom) {
      conditions.push(`<Geq><FieldRef Name='DocDate'/><Value Type='DateTime' IncludeTimeValue='FALSE'>${xml(filter.dateFrom)}</Value></Geq>`);
    }
    if (filter.dateTo) {
      conditions.push(`<Leq><FieldRef Name='DocDate'/><Value Type='DateTime' IncludeTimeValue='FALSE'>${xml(filter.dateTo)}</Value></Leq>`);
    }
    if (filter.fileTypes.length) {
      const values: string = filter.fileTypes.map(t => `<Value Type='Text'>${xml(t)}</Value>`).join('');
      conditions.push(`<In><FieldRef Name='${DocumentField.FileType}'/><Values>${values}</Values></In>`);
    }
    const where: string = conditions.length ? `<Where>${andAll(conditions)}</Where>` : '';
    const viewXml: string = `<View Scope='Recursive'><Query>${where}<OrderBy><FieldRef Name='ID'/></OrderBy></Query>` +
      `<ViewFields>${viewFieldsXml()}</ViewFields><RowLimit Paged='TRUE'>${PAGE_SIZE}</RowLimit></View>`;
    return this._renderListData(viewXml);
  }

  /** Full-text search (name + content) limited to the documents library. Depends on the search index being current. */
  public async searchDocuments(text: string): Promise<IDocumentQueryResult> {
    const libraryAbsoluteUrl: string = new URL(this._libraryServerRelativeUrl, this._webAbsoluteUrl).href;
    const query: string = `${text.trim()} path:"${libraryAbsoluteUrl}" IsDocument:1`.replace(/'/g, "''");
    const documents: ISelectedDocument[] = [];
    let startRow: number = 0;
    let total: number = 0;
    do {
      const url: string = `${this._webAbsoluteUrl}/_api/search/query?querytext='${encodeURIComponent(query)}'` +
        `&selectproperties='ListItemID,UniqueId,Filename,FileType,Path'&rowlimit=${PAGE_SIZE}&startrow=${startRow}&trimduplicates=false`;
      const response: SPHttpClientResponse = await this._http.get(url, SPHttpClient.configurations.v1, { headers: JSON_HEADERS });
      await this._ensureOk(response, 'search documents');
      const body: ISearchResponse = await response.json();
      const results = body.PrimaryQueryResult?.RelevantResults;
      if (!results) {
        break;
      }
      total = results.TotalRows;
      for (const row of results.Table.Rows) {
        const cells: Record<string, string> = {};
        row.Cells.forEach(c => { cells[c.Key] = c.Value ?? ''; });
        const doc: ISelectedDocument | undefined = toDocument(name => ({
          [DocumentField.Id]: cells.ListItemID,
          [DocumentField.UniqueId]: cells.UniqueId,
          [DocumentField.FileRef]: cells.Path ? decodeURI(new URL(cells.Path).pathname) : '',
          [DocumentField.FileLeafRef]: cells.Filename,
          [DocumentField.FileType]: cells.FileType,
          [DocumentField.ObjectType]: '0'
        } as Record<string, string>)[name]);
        if (doc) {
          documents.push(doc);
        }
      }
      startRow += results.Table.Rows.length;
      if (results.Table.Rows.length === 0) {
        break;
      }
    } while (startRow < total && documents.length < MAX_RESULTS);
    return { documents: documents.slice(0, MAX_RESULTS), truncated: total > MAX_RESULTS };
  }

  private async _renderListData(viewXml: string, folderServerRelativeUrl?: string): Promise<IDocumentQueryResult> {
    const base: string = `${this._webAbsoluteUrl}/_api/web/GetList(@l)/RenderListDataAsStream` +
      `?@l='${encodeURIComponent(this._libraryServerRelativeUrl.replace(/'/g, "''"))}'`;
    const parameters: Record<string, unknown> = { RenderOptions: 2, ViewXml: viewXml, DatesInUtc: true };
    if (folderServerRelativeUrl) {
      parameters.FolderServerRelativeUrl = folderServerRelativeUrl;
    }
    const documents: ISelectedDocument[] = [];
    let nextHref: string | undefined = '';
    while (nextHref !== undefined && documents.length < MAX_RESULTS) {
      const url: string = nextHref ? `${base}&${nextHref.replace(/^\?/, '')}` : base;
      const response: SPHttpClientResponse = await this._http.post(url, SPHttpClient.configurations.v1, {
        headers: JSON_HEADERS,
        body: JSON.stringify({ parameters })
      });
      await this._ensureOk(response, 'load documents');
      const body: IListDataResponse = await response.json();
      body.Row.forEach(row => {
        const doc: ISelectedDocument | undefined = toDocument(name => row[name]);
        if (doc) {
          documents.push(doc);
        }
      });
      nextHref = body.NextHref || undefined;
    }
    return { documents: documents.slice(0, MAX_RESULTS), truncated: nextHref !== undefined || documents.length > MAX_RESULTS };
  }

  private async _ensureOk(response: SPHttpClientResponse, action: string): Promise<void> {
    if (response.ok) {
      return;
    }
    let detail: string = '';
    try {
      const body: { 'odata.error'?: { message?: { value?: string } }; error?: { message?: string } } = await response.json();
      detail = body['odata.error']?.message?.value ?? body.error?.message ?? '';
    } catch {
      detail = response.statusText;
    }
    throw new Error(`Couldn't ${action} (HTTP ${response.status})${detail ? `: ${detail}` : ''}`);
  }
}
