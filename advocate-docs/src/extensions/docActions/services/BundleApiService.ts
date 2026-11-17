import { AadHttpClient, type AadHttpClientFactory, type HttpClientResponse } from '@microsoft/sp-http';

/** URL names of the matter site's lists and libraries the build reads and writes. */
export interface IMatterLists {
  documents: string;
  bundles: string;
  registry: string;
  items: string;
  manifests: string;
  /** Hidden library for temporary copies the Function converts to PDF (email attachments). */
  conversion: string;
}

export interface IBuildProgress {
  stage: string;
  done: number;
  total: number;
}

export interface IBuildResult {
  pdfUrl: string;
  version: number;
  pageCount: number;
  documents: number;
  placeholders: number;
  /** Email attachments added as their own index entries (absent from older Function versions). */
  attachments?: number;
}

/** Durable Functions runtime status names. */
export type BuildRuntimeStatus = 'Pending' | 'Running' | 'Completed' | 'Failed' | 'Terminated' | 'Suspended';

export interface IBuildStatus {
  status: BuildRuntimeStatus;
  progress: IBuildProgress | undefined;
  result: IBuildResult | undefined;
  error: string | undefined;
}

/** Calls the AdvocateDocs Function API as the signed-in user (Entra token for the API app). */
export class BundleApiService {
  private _client: AadHttpClient | undefined;

  constructor(
    private readonly _factory: AadHttpClientFactory,
    /** Application ID URI of the API app registration, e.g. api://{clientId}. */
    private readonly _apiResource: string,
    private readonly _baseUrl: string,
    private readonly _webAbsoluteUrl: string,
    private readonly _lists: IMatterLists
  ) {}

  /** Queues a build of the saved bundle; returns the build id to poll. */
  public async startBuild(bundleId: string): Promise<string> {
    const client: AadHttpClient = await this._getClient();
    const response: HttpClientResponse = await client.post(`${this._baseUrl}/api/bundles/build`, AadHttpClient.configurations.v1, {
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ webUrl: this._webAbsoluteUrl, bundleId, lists: this._lists })
    });
    const body: { instanceId?: string; error?: string } = await response.json();
    if (!response.ok || !body.instanceId) {
      throw new Error(`Couldn't start the build (HTTP ${response.status}): ${body.error ?? response.statusText}`);
    }
    return body.instanceId;
  }

  public async getStatus(buildId: string): Promise<IBuildStatus> {
    const client: AadHttpClient = await this._getClient();
    const response: HttpClientResponse = await client.get(`${this._baseUrl}/api/bundles/build/${encodeURIComponent(buildId)}`,
      AadHttpClient.configurations.v1);
    const body: IBuildStatus & { error?: string } = await response.json();
    if (!response.ok) {
      throw new Error(`Couldn't read build status (HTTP ${response.status}): ${body.error ?? response.statusText}`);
    }
    return body;
  }

  /** The documents (any type, email attachments included) as one A4 PDF, ready to print. */
  public async print(itemIds: number[]): Promise<Blob> {
    const client: AadHttpClient = await this._getClient();
    const response: HttpClientResponse = await client.post(`${this._baseUrl}/api/print`, AadHttpClient.configurations.v1, {
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ webUrl: this._webAbsoluteUrl, documentsList: this._lists.documents, conversionList: this._lists.conversion, itemIds })
    });
    if (!response.ok) {
      const body: { error?: string } = await response.json().catch(() => ({}));
      throw new Error(`Couldn't prepare the print (HTTP ${response.status}): ${body.error ?? response.statusText}`);
    }
    return response.blob();
  }

  private async _getClient(): Promise<AadHttpClient> {
    this._client = this._client ?? await this._factory.getClient(this._apiResource);
    return this._client;
  }
}
