import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';
import { BATCH_SIZE, IBatchResult, IBatchWrite, buildBatchBody, parseBatchResponse, randomId } from './batch';

export type { IBatchResult, IBatchWrite } from './batch';
export { mapLimit } from './batch';

export type SPItem = Record<string, unknown>;

export class SPError extends Error {
  public readonly status: number;
  public constructor(message: string, status: number) {
    super(message);
    this.status = status;
    // Keep instanceof working after down-levelling to ES5.
    Object.setPrototypeOf(this, SPError.prototype);
  }
}

const NOMETA = 'application/json;odata=nometadata';

/** Escape a value for use inside a single-quoted OData / REST string literal. */
export function odataString(value: string): string {
  return "'" + value.replace(/'/g, "''") + "'";
}

/**
 * Thin wrapper around SPHttpClient for the host site's REST API (_api). No other service is ever called.
 * All requests are relative to the web the web part is on.
 */
export class SPClient {
  private readonly _http: SPHttpClient;
  public readonly webUrl: string;

  public constructor(http: SPHttpClient, webUrl: string) {
    this._http = http;
    this.webUrl = webUrl.replace(/\/$/, '');
  }

  /** A client for another site of the same tenant (Time off Manager 365 on its own site). */
  public forWeb(webUrl: string): SPClient {
    return new SPClient(this._http, webUrl);
  }

  public api(path: string): string {
    return path.indexOf('http') === 0 ? path : `${this.webUrl}/_api/${path.replace(/^\//, '')}`;
  }

  private async _parse<T>(res: SPHttpClientResponse): Promise<T> {
    if (res.status === 204) return undefined as unknown as T;
    const text = await res.text();
    if (!res.ok) {
      let message = `${res.status} ${res.statusText}`;
      try {
        const body = JSON.parse(text);
        const err = body['odata.error'] || body.error;
        if (err && err.message) message = typeof err.message === 'string' ? err.message : err.message.value;
      } catch {
        // not JSON — keep the status line
      }
      throw new SPError(message, res.status);
    }
    if (!text) return undefined as unknown as T;
    return JSON.parse(text) as T;
  }

  public async get<T = SPItem>(path: string): Promise<T> {
    const res = await this._http.get(this.api(path), SPHttpClient.configurations.v1, {
      headers: { Accept: NOMETA }
    });
    return this._parse<T>(res);
  }

  /** GET a collection, following odata.nextLink until `max` rows are read. */
  public async getAll<T = SPItem>(path: string, max: number = 20000): Promise<T[]> {
    let url: string | undefined = this.api(path);
    const rows: T[] = [];
    while (url && rows.length < max) {
      const page: { value: T[]; 'odata.nextLink'?: string } = await this.get(url);
      rows.push(...(page.value || []));
      url = page['odata.nextLink'];
    }
    return rows;
  }

  public async post<T = SPItem>(path: string, body?: unknown, headers?: Record<string, string>): Promise<T> {
    const options: ISPHttpClientOptions = {
      headers: { Accept: NOMETA, 'Content-Type': NOMETA, ...(headers || {}) }
    };
    if (body !== undefined) {
      options.body = body instanceof ArrayBuffer || body instanceof Blob ? body : JSON.stringify(body);
    }
    const res = await this._http.post(this.api(path), SPHttpClient.configurations.v1, options);
    return this._parse<T>(res);
  }

  /** POST using the verbose OData format (needed for a few calls that require __metadata type hints). */
  public async postVerbose<T = SPItem>(path: string, body: unknown, merge: boolean = false): Promise<T> {
    const headers: Record<string, string> = {
      Accept: NOMETA,
      'Content-Type': 'application/json;odata=verbose',
      'odata-version': ''
    };
    if (merge) {
      headers['X-HTTP-Method'] = 'MERGE';
      headers['IF-MATCH'] = '*';
    }
    const res = await this._http.post(this.api(path), SPHttpClient.configurations.v1, {
      headers,
      body: JSON.stringify(body)
    });
    return this._parse<T>(res);
  }

  /** The SharePoint server's clock (Date header), so licence and trial checks do not depend on the PC clock. */
  public async serverTime(): Promise<Date> {
    try {
      const res = await this._http.get(this.api('web?$select=Id'), SPHttpClient.configurations.v1);
      const d = Date.parse(res.headers.get('Date') || '');
      return isNaN(d) ? new Date() : new Date(d);
    } catch {
      return new Date();
    }
  }

  /** Raw bytes of a file in this site (e.g. a list item attachment), by server-relative URL. */
  public async getFile(serverRelativeUrl: string): Promise<ArrayBuffer> {
    const url = this.api(`web/GetFileByServerRelativePath(decodedurl=@a)/$value?@a=${encodeURIComponent(odataString(serverRelativeUrl))}`);
    const res = await this._http.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) throw new SPError(`Could not read ${serverRelativeUrl} (${res.status})`, res.status);
    return res.arrayBuffer();
  }

  public async merge(path: string, body: unknown): Promise<void> {
    await this.post(path, body, { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
  }

  public async del(path: string): Promise<void> {
    await this.post(path, undefined, { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' });
  }

  public list(title: string): string {
    return `web/lists/getbytitle(${odataString(title)})`;
  }

  public async ensureUser(logonName: string): Promise<number> {
    const u = await this.post<{ Id: number }>('web/ensureuser', { logonName });
    return u.Id;
  }

  /**
   * Send many writes in one `$batch` request (at most BATCH_SIZE per call; larger inputs are split). Each write gets its
   * own change set, so one failure does not stop the others. Resolves one result per write, in order.
   */
  public async batch(writes: IBatchWrite[]): Promise<IBatchResult[]> {
    const results: IBatchResult[] = [];
    for (let i = 0; i < writes.length; i += BATCH_SIZE) {
      const chunk = writes.slice(i, i + BATCH_SIZE);
      const boundary = 'batch_' + randomId();
      const body = buildBatchBody(chunk.map((w) => ({ ...w, url: this.api(w.path) })), boundary);
      const res = await this._http.post(`${this.webUrl}/_api/$batch`, SPHttpClient.configurations.v1, {
        headers: { 'Content-Type': `multipart/mixed; boundary=${boundary}`, Accept: 'multipart/mixed' },
        body
      });
      const text = await res.text();
      if (!res.ok) throw new SPError(`Batch failed: ${res.status} ${res.statusText}`, res.status);
      const parsed = parseBatchResponse(text);
      chunk.forEach((_, j) => {
        results.push(parsed[j] || { status: 0, error: 'No response for this change' });
      });
    }
    return results;
  }
}
