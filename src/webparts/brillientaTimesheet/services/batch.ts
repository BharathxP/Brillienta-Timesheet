// SharePoint REST $batch encoding (multipart/mixed). Pure functions, unit tested (tests/platform.test.cjs).

export type SPItem = Record<string, unknown>;

const NOMETA = 'application/json;odata=nometadata';

export interface IBatchWrite {
  method: 'POST' | 'MERGE' | 'DELETE';
  /** Relative API path (e.g. web/lists/getbytitle('X')/items(5)). */
  path: string;
  body?: unknown;
}

export interface IBatchResult {
  status: number;
  /** Parsed JSON response of a successful create. */
  data?: SPItem;
  error?: string;
}

export const BATCH_SIZE = 100;

export function randomId(): string {
  return Math.random().toString(36).substring(2) + Date.now().toString(36);
}

/** multipart/mixed body for SharePoint's OData $batch endpoint, one change set per write. */
export function buildBatchBody(writes: Array<IBatchWrite & { url: string }>, boundary: string): string {
  const lines: string[] = [];
  writes.forEach((w, i) => {
    const cs = `changeset_${boundary}_${i}`;
    // SharePoint's OData v4 batch reader rejects change set operations without a Content-ID.
    lines.push(
      `--${boundary}`,
      `Content-Type: multipart/mixed; boundary=${cs}`,
      '',
      `--${cs}`,
      'Content-Type: application/http',
      'Content-Transfer-Encoding: binary',
      `Content-ID: ${i + 1}`,
      ''
    );
    // MERGE/DELETE go as POST with X-HTTP-Method, which SharePoint accepts inside a batch as well.
    lines.push(`POST ${w.url} HTTP/1.1`);
    lines.push(`Accept: ${NOMETA}`);
    lines.push(`Content-Type: ${NOMETA}`);
    if (w.method !== 'POST') {
      lines.push(`X-HTTP-Method: ${w.method}`, 'IF-MATCH: *');
    }
    lines.push('');
    if (w.body !== undefined) lines.push(JSON.stringify(w.body));
    lines.push('', `--${cs}--`);
  });
  lines.push(`--${boundary}--`, '');
  return lines.join('\r\n');
}

/** Status (and JSON body when present) of every operation in a $batch response, in order. */
export function parseBatchResponse(text: string): IBatchResult[] {
  const out: IBatchResult[] = [];
  const re = /HTTP\/1\.1 (\d{3})[^\r\n]*\r?\n([\s\S]*?)(?=\r?\n--(?:changesetresponse|batchresponse)|HTTP\/1\.1 \d{3}|$)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    const status = Number(m[1]);
    const rest = m[2] || '';
    const split = rest.split(/\r?\n\r?\n/);
    const payload = split.slice(1).join('\n\n').trim();
    const result: IBatchResult = { status };
    if (payload) {
      try {
        const json = JSON.parse(payload);
        if (status >= 400) {
          const err = json['odata.error'] || json.error;
          result.error = err && err.message ? (typeof err.message === 'string' ? err.message : err.message.value) : payload;
        } else {
          result.data = json;
        }
      } catch {
        if (status >= 400) result.error = payload.substring(0, 300);
      }
    } else if (status >= 400) {
      result.error = `HTTP ${status}`;
    }
    out.push(result);
  }
  return out;
}

/** Run async tasks with limited parallelism, preserving result order. */
export async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T, index: number) => Promise<R>): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let next = 0;
  const workers: Promise<void>[] = [];
  const run = async (): Promise<void> => {
    while (next < items.length) {
      const i = next++;
      results[i] = await fn(items[i], i);
    }
  };
  for (let w = 0; w < Math.min(limit, items.length); w++) workers.push(run());
  await Promise.all(workers);
  return results;
}

/**
 * Delete each id with `del`, a few at a time; every id is attempted. Failures are collected, not thrown; an item that
 * is already gone (404) counts as removed.
 */
export async function removeEach(ids: number[], del: (id: number) => Promise<void>, limit: number = 5): Promise<{ removed: number; errors: string[] }> {
  const errors: string[] = [];
  let removed = 0;
  await mapLimit(ids, limit, async (id) => {
    try {
      await del(id);
      removed++;
    } catch (e) {
      if ((e as { status?: number }).status === 404) removed++;
      else errors.push((e as Error).message || String(e));
    }
  });
  return { removed, errors };
}
