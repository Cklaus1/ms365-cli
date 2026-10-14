import { AuthManager } from './auth.js';
import logger from './logger.js';
import { CliError, ExitCode } from './errors.js';
import { recordAudit } from './audit.js';

const MAX_RETRIES = 3;
const MAX_PAGES = 100;
const MAX_REDIRECTS = 5;
const REQUEST_TIMEOUT_MS = 30_000;

/**
 * How the caller wants the response body:
 * - auto (default): JSON → parsed; text/html/xml → { rawResponse }; anything else → base64 contentBytes
 * - json / text / binary: force that decoding
 * - stream: body is not read; GraphResponse.stream carries it (for large downloads)
 */
export type ResponseType = 'auto' | 'json' | 'text' | 'binary' | 'stream';

export interface GraphRequestOptions {
  method?: string;
  headers?: Record<string, string>;
  body?: string | Uint8Array;
  queryParams?: Record<string, string>;
  responseType?: ResponseType;
  /** Timeout until response headers (default 30s). */
  timeoutMs?: number;
}

export interface GraphResponse {
  data: unknown;
  status: number;
  nextLink?: string;
  deltaLink?: string;
  count?: number;
  headers?: Headers;
  /** Set when responseType is 'stream'. */
  stream?: ReadableStream<Uint8Array> | null;
}

export interface BatchRequest {
  method: string;
  /** Path relative to the API version, e.g. "/me/messages/123". */
  url: string;
  body?: unknown;
  headers?: Record<string, string>;
}

export interface BatchResult {
  status: number;
  body: unknown;
  headers?: Record<string, string>;
}

const BATCH_SIZE = 20;

/** Methods that are safe to retry after a timeout or 5xx (the server may have processed them). */
const IDEMPOTENT_METHODS = new Set(['GET', 'HEAD', 'OPTIONS', 'PUT', 'DELETE', 'PATCH']);

/** CliError codes raised before anything was sent to Graph (not audited). */
const NOT_SENT_CODES = new Set<string | undefined>(['read_only', 'untrusted_url', 'auth_required']);

/** Hosts that may receive our bearer token when following redirects. */
export function isTrustedHost(url: string): boolean {
  try {
    const host = new URL(url).hostname.toLowerCase();
    return host === 'graph.microsoft.com'
      || host.endsWith('.sharepoint.com')
      || host.endsWith('.svc.ms');
  } catch {
    return false;
  }
}

/** Connection errors raised before any request bytes reached the server. */
function isPreSendError(err: unknown): boolean {
  const code = (err as { cause?: { code?: string } })?.cause?.code;
  return code === 'ECONNREFUSED' || code === 'ENOTFOUND' || code === 'EAI_AGAIN';
}

export class GraphClient {
  private auth: AuthManager;
  private dryRun: boolean;
  private readOnly: boolean;
  private apiVersion: string;
  private baseUrl: string;

  constructor(auth: AuthManager, dryRun = false, apiVersion = 'v1.0', readOnly = false) {
    this.auth = auth;
    this.dryRun = dryRun;
    this.readOnly = readOnly;
    this.apiVersion = apiVersion;
    this.baseUrl = `https://graph.microsoft.com/${this.apiVersion}`;
  }

  buildUrl(endpoint: string, queryParams?: Record<string, string>): string {
    let url = /^https:\/\//.test(endpoint) ? endpoint : `${this.baseUrl}${endpoint}`;

    if (queryParams && Object.keys(queryParams).length > 0) {
      const qs = Object.entries(queryParams)
        .map(([k, v]) => {
          const key = k.startsWith('$') ? '$' + encodeURIComponent(k.slice(1)) : encodeURIComponent(k);
          return `${key}=${encodeURIComponent(v)}`;
        })
        .join('&');
      url += (url.includes('?') ? '&' : '?') + qs;
    }

    return url;
  }

  /**
   * Send a Graph request. `endpoint` is a path under the API base, or an absolute
   * https://graph.microsoft.com URL (pagination nextLinks).
   */
  async request(endpoint: string, options: GraphRequestOptions = {}): Promise<GraphResponse> {
    const method = (options.method || 'GET').toUpperCase();
    if (method === 'GET') return this.send(endpoint, options);
    return this.audited(method, endpoint, () => this.send(endpoint, options));
  }

  // --- audit hook (T37): one line per non-GET call that reached Graph ---
  /** Run `fn` and record its outcome. Dry runs (status 0) and calls refused before sending are not logged. */
  private async audited(method: string, endpoint: string, fn: () => Promise<GraphResponse>): Promise<GraphResponse> {
    let res: GraphResponse;
    try {
      res = await fn();
    } catch (err) {
      const e = err as CliError;
      if (!NOT_SENT_CODES.has(e.code)) this.audit(method, endpoint, e.status ?? 0, e.requestId, false);
      throw err;
    }
    if (res.status !== 0) this.audit(method, endpoint, res.status, res.headers?.get('request-id') ?? undefined, true);
    return res;
  }

  private audit(method: string, endpoint: string, status: number, requestId: string | undefined, ok: boolean): void {
    recordAudit({ method, endpoint, status, requestId, ok });
  }
  private auditBatchItem(r: BatchRequest, status: number, requestId: string | undefined): void {
    if (r.method.toUpperCase() !== 'GET') this.audit(r.method.toUpperCase(), r.url, status, requestId, status >= 200 && status < 300);
  }
  // --- end audit hook ---

  /** request() without the audit hook. */
  private async send(endpoint: string, options: GraphRequestOptions = {}): Promise<GraphResponse> {
    if (/^https:\/\//.test(endpoint) && new URL(endpoint).hostname !== 'graph.microsoft.com') {
      throw new CliError(`Refusing to send credentials to non-Graph URL: ${new URL(endpoint).hostname}`, { code: 'untrusted_url' });
    }
    const url = this.buildUrl(endpoint, options.queryParams);
    const method = (options.method || 'GET').toUpperCase();
    const responseType = options.responseType ?? 'auto';

    logger.debug(`${method} ${url}`, options.body ? { bodyLength: options.body.length } : undefined);


    if (this.readOnly && method !== 'GET') {
      throw new CliError(`Read-only mode: ${method} requests are blocked. Remove --read-only to proceed.`, { code: 'read_only', exitCode: ExitCode.PERMISSION });
    }

    if (this.dryRun) {
      console.log(`[dry-run] ${method} ${url}`);
      if (typeof options.body === 'string') {
        console.log(`[dry-run] Body: ${options.body}`);
      } else if (options.body) {
        console.log(`[dry-run] Body: <${options.body.length} bytes>`);
      }
      if (options.headers && Object.keys(options.headers).length > 0) {
        console.log(`[dry-run] Headers: ${JSON.stringify(options.headers)}`);
      }
      return { data: { dryRun: true, method, url }, status: 0 };
    }

    const token = await this.auth.getToken();

    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...options.headers,
    };

    const idempotent = IDEMPOTENT_METHODS.has(method);
    let lastError: Error | null = null;
    let retriedWith401 = false;

    for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
      let response: Response;
      try {
        response = await this.fetchFollowingRedirects(url, method, headers, options.body, options.timeoutMs ?? REQUEST_TIMEOUT_MS);
      } catch (err: unknown) {
        const timedOut = err instanceof Error && err.name === 'AbortError';
        const retryable = idempotent || isPreSendError(err);
        if (timedOut) {
          lastError = new CliError(
            `Request timed out after ${(options.timeoutMs ?? REQUEST_TIMEOUT_MS) / 1000}s: ${method} ${url}` + (idempotent ? '' : ' (not retried; it may have been processed)'),
            { code: 'timeout' },
          );
        } else {
          lastError = err instanceof Error ? err : new Error(String(err));
        }
        if (retryable && attempt < MAX_RETRIES - 1) {
          const waitMs = Math.pow(2, attempt) * 1000;
          logger.warn(`${timedOut ? 'Request timed out' : 'Connection error'}. Retrying in ${waitMs / 1000}s...`);
          await sleep(waitMs);
          continue;
        }
        throw lastError;
      }

      logger.debug(`Response: ${response.status} ${response.statusText}`, { attempt });

      // Rate limited — the request was not processed, so any method may retry
      if (response.status === 429 && attempt < MAX_RETRIES - 1) {
        const waitMs = parseRetryAfter(response.headers.get('Retry-After'), Math.pow(2, attempt) * 1000);
        logger.warn(`Rate limited. Retrying in ${waitMs / 1000}s...`);
        console.error(`Rate limited. Retrying in ${waitMs / 1000}s...`);
        await response.body?.cancel();
        await sleep(waitMs);
        continue;
      }

      // Server error — retry only idempotent methods (a POST may already have been applied)
      if (response.status >= 500 && idempotent && attempt < MAX_RETRIES - 1) {
        const waitMs = Math.pow(2, attempt) * 1000;
        logger.warn(`Server error (${response.status}). Retrying in ${waitMs / 1000}s...`);
        console.error(`Server error (${response.status}). Retrying in ${waitMs / 1000}s...`);
        await response.body?.cancel();
        await sleep(waitMs);
        continue;
      }

      if (!response.ok) {
        const errorText = await response.text();
        let errorMessage: string;
        let graphCode: string | undefined;
        let bodyRequestId: string | undefined;
        try {
          const errorJson = JSON.parse(errorText);
          errorMessage = errorJson.error?.message || errorText;
          graphCode = errorJson.error?.code;
          bodyRequestId = errorJson.error?.innerError?.['request-id'];
        } catch {
          errorMessage = errorText;
        }
        const requestId = response.headers.get('request-id') || bodyRequestId || undefined;
        const details = { status: response.status, graphCode, requestId };

        logger.error(`API error ${response.status}: ${errorMessage}`);

        // Retry once with a fresh token on 401
        if (response.status === 401 && !retriedWith401) {
          retriedWith401 = true;
          try {
            const freshToken = await this.auth.getToken(true);
            headers['Authorization'] = `Bearer ${freshToken}`;
            logger.info('Token refreshed, retrying request...');
            attempt--; // a token refresh is not a failed attempt
            continue;
          } catch {
            throw new CliError(`Authentication failed. Run: ms365 auth login`,
              { ...details, code: 'auth_failed', exitCode: ExitCode.AUTH, fix: 'ms365 auth login' });
          }
        }
        throw graphError(response.status, errorMessage, details);
      }

      return decodeResponse(response, responseType);
    }

    throw lastError || new Error('Request failed after retries');
  }

  /**
   * fetch with a 30s timeout to response headers. GET redirects are followed manually
   * so the bearer token is only re-sent to trusted Microsoft hosts (tempauth retirement).
   */
  private async fetchFollowingRedirects(
    url: string,
    method: string,
    headers: Record<string, string>,
    body: string | Uint8Array | undefined,
    timeoutMs: number,
  ): Promise<Response> {
    let currentUrl = url;
    for (let hop = 0; hop <= MAX_REDIRECTS; hop++) {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), timeoutMs);
      const hopHeaders = { ...headers };
      if (hop > 0 && !isTrustedHost(currentUrl)) delete hopHeaders['Authorization'];
      let response: Response;
      try {
        response = await fetch(currentUrl, {
          method,
          headers: hopHeaders,
          body: body as BodyInit | undefined,
          signal: controller.signal,
          redirect: method === 'GET' ? 'manual' : 'follow',
        });
      } finally {
        clearTimeout(timeout);
      }
      const location = response.headers.get('location');
      if (method === 'GET' && response.status >= 300 && response.status < 400 && location) {
        await response.body?.cancel();
        currentUrl = new URL(location, currentUrl).toString();
        continue;
      }
      return response;
    }
    throw new CliError(`Too many redirects: ${url}`, { code: 'redirect_loop' });
  }

  /**
   * Upload a file. Files <= 4MB use a simple PUT to `endpoint`; larger files use an
   * upload session derived from it (`…:/name:/content` → `…:/name:/createUploadSession`),
   * sent in 10MB chunks that retry and resume from the server's nextExpectedRanges.
   */
  async uploadFile(
    endpoint: string,
    filePath: string,
    contentType: string,
    options?: { maxSimpleSize?: number; chunkSize?: number },
  ): Promise<GraphResponse> {
    if (this.readOnly) {
      throw new CliError('Read-only mode: file uploads are blocked. Remove --read-only to proceed.', { code: 'read_only', exitCode: ExitCode.PERMISSION });
    }

    const { statSync, readFileSync } = await import('fs');
    const fileSize = statSync(filePath).size;
    const simpleLimit = options?.maxSimpleSize ?? 4 * 1024 * 1024;
    // Chunks must be a multiple of 320 KiB; 10 MiB = 32 × 320 KiB.
    const chunkSize = options?.chunkSize ?? 10 * 1024 * 1024;
    const isLarge = fileSize > simpleLimit;
    const sessionEndpoint = uploadSessionEndpoint(endpoint);

    logger.info(`Uploading ${filePath} (${fileSize} bytes, ${contentType})`);

    if (isLarge && !sessionEndpoint) {
      throw new CliError(`File is ${fileSize} bytes; ${endpoint} only accepts uploads up to ${simpleLimit} bytes.`, { code: 'file_too_large' });
    }

    if (this.dryRun) {
      if (isLarge) {
        console.log(`[dry-run] POST ${this.baseUrl}${sessionEndpoint}`);
        console.log(`[dry-run] Then PUT ${Math.ceil(fileSize / chunkSize)} chunk(s) to the session uploadUrl`);
      } else {
        console.log(`[dry-run] PUT ${this.baseUrl}${endpoint}`);
      }
      console.log(`[dry-run] File: ${filePath} (${fileSize} bytes, ${contentType})`);
      return { data: { dryRun: true, method: isLarge ? 'POST' : 'PUT', fileSize, contentType }, status: 0 };
    }

    if (!isLarge) {
      const response = await this.request(endpoint, {
        method: 'PUT',
        headers: { 'Content-Type': contentType },
        body: readFileSync(filePath),
        responseType: 'json',
        timeoutMs: 120_000,
      });
      return response.data && typeof response.data === 'object' && 'message' in (response.data as object)
        ? { ...response, data: { message: 'Upload complete', size: fileSize } }
        : response;
    }

    return this.audited('PUT', endpoint, () => this.uploadInSession(sessionEndpoint!, filePath, fileSize, chunkSize));
  }

  /** Large upload through an upload session (audited as one PUT by uploadFile). */
  private async uploadInSession(sessionEndpoint: string, filePath: string, fileSize: number, chunkSize: number): Promise<GraphResponse> {
    const { openSync, readSync, closeSync } = await import('fs');
    const session = await this.send(sessionEndpoint, {
      method: 'POST',
      body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }),
    });
    const uploadUrl = (session.data as { uploadUrl?: string }).uploadUrl;
    if (!uploadUrl) {
      throw new CliError('Upload session response did not include an uploadUrl.', { code: 'bad_response' });
    }

    let offset = 0;
    let lastResult: unknown = null;
    let failures = 0;
    const fd = openSync(filePath, 'r');
    try {
      while (offset < fileSize) {
        const length = Math.min(chunkSize, fileSize - offset);
        const buffer = Buffer.alloc(length);
        readSync(fd, buffer, 0, length, offset);
        const rangeEnd = offset + length - 1;
        logger.debug(`Uploading bytes ${offset}-${rangeEnd}/${fileSize}`);

        let res: Response | null = null;
        let failure: string | null = null;
        try {
          const controller = new AbortController();
          const timer = setTimeout(() => controller.abort(), 120_000);
          try {
            // uploadUrl is pre-authorized: never send the bearer token to it.
            res = await fetch(uploadUrl, {
              method: 'PUT',
              headers: { 'Content-Length': String(length), 'Content-Range': `bytes ${offset}-${rangeEnd}/${fileSize}` },
              body: buffer,
              signal: controller.signal,
            });
          } finally {
            clearTimeout(timer);
          }
        } catch (err) {
          failure = err instanceof Error ? err.message : String(err);
        }

        if (res && (res.status === 200 || res.status === 201)) {
          lastResult = await res.json();
          offset = fileSize;
        } else if (res && res.status === 202) {
          const body = await res.json().catch(() => ({})) as { nextExpectedRanges?: string[] };
          offset = nextOffset(body.nextExpectedRanges) ?? offset + length;
          failures = 0;
        } else {
          const status = res?.status;
          const text = res ? await res.text() : '';
          const retryable = !res || status === 429 || (status !== undefined && status >= 500) || status === 416;
          if (!retryable || ++failures >= MAX_RETRIES) {
            throw new CliError(`Chunk upload failed at byte ${offset}${status ? ` (${status})` : ''}: ${failure ?? text}`,
              { code: 'upload_failed', status, exitCode: status === 429 ? ExitCode.THROTTLED : ExitCode.GENERIC });
          }
          const waitMs = status === 429
            ? parseRetryAfter(res!.headers.get('Retry-After'), Math.pow(2, failures) * 1000)
            : Math.pow(2, failures - 1) * 1000;
          logger.warn(`Chunk at byte ${offset} failed (${status ?? failure}); resuming in ${waitMs / 1000}s`);
          await sleep(waitMs);
          // Ask the session which bytes it still needs, then resume from there.
          const resumeFrom = await this.querySessionOffset(uploadUrl);
          if (resumeFrom !== null) offset = resumeFrom;
          continue;
        }

        const pct = Math.round((Math.min(offset, fileSize) / fileSize) * 100);
        console.error(`Uploading... ${pct}%`);
      }
    } finally {
      closeSync(fd);
    }

    if (lastResult) removeODataProps(lastResult);
    return { data: lastResult || { message: 'Upload complete', size: fileSize }, status: 200 };
  }

  /** GET the upload session status and return the first byte it still expects. */
  private async querySessionOffset(uploadUrl: string): Promise<number | null> {
    try {
      const res = await fetch(uploadUrl);
      if (!res.ok) return null;
      const body = await res.json() as { nextExpectedRanges?: string[] };
      return nextOffset(body.nextExpectedRanges);
    } catch {
      return null;
    }
  }

  /**
   * Send requests through JSON batching (POST /$batch, 20 per batch). Results come back in
   * input order. Items throttled with 429 — and 5xx on idempotent methods — are retried in a
   * new batch after the longest Retry-After, up to MAX_RETRIES rounds.
   */
  async batch(requests: BatchRequest[]): Promise<BatchResult[]> {
    if (this.readOnly && requests.some(r => r.method.toUpperCase() !== 'GET')) {
      throw new CliError('Read-only mode: batch contains write requests. Remove --read-only to proceed.', { code: 'read_only', exitCode: ExitCode.PERMISSION });
    }
    if (this.dryRun) {
      for (const r of requests) {
        console.log(`[dry-run] ${r.method.toUpperCase()} ${this.baseUrl}${r.url}`);
        if (r.body !== undefined) console.log(`[dry-run] Body: ${JSON.stringify(r.body)}`);
      }
      console.log(`[dry-run] (sent as ${Math.ceil(requests.length / BATCH_SIZE)} POST ${this.baseUrl}/$batch request(s))`);
      return requests.map(() => ({ status: 0, body: { dryRun: true } }));
    }

    const results: Array<BatchResult | undefined> = new Array(requests.length);
    let pending = requests.map((_, i) => i);
    for (let round = 0; round < MAX_RETRIES && pending.length > 0; round++) {
      const retry: number[] = [];
      let waitMs = 0;
      for (let start = 0; start < pending.length; start += BATCH_SIZE) {
        const chunk = pending.slice(start, start + BATCH_SIZE);
        const payload = {
          requests: chunk.map((idx) => {
            const r = requests[idx];
            const item: Record<string, unknown> = { id: String(idx), method: r.method.toUpperCase(), url: r.url };
            if (r.body !== undefined) {
              item.body = r.body;
              item.headers = { 'Content-Type': 'application/json', ...r.headers };
            } else if (r.headers) {
              item.headers = r.headers;
            }
            return item;
          }),
        };
        let res: GraphResponse;
        try {
          res = await this.send('/$batch', { method: 'POST', body: JSON.stringify(payload), responseType: 'json' });
        } catch (err) {
          const e = err as CliError;
          if (!NOT_SENT_CODES.has(e.code)) for (const idx of chunk) this.auditBatchItem(requests[idx], e.status ?? 0, e.requestId);
          throw err;
        }
        const responses = ((res.data as { responses?: Array<{ id: string; status: number; body?: unknown; headers?: Record<string, string> }> }).responses) ?? [];
        for (const r of responses) {
          const idx = Number(r.id);
          const method = requests[idx].method.toUpperCase();
          const retryable = r.status === 429 || (r.status >= 500 && IDEMPOTENT_METHODS.has(method));
          if (retryable && round < MAX_RETRIES - 1) {
            retry.push(idx);
            const ra = r.headers?.['Retry-After'] ?? r.headers?.['retry-after'] ?? null;
            waitMs = Math.max(waitMs, parseRetryAfter(ra, Math.pow(2, round) * 1000));
          } else {
            if (r.body && typeof r.body === 'object') removeODataProps(r.body);
            results[idx] = { status: r.status, body: r.body ?? null, headers: r.headers };
            this.auditBatchItem(requests[idx], r.status, r.headers?.['request-id'] ?? res.headers?.get('request-id') ?? undefined);
          }
        }
      }
      if (retry.length > 0) {
        logger.warn(`${retry.length} batch item(s) throttled or failed; retrying in ${waitMs / 1000}s`);
        await sleep(waitMs);
      }
      pending = retry;
    }
    // Array.from, not map: map skips the holes left by items Graph never answered
    return Array.from({ length: requests.length }, (_, i) => results[i] ?? { status: 0, body: { error: { message: 'No response for batch item' } } });
  }

  /**
   * Async generator that yields items page-by-page for streaming output.
   * Use this for large result sets to avoid buffering everything in memory.
   * Next pages reuse the original headers (e.g. ConsistencyLevel) and retry rules.
   */
  async *requestPagesStreaming(
    endpoint: string,
    options: GraphRequestOptions = {},
    pageLimit?: number,
    pageDelayMs?: number,
  ): AsyncGenerator<unknown[], void, undefined> {
    const maxPages = pageLimit || MAX_PAGES;
    const delayMs = pageDelayMs || 0;

    const first = await this.request(endpoint, options);
    const data = first.data as Record<string, unknown>;

    if (!data.value || !Array.isArray(data.value)) {
      yield [first.data];
      return;
    }

    yield data.value;
    let nextLink = first.nextLink;
    let pageCount = 1;

    while (nextLink && pageCount < maxPages) {
      if (delayMs > 0) await sleep(delayMs);
      const nextResponse = await this.request(nextLink, { headers: options.headers });
      const nextData = nextResponse.data as Record<string, unknown>;

      if (nextData.value && Array.isArray(nextData.value)) {
        yield nextData.value;
      }
      nextLink = nextResponse.nextLink;
      pageCount++;
    }

    if (pageCount >= maxPages && nextLink) {
      console.error(`Warning: Reached page limit (${maxPages}). Results may be incomplete.`);
    }

    logger.info(`Paginated across ${pageCount} pages (streamed)`);
  }

  async requestAllPages(
    endpoint: string,
    options: GraphRequestOptions = {},
    pageLimit?: number,
    pageDelayMs?: number,
  ): Promise<GraphResponse> {
    const allItems: unknown[] = [];
    for await (const page of this.requestPagesStreaming(endpoint, options, pageLimit, pageDelayMs)) {
      allItems.push(...page);
    }
    return { data: { value: allItems }, status: 200 };
  }
}

function isTextContentType(contentType: string): boolean {
  return contentType.startsWith('text/') || /\b(xml|html)\b/.test(contentType);
}

/** Decode a successful response according to the caller's responseType. */
async function decodeResponse(response: Response, responseType: ResponseType): Promise<GraphResponse> {
  const contentType = response.headers.get('content-type') || '';
  const base = { status: response.status, headers: response.headers };

  if (responseType === 'stream') {
    return { ...base, data: { contentType, size: Number(response.headers.get('content-length')) || undefined }, stream: response.body };
  }

  const isJson = contentType.includes('json');
  if (responseType === 'binary' || (responseType === 'auto' && contentType && !isJson && !isTextContentType(contentType))) {
    const buffer = Buffer.from(await response.arrayBuffer());
    return {
      ...base,
      data: { contentBytes: buffer.toString('base64'), contentType, size: buffer.byteLength },
    };
  }

  const text = await response.text();
  if (responseType === 'text') {
    return { ...base, data: text };
  }
  if (!text) {
    return { ...base, data: { message: 'OK' } };
  }

  try {
    const json = JSON.parse(text);
    const meta = extractODataMeta(json);
    removeODataProps(json);
    return { ...base, data: json, ...meta };
  } catch {
    if (responseType === 'json') {
      throw new CliError(`Expected JSON response but got ${contentType || 'unknown content type'}`, { code: 'bad_response' });
    }
    return { ...base, data: { rawResponse: text } };
  }
}

function extractODataMeta(json: unknown): { nextLink?: string; deltaLink?: string; count?: number } {
  if (typeof json !== 'object' || json === null || Array.isArray(json)) return {};
  const o = json as Record<string, unknown>;
  const meta: { nextLink?: string; deltaLink?: string; count?: number } = {};
  if (typeof o['@odata.nextLink'] === 'string') meta.nextLink = o['@odata.nextLink'];
  if (typeof o['@odata.deltaLink'] === 'string') meta.deltaLink = o['@odata.deltaLink'];
  if (typeof o['@odata.count'] === 'number') meta.count = o['@odata.count'];
  return meta;
}

/** Map a failed Graph response to a CliError with the matching exit code. */
export function graphError(
  status: number,
  message: string,
  details: { status?: number; graphCode?: string; requestId?: string } = {},
): CliError {
  const d = { ...details, status };
  if (status === 401) {
    return new CliError(`Authentication failed. Run: ms365 auth login`,
      { ...d, code: 'auth_failed', exitCode: ExitCode.AUTH, fix: 'ms365 auth login' });
  }
  if (status === 403) {
    const hint = message.includes('scope') || message.includes('permission') ? ' Try running with --org-mode.' : '';
    return new CliError(`Permission denied (403): ${message}${hint}`,
      { ...d, code: 'permission_denied', exitCode: ExitCode.PERMISSION });
  }
  if (status === 404) {
    return new CliError(`Graph API error (404): ${message}`, { ...d, code: 'not_found', exitCode: ExitCode.NOT_FOUND });
  }
  if (status === 429) {
    return new CliError(`Graph API error (429): ${message}`, { ...d, code: 'throttled', exitCode: ExitCode.THROTTLED });
  }
  return new CliError(`Graph API error (${status}): ${message}`, { ...d, code: 'graph_error' });
}

/**
 * Derive the upload-session endpoint from a content endpoint:
 * `/drives/D/items/P:/a.zip:/content` → `/drives/D/items/P:/a.zip:/createUploadSession`.
 * Returns null for endpoints that don't support sessions (e.g. `/me/photo/$value`).
 */
export function uploadSessionEndpoint(endpoint: string): string | null {
  const m = endpoint.match(/^(.*):\/content$/);
  return m ? `${m[1]}:/createUploadSession` : null;
}

/** First byte offset from nextExpectedRanges like ["26-", "40-60"]. */
function nextOffset(ranges?: string[]): number | null {
  if (!ranges || ranges.length === 0) return null;
  const start = parseInt(ranges[0].split('-')[0], 10);
  return Number.isFinite(start) ? start : null;
}

/** Keys that are transport metadata, not resource data. Type/etag/count are kept. */
const STRIPPED_ODATA_KEYS = new Set(['@odata.context', '@odata.nextLink', '@odata.deltaLink']);

export function removeODataProps(obj: unknown): void {
  if (typeof obj !== 'object' || obj === null) return;
  if (Array.isArray(obj)) {
    obj.forEach(removeODataProps);
    return;
  }
  for (const key of Object.keys(obj as Record<string, unknown>)) {
    if (STRIPPED_ODATA_KEYS.has(key)) {
      delete (obj as Record<string, unknown>)[key];
    } else {
      removeODataProps((obj as Record<string, unknown>)[key]);
    }
  }
}

function parseRetryAfter(header: string | null, fallbackMs: number): number {
  if (!header) return fallbackMs;
  const seconds = parseInt(header, 10);
  if (!isNaN(seconds)) return seconds * 1000;
  // Retry-After can be an HTTP date
  const date = Date.parse(header);
  if (!isNaN(date)) return Math.max(date - Date.now(), 1000);
  return fallbackMs;
}

function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}
