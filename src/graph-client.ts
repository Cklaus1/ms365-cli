import { closeSync, openSync, readFileSync, readSync, statSync } from 'node:fs';
import path from 'node:path';
import type { AuthManager } from './auth.js';
import logger from './logger.js';
import { detectMimeType } from './mime.js';
import { formatSize, info } from './utils.js';

/** Attempts per request (first try + retries) for retryable failures. */
const MAX_ATTEMPTS = 3;
const MAX_PAGES = 100;
const DEFAULT_TIMEOUT_MS = 30_000;
/** Largest file sent as one PUT; Graph rejects simple uploads above 4 MiB. */
const SIMPLE_UPLOAD_MAX = 4 * 1024 * 1024;
/** Upload-session chunk size: must be a multiple of 320 KiB. */
const UPLOAD_CHUNK_SIZE = 32 * 320 * 1024;
const UPLOAD_TIMEOUT_MS = 120_000;

/** Methods that are safe to repeat after a timeout or 5xx. POST/PATCH are retried only on 429. */
const IDEMPOTENT_METHODS: Record<string, true> = { GET: true, HEAD: true, PUT: true, DELETE: true };

/**
 * POST actions that only read data; allowed under --read-only. Matched against the endpoint path
 * (query string ignored): Microsoft Search, free/busy, and meeting-time suggestions.
 */
const READ_ONLY_POST_ACTIONS: RegExp[] = [
  /^\/search\/query$/,
  /^\/(me|users\/[^/]+)\/calendar\/getSchedule$/,
  /^\/(me|users\/[^/]+)\/findMeetingTimes$/,
];

/** Whether `--read-only` permits this request. */
export function isReadOnlyAllowed(method: string, endpoint: string): boolean {
  const m = method.toUpperCase();
  if (m === 'GET' || m === 'HEAD') return true;
  const path = endpoint.split('?')[0];
  return m === 'POST' && READ_ONLY_POST_ACTIONS.some((re) => re.test(path));
}

export type ApiVersion = 'v1.0' | 'beta';

export interface GraphRequestOptions {
  method?: string;
  headers?: Record<string, string>;
  body?: string;
  queryParams?: Record<string, string>;
  /** Overrides the client's API version for this request (and its nextLink pages, which are absolute). */
  apiVersion?: ApiVersion;
}

/**
 * `data` is the raw Graph payload, `@odata.*` annotations included (`@odata.etag`, `@odata.type`, …);
 * rendering strips them (formatter). The paging/delta/count annotations are also lifted out here.
 */
export interface GraphResponse {
  data: unknown;
  /** HTTP status; 0 for a dry-run preview. */
  status: number;
  /** `@odata.nextLink`: absolute URL of the next page (for requestAllPages: the first page not fetched). */
  nextLink?: string;
  /** `@odata.deltaLink`: absolute URL to resume a delta query (last page). */
  deltaLink?: string;
  /** `@odata.count` (requires `$count=true`). */
  count?: number;
}

/** An unparsed response body (GRAPH-2). `body` is null for a dry-run preview or an empty body. */
export interface BinaryResponse {
  body: ReadableStream<Uint8Array> | null;
  /** HTTP status; 0 for a dry-run preview. */
  status: number;
  contentType: string;
  /** From Content-Length, when the server sent one. */
  contentLength?: number;
}

/** A non-2xx answer from Graph after the retry policy gave up. */
export class GraphError extends Error {
  constructor(message: string, readonly status: number, readonly code?: string) {
    super(message);
    this.name = 'GraphError';
  }
}

interface SendInit {
  method: string;
  headers: Record<string, string>;
  body?: string | Uint8Array<ArrayBuffer>;
  timeoutMs?: number;
  /** Overrides the per-method default: may a timeout/5xx be retried? */
  idempotent?: boolean;
}

export type ConflictBehavior = 'rename' | 'replace' | 'fail';

export interface UploadOptions {
  /** `@microsoft.graph.conflictBehavior` when the target name exists. */
  conflict: ConflictBehavior;
  /** Defaults to the MIME type detected from the local file name. */
  contentType?: string;
  /** Files up to this size use one PUT (default 4 MiB); larger ones an upload session. */
  maxSimpleSize?: number;
  /** Upload-session chunk size (default 10 MiB; Graph requires multiples of 320 KiB). */
  chunkSize?: number;
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

  buildUrl(endpoint: string, queryParams?: Record<string, string>, apiVersion?: ApiVersion): string {
    let url = apiVersion ? `https://graph.microsoft.com/${apiVersion}${endpoint}` : `${this.baseUrl}${endpoint}`;

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

  async request(endpoint: string, options: GraphRequestOptions = {}): Promise<GraphResponse> {
    const url = this.buildUrl(endpoint, options.queryParams, options.apiVersion);
    const method = (options.method || 'GET').toUpperCase();

    logger.debug(`${method} ${url}`, options.body ? { bodyLength: options.body.length } : undefined);

    this.assertAllowed(method, endpoint);

    if (this.dryRun) {
      const data: Record<string, unknown> = { dryRun: true, method, url };
      if (options.body) data.body = parseJsonOr(options.body);
      this.preview(method, url, options);
      return { data, status: 0 };
    }

    const response = await this.send(url, {
      method,
      headers: { 'Content-Type': 'application/json', ...options.headers },
      body: options.body,
    });
    return parseResponse(response);
  }

  /**
   * GET (or `options.method`) whose body is never parsed or decoded: returns the byte stream as-is,
   * whatever the content type (a `.json` file comes back verbatim). Same auth/timeout/retry/401 core
   * and read-only/dry-run handling as `request()`; dry-run returns `{ body: null, status: 0 }`.
   */
  async requestBinary(endpoint: string, options: GraphRequestOptions = {}): Promise<BinaryResponse> {
    const url = this.buildUrl(endpoint, options.queryParams, options.apiVersion);
    const method = (options.method || 'GET').toUpperCase();
    logger.debug(`${method} ${url} (binary)`);

    this.assertAllowed(method, endpoint);
    if (this.dryRun) {
      this.preview(method, url, options);
      return { body: null, status: 0, contentType: '' };
    }

    const response = await this.send(url, { method, headers: { ...options.headers }, body: options.body });
    const length = response.headers.get('content-length');
    return {
      body: response.body,
      status: response.status,
      contentType: response.headers.get('content-type') ?? '',
      contentLength: length === null ? undefined : Number(length),
    };
  }

  /** Dry-run preview, on stderr unconditionally: --quiet must not hide what would be sent, stdout stays data-only. */
  private preview(method: string, url: string, options: { body?: string; headers?: Record<string, string> }, extra?: string): void {
    process.stderr.write(`[dry-run] ${method} ${url}\n`);
    if (options.body) process.stderr.write(`[dry-run] Body: ${options.body}\n`);
    if (options.headers && Object.keys(options.headers).length > 0) {
      process.stderr.write(`[dry-run] Headers: ${JSON.stringify(options.headers)}\n`);
    }
    if (extra) process.stderr.write(`[dry-run] ${extra}\n`);
  }

  /**
   * The single request core: auth header, timeout, retry policy, 401 refresh, error mapping.
   * Resolves only with a 2xx response (body unread); throws GraphError otherwise.
   * `auth: false` sends no Authorization header (pre-authenticated upload URLs).
   */
  private async send(url: string, init: SendInit, auth = true): Promise<Response> {
    const method = init.method.toUpperCase();
    const idempotent = init.idempotent ?? IDEMPOTENT_METHODS[method] === true;
    const timeoutMs = init.timeoutMs ?? DEFAULT_TIMEOUT_MS;
    const headers = { ...init.headers };
    if (auth) headers.Authorization = `Bearer ${await this.auth.getToken()}`;
    let refreshed = false;

    for (let attempt = 1; ; attempt++) {
      let response: Response;
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs);
      try {
        response = await fetch(url, { method, headers, body: init.body, signal: controller.signal });
      } catch (err: unknown) {
        if (!(err instanceof Error && err.name === 'AbortError')) throw err;
        const what = `Request timed out after ${timeoutMs / 1000}s: ${method} ${url}`;
        if (!idempotent) {
          throw new Error(`${what}. The request may have been applied; not retried — check before re-running.`);
        }
        if (attempt >= MAX_ATTEMPTS) throw new Error(`${what} (${attempt} attempts)`);
        logger.warn(`${what}. Retrying...`);
        info(`Request timed out. Retrying (${attempt + 1}/${MAX_ATTEMPTS})...`);
        continue;
      } finally {
        clearTimeout(timer);
      }

      logger.debug(`Response: ${response.status} ${response.statusText}`, { attempt });
      if (response.ok) return response;

      if (response.status === 401 && auth && !refreshed) {
        refreshed = true;
        await response.body?.cancel();
        let fresh: string;
        try {
          fresh = await this.auth.getToken(true);
        } catch {
          throw new GraphError('Authentication failed. Run: ms365 auth login', 401);
        }
        headers.Authorization = `Bearer ${fresh}`;
        logger.info('Token refreshed, retrying request...');
        attempt--; // the refresh retry does not consume an attempt
        continue;
      }

      const retryable = response.status === 429 || (response.status >= 500 && idempotent);
      if (retryable && attempt < MAX_ATTEMPTS) {
        const waitMs = parseRetryAfter(response.headers.get('Retry-After'), Math.pow(2, attempt - 1) * 1000);
        await response.body?.cancel();
        const reason = response.status === 429 ? 'Rate limited (429)' : `Server error (${response.status})`;
        logger.warn(`${reason}. Retrying in ${waitMs / 1000}s...`);
        info(`${reason}. Retrying in ${waitMs / 1000}s...`);
        await sleep(waitMs);
        continue;
      }

      throw await toGraphError(response, method, idempotent, attempt);
    }
  }

  /** GET an absolute nextLink URL through the same core, with the caller's headers. */
  private async requestNext(nextLink: string, options: GraphRequestOptions): Promise<GraphResponse> {
    logger.debug(`GET ${nextLink}`);
    const response = await this.send(nextLink, {
      method: 'GET',
      headers: { 'Content-Type': 'application/json', ...options.headers },
    });
    return parseResponse(response);
  }

  /**
   * Upload `localPath` to the drive item addressed by `itemPath`: a colon path WITHOUT the trailing
   * colon whose segments are already encoded (`/me/drive/root:/Docs/a.pdf`, `/drives/{d}/items/{p}:/a.pdf`).
   * Files <= `maxSimpleSize` go as one PUT to `{itemPath}:/content?@microsoft.graph.conflictBehavior=<c>`;
   * larger ones via `{itemPath}:/createUploadSession` (`item['@microsoft.graph.conflictBehavior']`) in
   * chunks that resume from the server's `nextExpectedRanges` after a failure; a session that cannot
   * finish is cancelled (DELETE). `data` is the stored driveItem; an `info()` line names the stored file
   * and flags a rename. With conflict `rename` a timed-out/5xx simple PUT is not retried (a retry could
   * store a second copy). Read-only blocks it; dry-run previews the first request on stderr.
   */
  async uploadDriveItem(itemPath: string, localPath: string, opts: UploadOptions): Promise<GraphResponse> {
    this.assertAllowed('PUT', itemPath);
    const fileSize = statSync(localPath).size;
    const contentType = opts.contentType ?? detectMimeType(path.basename(localPath));
    const fileNote = `File: ${localPath} (${fileSize} bytes, ${contentType})`;

    let result: GraphResponse;
    if (fileSize <= (opts.maxSimpleSize ?? SIMPLE_UPLOAD_MAX)) {
      const url = `${this.buildUrl(`${itemPath}:/content`)}?@microsoft.graph.conflictBehavior=${opts.conflict}`;
      if (this.dryRun) {
        this.preview('PUT', url, {}, fileNote);
        return { data: { dryRun: true, method: 'PUT', url, fileSize, contentType, conflict: opts.conflict }, status: 0 };
      }
      const response = await this.send(url, {
        method: 'PUT',
        headers: { 'Content-Type': contentType },
        body: readFileSync(localPath),
        timeoutMs: UPLOAD_TIMEOUT_MS,
        idempotent: opts.conflict !== 'rename',
      });
      result = await parseResponse(response);
    } else {
      const url = this.buildUrl(`${itemPath}:/createUploadSession`);
      const body = JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': opts.conflict } });
      if (this.dryRun) {
        this.preview('POST', url, { body }, fileNote);
        return { data: { dryRun: true, method: 'POST', url, fileSize, contentType, conflict: opts.conflict }, status: 0 };
      }
      const session = await parseResponse(await this.send(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body,
      }));
      const uploadUrl = stringField(session.data, 'uploadUrl');
      if (!uploadUrl) throw new Error('Upload session response has no uploadUrl.');
      result = await this.uploadChunks(uploadUrl, localPath, fileSize, opts.chunkSize ?? UPLOAD_CHUNK_SIZE);
    }

    const requested = decodeSegment(itemPath.slice(itemPath.lastIndexOf('/') + 1));
    const stored = stringField(result.data, 'name') ?? requested;
    info(stored === requested
      ? `Uploaded "${stored}" (${formatSize(fileSize)}).`
      : `Uploaded as "${stored}" (requested "${requested}"; conflict behavior ${opts.conflict}).`);
    return result;
  }

  /** Session chunk loop; resumes from `nextExpectedRanges` after a failed chunk, cancels the session on giving up. */
  private async uploadChunks(uploadUrl: string, localPath: string, fileSize: number, chunkSize: number): Promise<GraphResponse> {
    const fd = openSync(localPath, 'r');
    let offset = 0;
    let failures = 0;
    try {
      for (;;) {
        const length = Math.min(chunkSize, fileSize - offset);
        const chunk = Buffer.alloc(length);
        readSync(fd, chunk, 0, length, offset);
        const end = offset + length - 1;
        logger.debug(`Uploading bytes ${offset}-${end}/${fileSize}`);

        let response: Response;
        try {
          // The uploadUrl is pre-authenticated: never send it the bearer token.
          response = await this.send(uploadUrl, {
            method: 'PUT',
            headers: { 'Content-Range': `bytes ${offset}-${end}/${fileSize}` },
            body: chunk,
            timeoutMs: UPLOAD_TIMEOUT_MS,
          }, false);
        } catch (err: unknown) {
          if (++failures >= MAX_ATTEMPTS) throw err;
          const resumeAt = await this.sessionOffset(uploadUrl);
          if (resumeAt === undefined) throw err;
          offset = resumeAt;
          info(`Chunk upload failed (${err instanceof Error ? err.message : String(err)}). Resuming at byte ${offset}...`);
          continue;
        }

        if (response.status === 200 || response.status === 201) return await parseResponse(response);
        failures = 0;
        offset = nextExpectedOffset((await parseResponse(response)).data) ?? offset + length;
        info(`Uploading... ${Math.round((offset / fileSize) * 100)}%`);
      }
    } catch (err: unknown) {
      await this.cancelSession(uploadUrl);
      throw err;
    } finally {
      closeSync(fd);
    }
  }

  /** The first byte the session still expects, from GET uploadUrl; undefined if the session is unusable. */
  private async sessionOffset(uploadUrl: string): Promise<number | undefined> {
    try {
      const status = await this.send(uploadUrl, { method: 'GET', headers: {} }, false);
      return nextExpectedOffset((await parseResponse(status)).data);
    } catch (err: unknown) {
      logger.warn(`Upload session status failed: ${err instanceof Error ? err.message : String(err)}`);
      return undefined;
    }
  }

  private async cancelSession(uploadUrl: string): Promise<void> {
    try {
      const response = await this.send(uploadUrl, { method: 'DELETE', headers: {} }, false);
      await response.body?.cancel();
    } catch (err: unknown) {
      logger.warn(`Cancelling upload session failed: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  /**
   * Plain PUT of a local file's bytes to `endpoint` (e.g. `/me/photo/$value`, URL used exactly as given).
   * Same core as `request()` (retry, 401 refresh, GraphError on non-2xx); read-only blocks it; dry-run
   * previews on stderr and returns status 0. Callers enforce size limits.
   */
  async putBinary(endpoint: string, localPath: string, contentType: string): Promise<GraphResponse> {
    this.assertAllowed('PUT', endpoint);
    const url = this.buildUrl(endpoint);
    const fileSize = statSync(localPath).size;
    if (this.dryRun) {
      this.preview('PUT', url, {}, `File: ${localPath} (${fileSize} bytes, ${contentType})`);
      return { data: { dryRun: true, method: 'PUT', url, fileSize, contentType }, status: 0 };
    }
    const response = await this.send(url, {
      method: 'PUT',
      headers: { 'Content-Type': contentType },
      body: readFileSync(localPath),
      timeoutMs: UPLOAD_TIMEOUT_MS,
    });
    return parseResponse(response);
  }

  private assertAllowed(method: string, endpoint: string): void {
    if (this.readOnly && !isReadOnlyAllowed(method, endpoint)) {
      throw new Error(`Read-only mode: ${method} requests are blocked. Remove --read-only to proceed.`);
    }
  }

  /**
   * Yields one GraphResponse per page: the first via `request()`, then each `@odata.nextLink`
   * through the same core with the caller's headers, up to `pageLimit` pages (default 100).
   */
  private async *pageResponses(
    endpoint: string,
    options: GraphRequestOptions,
    pageLimit?: number,
    pageDelayMs?: number,
  ): AsyncGenerator<GraphResponse, void, undefined> {
    const maxPages = pageLimit || MAX_PAGES;
    const delayMs = pageDelayMs || 0;

    let page = await this.request(endpoint, options);
    yield page;
    let pageCount = 1;

    while (page.nextLink && pageCount < maxPages) {
      if (delayMs > 0) await sleep(delayMs);
      page = await this.requestNext(page.nextLink, options);
      yield page;
      pageCount++;
    }

    if (page.nextLink) {
      console.error(`Warning: Reached page limit (${maxPages}). Results may be incomplete.`);
    }
    logger.info(`Paginated across ${pageCount} pages`);
  }

  /**
   * Async generator that yields items page-by-page for streaming output.
   * Use this for large result sets to avoid buffering everything in memory.
   * A non-collection payload is yielded once as a single-item page.
   */
  async *requestPagesStreaming(
    endpoint: string,
    options: GraphRequestOptions = {},
    pageLimit?: number,
    pageDelayMs?: number,
  ): AsyncGenerator<unknown[], void, undefined> {
    for await (const page of this.pageResponses(endpoint, options, pageLimit, pageDelayMs)) {
      const items = collectionItems(page.data);
      if (!items) {
        yield [page.data];
        return;
      }
      yield items;
    }
  }

  /**
   * Fetch every page into `{ value: [...] }`. `count` comes from the first page, `deltaLink` from the
   * last; `nextLink` is set only when `pageLimit` stopped paging early (resume point).
   */
  async requestAllPages(
    endpoint: string,
    options: GraphRequestOptions = {},
    pageLimit?: number,
    pageDelayMs?: number,
  ): Promise<GraphResponse> {
    const allItems: unknown[] = [];
    let first: GraphResponse | undefined;
    let last: GraphResponse | undefined;
    for await (const page of this.pageResponses(endpoint, options, pageLimit, pageDelayMs)) {
      first ??= page;
      last = page;
      const items = collectionItems(page.data);
      if (!items) {
        allItems.push(page.data);
        break;
      }
      allItems.push(...items);
    }
    return {
      data: { value: allItems },
      status: last?.status ?? 200,
      nextLink: last?.nextLink,
      deltaLink: last?.deltaLink,
      count: first?.count,
    };
  }
}

function collectionItems(data: unknown): unknown[] | undefined {
  if (typeof data === 'object' && data !== null && 'value' in data && Array.isArray(data.value)) return data.value;
  return undefined;
}

function stringField(data: unknown, key: string): string | undefined {
  if (typeof data !== 'object' || data === null || !(key in data)) return undefined;
  const value: unknown = (data as Record<string, unknown>)[key];
  return typeof value === 'string' ? value : undefined;
}

/** Start of the first `nextExpectedRanges` entry (`"26-"` or `"26-49"` -> 26) of an upload-session payload. */
function nextExpectedOffset(data: unknown): number | undefined {
  if (typeof data !== 'object' || data === null || !('nextExpectedRanges' in data)) return undefined;
  const ranges = data.nextExpectedRanges;
  if (!Array.isArray(ranges) || typeof ranges[0] !== 'string') return undefined;
  const start = Number.parseInt(ranges[0], 10);
  return Number.isNaN(start) ? undefined : start;
}

function decodeSegment(segment: string): string {
  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
}

function parseJsonOr(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

async function parseResponse(response: Response): Promise<GraphResponse> {
  // Handle binary content
  const contentType = response.headers.get('content-type') || '';
  if (
    contentType.startsWith('audio/') ||
    contentType.startsWith('video/') ||
    contentType.startsWith('image/') ||
    contentType === 'application/octet-stream'
  ) {
    const buffer = await response.arrayBuffer();
    return {
      data: {
        contentBytes: Buffer.from(buffer).toString('base64'),
        contentType,
        size: buffer.byteLength,
      },
      status: response.status,
    };
  }

  const text = await response.text();
  if (!text) {
    return { data: { message: 'OK' }, status: response.status };
  }

  let json: unknown;
  try {
    json = JSON.parse(text);
  } catch {
    return { data: { rawResponse: text }, status: response.status };
  }
  const result: GraphResponse = { data: json, status: response.status };
  if (typeof json === 'object' && json !== null) {
    if ('@odata.nextLink' in json && typeof json['@odata.nextLink'] === 'string') result.nextLink = json['@odata.nextLink'];
    if ('@odata.deltaLink' in json && typeof json['@odata.deltaLink'] === 'string') result.deltaLink = json['@odata.deltaLink'];
    if ('@odata.count' in json && typeof json['@odata.count'] === 'number') result.count = json['@odata.count'];
  }
  return result;
}

/** Map a final non-2xx response to a user-facing GraphError. */
async function toGraphError(response: Response, method: string, idempotent: boolean, attempts: number): Promise<GraphError> {
  const status = response.status;
  const text = await response.text().catch(() => '');
  let message = text || response.statusText;
  let code: string | undefined;
  try {
    const parsed: unknown = JSON.parse(text);
    if (typeof parsed === 'object' && parsed !== null && 'error' in parsed) {
      const error = parsed.error;
      if (typeof error === 'object' && error !== null) {
        if ('message' in error && typeof error.message === 'string' && error.message) message = error.message;
        if ('code' in error && typeof error.code === 'string') code = error.code;
      }
    }
  } catch {
    // not JSON: keep the raw text
  }
  logger.error(`API error ${status}: ${message}`);

  if (status === 401) return new GraphError('Authentication failed. Run: ms365 auth login', status, code);
  if (status === 403) {
    const hint = message.includes('scope') || message.includes('permission') ? ' Try running with --org-mode.' : '';
    return new GraphError(`Permission denied (403): ${message}${hint}`, status, code);
  }
  let suffix = '';
  if (status >= 500 && !idempotent) suffix = ` — the ${method} request may have been applied; not retried.`;
  else if (attempts > 1) suffix = ` (after ${attempts} attempts)`;
  return new GraphError(`Graph API error (${status}): ${message}${suffix}`, status, code);
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
