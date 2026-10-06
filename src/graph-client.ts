import type { AuthManager } from './auth.js';
import logger from './logger.js';
import { info } from './utils.js';

/** Attempts per request (first try + retries) for retryable failures. */
const MAX_ATTEMPTS = 3;
const MAX_PAGES = 100;
const DEFAULT_TIMEOUT_MS = 30_000;

/** Methods that are safe to repeat after a timeout or 5xx. POST/PATCH are retried only on 429. */
const IDEMPOTENT_METHODS: Record<string, true> = { GET: true, HEAD: true, PUT: true, DELETE: true };

export interface GraphRequestOptions {
  method?: string;
  headers?: Record<string, string>;
  body?: string;
  queryParams?: Record<string, string>;
}

export interface GraphResponse {
  data: unknown;
  status: number;
  nextLink?: string;
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
    let url = `${this.baseUrl}${endpoint}`;

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
    const url = this.buildUrl(endpoint, options.queryParams);
    const method = (options.method || 'GET').toUpperCase();

    logger.debug(`${method} ${url}`, options.body ? { bodyLength: options.body.length } : undefined);

    if (this.readOnly && method !== 'GET') {
      throw new Error(`Read-only mode: ${method} requests are blocked. Remove --read-only to proceed.`);
    }

    if (this.dryRun) {
      console.log(`[dry-run] ${method} ${url}`);
      if (options.body) {
        console.log(`[dry-run] Body: ${options.body}`);
      }
      if (options.headers && Object.keys(options.headers).length > 0) {
        console.log(`[dry-run] Headers: ${JSON.stringify(options.headers)}`);
      }
      return { data: { dryRun: true, method, url }, status: 0 };
    }

    const response = await this.send(url, {
      method,
      headers: { 'Content-Type': 'application/json', ...options.headers },
      body: options.body,
    });
    return parseResponse(response);
  }

  /**
   * The single request core: auth header, timeout, retry policy, 401 refresh, error mapping.
   * Resolves only with a 2xx response (body unread); throws GraphError otherwise.
   * `auth: false` sends no Authorization header (pre-authenticated upload URLs).
   */
  private async send(url: string, init: SendInit, auth = true): Promise<Response> {
    const method = init.method.toUpperCase();
    const idempotent = IDEMPOTENT_METHODS[method] === true;
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
   * Stream a file upload using chunked reads instead of loading into memory.
   * For files <= 4MB, uses simple PUT. For larger files, uses upload session.
   */
  async uploadFile(
    endpoint: string,
    filePath: string,
    contentType: string,
    options?: { maxSimpleSize?: number },
  ): Promise<GraphResponse> {
    if (this.readOnly) {
      throw new Error('Read-only mode: file uploads are blocked. Remove --read-only to proceed.');
    }

    const { createReadStream, statSync } = await import('fs');
    const stat = statSync(filePath);
    const fileSize = stat.size;

    logger.info(`Uploading ${filePath} (${fileSize} bytes, ${contentType})`);

    if (this.dryRun) {
      console.log(`[dry-run] PUT ${this.baseUrl}${endpoint}`);
      console.log(`[dry-run] File: ${filePath} (${fileSize} bytes, ${contentType})`);
      return { data: { dryRun: true, method: 'PUT', fileSize, contentType }, status: 0 };
    }

    const token = await this.auth.getToken();

    const simpleLimit = options?.maxSimpleSize ?? 4 * 1024 * 1024;
    if (fileSize <= simpleLimit) {
      // Simple upload for files <= 4MB
      const { readFileSync } = await import('fs');
      const content = readFileSync(filePath);

      const response = await fetch(`${this.baseUrl}${endpoint}`, {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': contentType,
          'Content-Length': String(fileSize),
        },
        body: content,
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Upload failed (${response.status}): ${errorText}`);
      }

      const text = await response.text();
      try {
        const json = JSON.parse(text);
        removeODataProps(json);
        return { data: json, status: response.status };
      } catch {
        return { data: { message: 'Upload complete', size: fileSize }, status: response.status };
      }
    }

    // Large file upload session for > 4MB
    const sessionResponse = await fetch(
      `${this.baseUrl}${endpoint}/createUploadSession`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          item: { '@microsoft.graph.conflictBehavior': 'replace' },
        }),
      }
    );

    if (!sessionResponse.ok) {
      const errorText = await sessionResponse.text();
      throw new Error(`Failed to create upload session (${sessionResponse.status}): ${errorText}`);
    }

    const session = await sessionResponse.json() as { uploadUrl: string };
    const chunkSize = 10 * 1024 * 1024; // 10MB chunks
    let offset = 0;
    let lastResult: unknown = null;

    const { openSync, readSync, closeSync } = await import('fs');
    const fd = openSync(filePath, 'r');

    try {
      while (offset < fileSize) {
        const length = Math.min(chunkSize, fileSize - offset);
        const buffer = Buffer.alloc(length);
        readSync(fd, buffer, 0, length, offset);

        const rangeEnd = offset + length - 1;
        logger.debug(`Uploading bytes ${offset}-${rangeEnd}/${fileSize}`);

        const chunkResponse = await fetch(session.uploadUrl, {
          method: 'PUT',
          headers: {
            'Content-Length': String(length),
            'Content-Range': `bytes ${offset}-${rangeEnd}/${fileSize}`,
          },
          body: buffer,
        });

        if (!chunkResponse.ok && chunkResponse.status !== 202) {
          const errorText = await chunkResponse.text();
          throw new Error(`Chunk upload failed at byte ${offset} (${chunkResponse.status}): ${errorText}`);
        }

        if (chunkResponse.status === 200 || chunkResponse.status === 201) {
          lastResult = await chunkResponse.json();
        } else {
          // Drain the response body to prevent connection leaks on 202
          await chunkResponse.text();
        }

        offset += length;
        const pct = Math.round((offset / fileSize) * 100);
        console.error(`Uploading... ${pct}%`);
      }
    } finally {
      closeSync(fd);
    }

    if (lastResult) {
      removeODataProps(lastResult as Record<string, unknown>);
    }

    return {
      data: lastResult || { message: 'Upload complete', size: fileSize },
      status: 200,
    };
  }

  /**
   * Async generator that yields items page-by-page for streaming output.
   * Use this for large result sets to avoid buffering everything in memory.
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
      const nextResponse = await this.requestNext(nextLink, options);
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

  try {
    const json = JSON.parse(text);
    const nextLink = json['@odata.nextLink'] || undefined;
    removeODataProps(json);
    return { data: json, status: response.status, nextLink };
  } catch {
    return { data: { rawResponse: text }, status: response.status };
  }
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

function removeODataProps(obj: unknown): void {
  if (typeof obj !== 'object' || obj === null) return;
  if (Array.isArray(obj)) {
    obj.forEach(removeODataProps);
    return;
  }
  for (const key of Object.keys(obj as Record<string, unknown>)) {
    if (key.startsWith('@odata.')) {
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
