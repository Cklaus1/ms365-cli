import { AuthManager } from './auth.js';
import logger from './logger.js';

const MAX_RETRIES = 3;
const MAX_PAGES = 100;

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
    const method = options.method || 'GET';

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

    const token = await this.auth.getToken();

    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...options.headers,
    };

    let lastError: Error | null = null;
    let retriedWith401 = false;

    for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 30_000);

      let response: Response;
      try {
        response = await fetch(url, {
          method,
          headers,
          body: options.body,
          signal: controller.signal,
        });
      } catch (err: unknown) {
        clearTimeout(timeout);
        if (err instanceof Error && err.name === 'AbortError') {
          lastError = new Error(`Request timed out after 30s: ${method} ${url}`);
          if (attempt < MAX_RETRIES - 1) {
            logger.warn(`Request timed out. Retrying...`);
            continue;
          }
          break;
        }
        throw err;
      }
      clearTimeout(timeout);

      logger.debug(`Response: ${response.status} ${response.statusText}`, { attempt });

      // Rate limited — retry with backoff
      if (response.status === 429) {
        const waitMs = parseRetryAfter(response.headers.get('Retry-After'), Math.pow(2, attempt) * 1000);
        lastError = new Error(`Rate limited (429) after ${attempt + 1} attempts`);
        if (attempt < MAX_RETRIES - 1) {
          logger.warn(`Rate limited. Retrying in ${waitMs / 1000}s...`);
          console.error(`Rate limited. Retrying in ${waitMs / 1000}s...`);
          await sleep(waitMs);
          continue;
        }
      }

      // Server error — retry with backoff
      if (response.status >= 500) {
        lastError = new Error(`Server error (${response.status}) after ${attempt + 1} attempts`);
        if (attempt < MAX_RETRIES - 1) {
          const waitMs = Math.pow(2, attempt) * 1000;
          logger.warn(`Server error (${response.status}). Retrying in ${waitMs / 1000}s...`);
          console.error(`Server error (${response.status}). Retrying in ${waitMs / 1000}s...`);
          await sleep(waitMs);
          continue;
        }
      }

      if (!response.ok) {
        const errorText = await response.text();
        let errorMessage: string;
        try {
          const errorJson = JSON.parse(errorText);
          errorMessage = errorJson.error?.message || errorText;
        } catch {
          errorMessage = errorText;
        }

        logger.error(`API error ${response.status}: ${errorMessage}`);

        // Retry once with a fresh token on 401
        if (response.status === 401 && !retriedWith401) {
          retriedWith401 = true;
          try {
            const freshToken = await this.auth.getToken(true);
            headers['Authorization'] = `Bearer ${freshToken}`;
            logger.info('Token refreshed, retrying request...');
            continue;
          } catch {
            throw new Error(`Authentication failed. Run: ms365 auth login`);
          }
        }
        if (response.status === 401) {
          throw new Error(`Authentication failed. Run: ms365 auth login`);
        }
        if (response.status === 403) {
          const hint = errorMessage.includes('scope') || errorMessage.includes('permission')
            ? ' Try running with --org-mode.'
            : '';
          throw new Error(`Permission denied (403): ${errorMessage}${hint}`);
        }
        throw new Error(`Graph API error (${response.status}): ${errorMessage}`);
      }

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

    throw lastError || new Error('Request failed after retries');
  }

  /** Fetch a full URL directly (used for pagination nextLinks). Retries on 429/5xx. */
  private async requestDirect(fullUrl: string): Promise<GraphResponse> {
    for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
      const token = await this.auth.getToken();
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 30_000);

      let response: Response;
      try {
        response = await fetch(fullUrl, {
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          signal: controller.signal,
        });
      } catch (err: unknown) {
        clearTimeout(timeout);
        if (err instanceof Error && err.name === 'AbortError') {
          if (attempt < MAX_RETRIES - 1) continue;
          throw new Error(`Pagination request timed out: ${fullUrl}`);
        }
        throw err;
      }
      clearTimeout(timeout);

      if (response.status === 429 || response.status >= 500) {
        if (attempt < MAX_RETRIES - 1) {
          const waitMs = parseRetryAfter(response.headers.get('Retry-After'), Math.pow(2, attempt) * 1000);
          await sleep(waitMs);
          continue;
        }
      }

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph API error (${response.status}): ${errorText}`);
      }

      const text = await response.text();
      if (!text) return { data: { message: 'OK' }, status: response.status };

      try {
        const json = JSON.parse(text);
        const nextLink = json['@odata.nextLink'] || undefined;
        removeODataProps(json);
        return { data: json, status: response.status, nextLink };
      } catch {
        return { data: { rawResponse: text }, status: response.status };
      }
    }

    throw new Error(`Pagination request failed after ${MAX_RETRIES} retries: ${fullUrl}`);
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
      const nextResponse = await this.requestDirect(nextLink);
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
