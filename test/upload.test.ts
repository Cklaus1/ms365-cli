import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { GraphClient, uploadSessionEndpoint } from '../src/graph-client.js';

describe('uploadSessionEndpoint (C-2)', () => {
  it.each([
    ['/drives/D/items/P:/big.zip:/content', '/drives/D/items/P:/big.zip:/createUploadSession'],
    ['/me/drive/root:/Docs/big.zip:/content', '/me/drive/root:/Docs/big.zip:/createUploadSession'],
    ['/me/photo/$value', null],
  ])('%s', (input, want) => { expect(uploadSessionEndpoint(input)).toBe(want); });
});

describe('uploadFile', () => {
  let originalFetch: typeof globalThis.fetch;
  let dir: string;
  const auth = { getToken: vi.fn().mockResolvedValue('tok') };
  const writeFile = (size: number) => {
    const p = path.join(dir, 'f.bin');
    fs.writeFileSync(p, Buffer.alloc(size, 7));
    return p;
  };
  beforeEach(() => {
    originalFetch = globalThis.fetch;
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-up-'));
    vi.spyOn(console, 'error').mockImplementation(() => {});
  });
  afterEach(() => { globalThis.fetch = originalFetch; vi.restoreAllMocks(); });

  it('large file: session URL drops /content, chunks resume from nextExpectedRanges after a 503', async () => {
    const file = writeFile(3000);
    const calls: Array<{ url: string; method: string; range?: string; auth?: string }> = [];
    let chunkPuts = 0;
    globalThis.fetch = vi.fn(async (url: string, init: RequestInit = {}) => {
      const h = (init.headers ?? {}) as Record<string, string>;
      calls.push({ url, method: init.method ?? 'GET', range: h['Content-Range'], auth: h.Authorization });
      if (url.endsWith('/createUploadSession')) {
        return new Response(JSON.stringify({ uploadUrl: 'https://up.example.net/s' }), { status: 200, headers: { 'content-type': 'application/json' } });
      }
      if (init.method === 'PUT') {
        chunkPuts++;
        if (chunkPuts === 1) return new Response(JSON.stringify({ nextExpectedRanges: ['1024-'] }), { status: 202 });
        if (chunkPuts === 2) return new Response('busy', { status: 503 });
        if (chunkPuts === 3) return new Response(JSON.stringify({ nextExpectedRanges: ['2048-'] }), { status: 202 });
        return new Response(JSON.stringify({ id: 'item1', size: 3000 }), { status: 201 });
      }
      // session status query after the failure
      return new Response(JSON.stringify({ nextExpectedRanges: ['1024-'] }), { status: 200 });
    }) as never;

    const client = new GraphClient(auth as never);
    const r = await client.uploadFile('/me/drive/root:/Docs/big.bin:/content', file, 'application/octet-stream',
      { maxSimpleSize: 1000, chunkSize: 1024 });

    expect(calls[0].url).toBe('https://graph.microsoft.com/v1.0/me/drive/root:/Docs/big.bin:/createUploadSession');
    const puts = calls.filter(c => c.method === 'PUT');
    expect(puts.map(p => p.range)).toEqual([
      'bytes 0-1023/3000', 'bytes 1024-2047/3000', 'bytes 1024-2047/3000', 'bytes 2048-2999/3000',
    ]);
    expect(puts.every(p => p.auth === undefined)).toBe(true); // never send the token to uploadUrl
    expect((r.data as { id: string }).id).toBe('item1');
  }, 15_000);

  it('dry-run of a large file shows the session URL', async () => {
    const file = writeFile(5 * 1024 * 1024);
    const out: string[] = [];
    vi.spyOn(console, 'error').mockImplementation((s: string) => { out.push(s); });
    const client = new GraphClient(auth as never, true);
    await client.uploadFile('/me/drive/root:/x/big.bin:/content', file, 'application/octet-stream');
    expect(out[0]).toBe('[dry-run] POST https://graph.microsoft.com/v1.0/me/drive/root:/x/big.bin:/createUploadSession');
  });

  it('rejects a large file for an endpoint without upload sessions', async () => {
    const file = writeFile(5 * 1024 * 1024);
    const client = new GraphClient(auth as never);
    await expect(client.uploadFile('/me/photo/$value', file, 'image/jpeg')).rejects.toThrow(/only accepts uploads up to/);
  });

  it('small file: simple PUT with its content type, retried on 503', async () => {
    const file = writeFile(10);
    let n = 0;
    const seen: string[] = [];
    globalThis.fetch = vi.fn(async (_url: string, init: RequestInit) => {
      seen.push((init.headers as Record<string, string>)['Content-Type']);
      if (++n === 1) return new Response('busy', { status: 503 });
      return new Response(JSON.stringify({ id: 'small' }), { status: 201, headers: { 'content-type': 'application/json' } });
    }) as never;
    const client = new GraphClient(auth as never);
    const r = await client.uploadFile('/me/drive/root:/a.txt:/content', file, 'text/plain');
    expect(n).toBe(2);
    expect(seen).toEqual(['text/plain', 'text/plain']);
    expect((r.data as { id: string }).id).toBe('small');
  }, 15_000);
});

// RT-1 (R-1, R-14): one session uploader for drive and Outlook, with both response conventions.
describe('uploadSession (shared session uploader)', () => {
  let originalFetch: typeof globalThis.fetch;
  let dir: string;
  const auth = { getToken: vi.fn().mockResolvedValue('tok') };
  const uploadUrl = 'https://up.example.net/session?authtoken=t';
  const sessionEndpoint = '/me/messages/M/attachments/createUploadSession';
  const sessionBody = { AttachmentItem: { attachmentType: 'file', name: 'f.bin', size: 0 } };
  const parseRange = (h: string) => {
    const m = /^bytes (\d+)-(\d+)\/(\d+)$/.exec(h)!;
    return { start: Number(m[1]), end: Number(m[2]) };
  };
  const writeRandom = (size: number) => {
    const p = path.join(dir, 'r.bin');
    fs.writeFileSync(p, Buffer.from(Array.from({ length: size }, (_, i) => (i * 131 + 7) % 251)));
    return p;
  };
  /** Advance fake timers until `p` settles (retries sleep, chunks time out). */
  const settle = async <T>(p: Promise<T>): Promise<T> => {
    let done = false;
    p.then(() => { done = true; }, () => { done = true; });
    for (let i = 0; i < 400 && !done; i++) await vi.advanceTimersByTimeAsync(1000);
    return p;
  };

  interface Call { method: string; url: string; range?: string; auth?: string }
  /**
   * A mock upload-session server. `style` picks the response convention: drive answers 202 + ranges and a final
   * 201 with the item; Outlook answers 200 + ranges and a final 201 with no body, and does not support GET.
   * `onPut(n)` can inject a failure for the n-th PUT: 'store-503' stores the bytes then answers 503 (response lost),
   * '503' and 'reset' drop the chunk.
   */
  const sessionServer = (size: number, style: 'drive' | 'outlook', onPut: (n: number) => string | undefined = () => undefined) => {
    const received = Buffer.alloc(size);
    let contiguous = 0;
    let puts = 0;
    const calls: Call[] = [];
    const store = (range: string, body: Buffer) => {
      const { start, end } = parseRange(range);
      expect(body.length).toBe(end - start + 1);
      body.copy(received, start);
      if (start <= contiguous) contiguous = Math.max(contiguous, end + 1);
    };
    const fetchMock = vi.fn(async (url: string, init: RequestInit = {}) => {
      const h = (init.headers ?? {}) as Record<string, string>;
      const method = init.method ?? 'GET';
      calls.push({ method, url, range: h['Content-Range'], auth: h.Authorization });
      if (url.endsWith('/createUploadSession')) {
        return new Response(JSON.stringify({ uploadUrl, nextExpectedRanges: ['0-'] }), { status: 200, headers: { 'content-type': 'application/json' } });
      }
      if (method === 'DELETE') return new Response(null, { status: 204 });
      if (method === 'GET') {
        if (style === 'outlook') return new Response('not supported', { status: 405 });
        return new Response(JSON.stringify({ nextExpectedRanges: [`${contiguous}-`] }), { status: 200 });
      }
      const fault = onPut(++puts);
      if (fault === 'reset') throw Object.assign(new TypeError('fetch failed'), { cause: { code: 'ECONNRESET' } });
      if (fault === '503') return new Response('busy', { status: 503 });
      store(h['Content-Range'], Buffer.from(init.body as Uint8Array));
      if (fault === 'store-503') return new Response('busy', { status: 503 });
      if (contiguous >= size) {
        return style === 'drive'
          ? new Response(JSON.stringify({ id: 'item1', size }), { status: 201 })
          : new Response(null, { status: 201 });
      }
      return new Response(JSON.stringify({ nextExpectedRanges: [`${contiguous}-`] }), { status: style === 'drive' ? 202 : 200 });
    });
    return { fetchMock, calls, received, puts: () => calls.filter(c => c.method === 'PUT').map(c => c.range) };
  };

  beforeEach(() => {
    originalFetch = globalThis.fetch;
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-us-'));
    vi.spyOn(console, 'error').mockImplementation(() => {});
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'Date'] });
  });
  afterEach(() => { vi.useRealTimers(); globalThis.fetch = originalFetch; vi.restoreAllMocks(); });

  // Critical-path gate: byte-exact round trip with a mid-upload failure and resume, for both conventions.
  it('drive style: a 503 after the server stored the chunk resumes from the session GET, byte-exact', async () => {
    const size = 5000;
    const file = writeRandom(size);
    const s = sessionServer(size, 'drive', n => (n === 3 ? 'store-503' : undefined));
    globalThis.fetch = s.fetchMock as never;
    const r = await settle(new GraphClient(auth as never).uploadSession(sessionEndpoint, sessionBody, file, size, { chunkSize: 1024 }));
    expect(s.received.equals(fs.readFileSync(file))).toBe(true);
    // chunk 3 (2048-3071) reached the server, so it is not sent again
    expect(s.puts()).toEqual(['bytes 0-1023/5000', 'bytes 1024-2047/5000', 'bytes 2048-3071/5000', 'bytes 3072-4095/5000', 'bytes 4096-4999/5000']);
    expect(s.calls.filter(c => c.method === 'GET')).toHaveLength(1);
    expect(s.calls.every(c => c.url === uploadUrl ? c.auth === undefined : true)).toBe(true);
    expect((r.data as { id: string }).id).toBe('item1');
    expect(s.calls.some(c => c.method === 'DELETE')).toBe(false);
  });

  it.each(['503', 'reset'])('Outlook style: %s on chunk 2 resumes from the last nextExpectedRanges, byte-exact', async (fault) => {
    const size = 3 * 1024 + 100;
    const file = writeRandom(size);
    const s = sessionServer(size, 'outlook', n => (n === 2 ? fault : undefined));
    globalThis.fetch = s.fetchMock as never;
    const r = await settle(new GraphClient(auth as never).uploadSession(sessionEndpoint, sessionBody, file, size, { chunkSize: 1024 }));
    expect(s.received.equals(fs.readFileSync(file))).toBe(true);
    expect(s.puts()).toEqual([
      'bytes 0-1023/3172', 'bytes 1024-2047/3172', 'bytes 1024-2047/3172', 'bytes 2048-3071/3172', 'bytes 3072-3171/3172',
    ]);
    // empty 201 body: success with a summary
    expect(r.data).toEqual({ message: 'Upload complete', size });
  });

  it('a 2xx without nextExpectedRanges advances one chunk; "{}" for every chunk sends each chunk once', async () => {
    const file = writeRandom(2049);
    const ranges: string[] = [];
    globalThis.fetch = vi.fn(async (url: string, init: RequestInit = {}) => {
      if (url.endsWith('/createUploadSession')) return new Response(JSON.stringify({ uploadUrl }), { status: 200, headers: { 'content-type': 'application/json' } });
      ranges.push((init.headers as Record<string, string>)['Content-Range']);
      return new Response('{}', { status: 200 });
    }) as never;
    const r = await settle(new GraphClient(auth as never).uploadSession(sessionEndpoint, sessionBody, file, 2049, { chunkSize: 1024 }));
    expect(ranges).toEqual(['bytes 0-1023/2049', 'bytes 1024-2047/2049', 'bytes 2048-2048/2049']);
    expect(r.data).toEqual({ message: 'Upload complete', size: 2049 });
  });

  it('a 201 ends the upload even before the last byte offset', async () => {
    const file = writeRandom(3000);
    let puts = 0;
    globalThis.fetch = vi.fn(async (url: string) => {
      if (url.endsWith('/createUploadSession')) return new Response(JSON.stringify({ uploadUrl }), { status: 200, headers: { 'content-type': 'application/json' } });
      return new Response(null, { status: ++puts === 1 ? 201 : 500 });
    }) as never;
    await settle(new GraphClient(auth as never).uploadSession(sessionEndpoint, sessionBody, file, 3000, { chunkSize: 1024 }));
    expect(puts).toBe(1);
  });

  it('posts the session body and sends chunks as application/octet-stream without the token', async () => {
    const file = writeRandom(10);
    const s = sessionServer(10, 'outlook');
    globalThis.fetch = s.fetchMock as never;
    await settle(new GraphClient(auth as never).uploadSession(sessionEndpoint, sessionBody, file, 10, { chunkSize: 1024 }));
    const [post, put] = s.fetchMock.mock.calls as unknown as Array<[string, RequestInit]>;
    expect(post[0]).toBe(`https://graph.microsoft.com/v1.0${sessionEndpoint}`);
    expect(JSON.parse(post[1].body as string)).toEqual(sessionBody);
    expect((put[1].headers as Record<string, string>)['Content-Type']).toBe('application/octet-stream');
    expect((put[1].headers as Record<string, string>).Authorization).toBeUndefined();
  });

  it('a chunk that stalls past the timeout aborts and is retried', async () => {
    const size = 2048;
    const file = writeRandom(size);
    const s = sessionServer(size, 'outlook');
    let stalled = false;
    globalThis.fetch = vi.fn(async (url: string, init: RequestInit = {}) => {
      if (init.method === 'PUT' && !stalled) {
        stalled = true;
        return new Promise((_, reject) => init.signal!.addEventListener('abort', () => reject(Object.assign(new Error('aborted'), { name: 'AbortError' }))));
      }
      return s.fetchMock(url, init);
    }) as never;
    await settle(new GraphClient(auth as never).uploadSession(sessionEndpoint, sessionBody, file, size, { chunkSize: 1024 }));
    expect(s.received.equals(fs.readFileSync(file))).toBe(true);
    expect(s.puts()).toEqual(['bytes 0-1023/2048', 'bytes 1024-2047/2048']);
  });

  it('a resume GET that stalls aborts after 30s and falls back to the last ranges (R-14)', async () => {
    const size = 2048;
    const file = writeRandom(size);
    const s = sessionServer(size, 'drive', n => (n === 2 ? '503' : undefined));
    let getAbortedAfter = -1;
    globalThis.fetch = vi.fn(async (url: string, init: RequestInit = {}) => {
      if ((init.method ?? 'GET') === 'GET' && url === uploadUrl) {
        const started = Date.now();
        expect(init.signal).toBeDefined();
        return new Promise((_, reject) => init.signal!.addEventListener('abort', () => {
          getAbortedAfter = Date.now() - started;
          reject(Object.assign(new Error('aborted'), { name: 'AbortError' }));
        }));
      }
      return s.fetchMock(url, init);
    }) as never;
    await settle(new GraphClient(auth as never).uploadSession(sessionEndpoint, sessionBody, file, size, { chunkSize: 1024 }));
    expect(getAbortedAfter).toBeGreaterThanOrEqual(30_000);
    expect(getAbortedAfter).toBeLessThan(31_000);
    expect(s.received.equals(fs.readFileSync(file))).toBe(true);
    expect(s.puts()).toEqual(['bytes 0-1023/2048', 'bytes 1024-2047/2048', 'bytes 1024-2047/2048']);
  });

  it('a non-retryable chunk error DELETEs the session and fails with upload_failed', async () => {
    const file = writeRandom(2048);
    const s = sessionServer(2048, 'outlook');
    globalThis.fetch = vi.fn(async (url: string, init: RequestInit = {}) => (
      init.method === 'PUT' ? new Response('bad', { status: 400 }) : s.fetchMock(url, init)
    )) as never;
    const err = await settle(new GraphClient(auth as never).uploadSession(sessionEndpoint, sessionBody, file, 2048, { chunkSize: 1024 })).catch(e => e);
    expect(err).toMatchObject({ code: 'upload_failed', status: 400 });
    expect(s.calls.filter(c => c.method === 'DELETE').map(c => [c.url, c.auth])).toEqual([[uploadUrl, undefined]]);
  });

  it('exhausted retries DELETE the session; a failing DELETE does not mask the upload error', async () => {
    const file = writeRandom(2048);
    let deletes = 0;
    globalThis.fetch = vi.fn(async (url: string, init: RequestInit = {}) => {
      if (url.endsWith('/createUploadSession')) return new Response(JSON.stringify({ uploadUrl }), { status: 200, headers: { 'content-type': 'application/json' } });
      if (init.method === 'DELETE') { deletes++; throw new TypeError('fetch failed'); }
      if (init.method === 'PUT') return new Response('busy', { status: 503 });
      return new Response('', { status: 404 });
    }) as never;
    const err = await settle(new GraphClient(auth as never).uploadSession(sessionEndpoint, sessionBody, file, 2048, { chunkSize: 1024 })).catch(e => e);
    expect(err).toMatchObject({ code: 'upload_failed', status: 503 });
    expect(deletes).toBe(1);
  });

  it('read-only blocks it; dry-run previews the POST and the chunk count without fetching', async () => {
    const file = writeRandom(2049);
    const f = vi.fn();
    globalThis.fetch = f as never;
    await expect(new GraphClient(auth as never, false, 'v1.0', true).uploadSession(sessionEndpoint, sessionBody, file, 2049))
      .rejects.toMatchObject({ code: 'read_only' });
    const out: string[] = [];
    vi.spyOn(console, 'error').mockImplementation((s: string) => { out.push(s); });
    const r = await new GraphClient(auth as never, true).uploadSession(sessionEndpoint, sessionBody, file, 2049, { chunkSize: 1024 });
    expect(out).toEqual([
      `[dry-run] POST https://graph.microsoft.com/v1.0${sessionEndpoint}`,
      `[dry-run] Body: ${JSON.stringify(sessionBody)}`,
      '[dry-run] Then PUT 3 chunk(s) to the session uploadUrl',
    ]);
    expect(r.status).toBe(0);
    expect(f).not.toHaveBeenCalled();
  });

  it('audit.log gets one PUT line (auditAs) for the whole upload', async () => {
    const prev = { dir: process.env.MS365_CLI_CONFIG_DIR, audit: process.env.MS365_CLI_AUDIT };
    process.env.MS365_CLI_CONFIG_DIR = path.join(dir, 'cfg');
    delete process.env.MS365_CLI_AUDIT;
    try {
      const file = writeRandom(3000);
      const s = sessionServer(3000, 'outlook');
      globalThis.fetch = s.fetchMock as never;
      await settle(new GraphClient(auth as never).uploadSession(sessionEndpoint, sessionBody, file, 3000, { chunkSize: 1024, auditAs: '/me/messages/M/attachments' }));
      const lines = fs.readFileSync(path.join(dir, 'cfg', 'audit.log'), 'utf8').trim().split('\n').map(l => JSON.parse(l));
      expect(lines).toEqual([expect.objectContaining({ method: 'PUT', endpoint: '/me/messages/M/attachments', ok: true })]);
    } finally {
      process.env.MS365_CLI_CONFIG_DIR = prev.dir;
      process.env.MS365_CLI_AUDIT = prev.audit;
    }
  });
});
