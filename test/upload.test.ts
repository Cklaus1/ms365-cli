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
    vi.spyOn(console, 'log').mockImplementation((s: string) => { out.push(s); });
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
