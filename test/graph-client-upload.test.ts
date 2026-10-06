import { describe, it, expect, afterEach, vi } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { GraphClient } from '../src/graph-client.js';
import { setQuiet } from '../src/utils.js';
import { mockFetch, jsonResponse, graphError, abortError, fakeAuth, type Handler, type RecordedCall } from './helpers/fetch-mock.js';

const BASE = 'https://graph.microsoft.com/v1.0';
const UPLOAD_URL = 'https://upload.example.com/session/abc';

let restore: (() => void) | undefined;
let stderr: string[] = [];
afterEach(() => {
  restore?.();
  restore = undefined;
  setQuiet(false);
  vi.restoreAllMocks();
});

function install(handler: Handler): RecordedCall[] {
  const m = mockFetch(handler);
  restore = m.restore;
  stderr = [];
  vi.spyOn(process.stderr, 'write').mockImplementation((chunk: string | Uint8Array) => {
    stderr.push(String(chunk));
    return true;
  });
  vi.spyOn(console, 'error').mockImplementation((...args: unknown[]) => { stderr.push(args.join(' ') + '\n'); });
  return m.calls;
}

function file(bytes: Buffer, name = 'report.docx'): string {
  const p = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-up-')), name);
  fs.writeFileSync(p, bytes);
  return p;
}

const SMALL = Buffer.from('hello\x00\xff world');
const LARGE = Buffer.from(Array.from({ length: 10 }, (_, i) => i));

/** Fake upload-session server: accepts Content-Range chunks, answers 202/201 like OneDrive. */
function sessionServer(opts: { failChunkAt?: Set<number>; storedName?: string } = {}): { handler: Handler; received: Buffer } {
  const state = { received: Buffer.alloc(0) };
  let chunkCalls = 0;
  const handler: Handler = (call) => {
    if (call.url.endsWith(':/createUploadSession')) return jsonResponse({ uploadUrl: UPLOAD_URL });
    if (call.url === UPLOAD_URL && call.method === 'GET') {
      return jsonResponse({ nextExpectedRanges: [`${state.received.length}-`] });
    }
    if (call.url === UPLOAD_URL && call.method === 'DELETE') return new Response(null, { status: 204 });
    if (call.url === UPLOAD_URL && call.method === 'PUT') {
      const index = chunkCalls++;
      const m = /bytes (\d+)-(\d+)\/(\d+)/.exec(call.headers['content-range'] ?? '');
      if (!m) return graphError(400, 'no range');
      const [start, , total] = [Number(m[1]), Number(m[2]), Number(m[3])];
      if (opts.failChunkAt?.has(index)) {
        // the server kept half the chunk before the connection broke
        const body = call.body as Buffer;
        state.received = Buffer.concat([state.received.subarray(0, start), body.subarray(0, Math.ceil(body.length / 2))]);
        throw new TypeError('fetch failed');
      }
      if (start !== state.received.length) return graphError(416, 'range not expected');
      state.received = Buffer.concat([state.received, call.body as Buffer]);
      if (state.received.length === total) return jsonResponse({ id: 'i1', name: opts.storedName ?? 'big.bin', size: total }, 201);
      return jsonResponse({ nextExpectedRanges: [`${state.received.length}-`] }, 202);
    }
    return graphError(404, `unexpected ${call.method} ${call.url}`);
  };
  return { handler, get received() { return state.received; } };
}

describe('GRAPH-3 uploadDriveItem', () => {
  it('simple upload appends :/content and sends the conflict behavior in the URL', async () => {
    const calls = install(() => jsonResponse({ id: 'i1', name: 'report.docx' }, 201));
    const local = file(SMALL);
    const res = await new GraphClient(fakeAuth()).uploadDriveItem('/me/drive/root:/Docs/report.docx', local, { conflict: 'rename' });
    expect(calls).toHaveLength(1);
    expect(calls[0].method).toBe('PUT');
    expect(calls[0].url).toBe(`${BASE}/me/drive/root:/Docs/report.docx:/content?@microsoft.graph.conflictBehavior=rename`);
    expect(Buffer.from(calls[0].body as Buffer).equals(SMALL)).toBe(true);
    expect(calls[0].headers['content-type']).toBe('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    expect(res.data).toMatchObject({ name: 'report.docx' });
  });

  it.each(['replace', 'fail'] as const)('conflict %s reaches the simple PUT URL', async (conflict) => {
    const calls = install(() => jsonResponse({ id: 'i1', name: 'report.docx' }, 200));
    await new GraphClient(fakeAuth()).uploadDriveItem('/me/drive/root:/report.docx', file(SMALL), { conflict });
    expect(calls[0].url).toContain(`conflictBehavior=${conflict}`);
  });

  it('reports the name Graph stored when it differs from the requested one', async () => {
    install(() => jsonResponse({ id: 'i1', name: 'report 1.docx' }, 201));
    const res = await new GraphClient(fakeAuth()).uploadDriveItem('/me/drive/root:/report.docx', file(SMALL), { conflict: 'rename' });
    expect(res.data).toMatchObject({ name: 'report 1.docx' });
    const err = stderr.join('');
    expect(err).toContain('report 1.docx');
    expect(err).toContain('report.docx');
  });

  it('large file: session URL ends :/createUploadSession, conflict in the item body, chunks reassemble', async () => {
    const server = sessionServer();
    const calls = install(server.handler);
    const res = await new GraphClient(fakeAuth()).uploadDriveItem('/drives/d1/items/p1:/big.bin', file(LARGE, 'big.bin'), {
      conflict: 'replace', maxSimpleSize: 4, chunkSize: 4,
    });
    expect(calls[0].method).toBe('POST');
    expect(calls[0].url).toBe(`${BASE}/drives/d1/items/p1:/big.bin:/createUploadSession`);
    expect(JSON.parse(String(calls[0].body))).toEqual({ item: { '@microsoft.graph.conflictBehavior': 'replace' } });
    const puts = calls.filter((c) => c.method === 'PUT');
    expect(puts.map((c) => c.headers['content-range'])).toEqual(['bytes 0-3/10', 'bytes 4-7/10', 'bytes 8-9/10']);
    // the pre-authenticated upload URL must not receive the bearer token
    expect(puts.every((c) => c.headers.authorization === undefined)).toBe(true);
    expect(server.received.equals(LARGE)).toBe(true);
    expect(res.data).toMatchObject({ name: 'big.bin', size: 10 });
  });

  it('a broken chunk resumes from the server nextExpectedRanges', async () => {
    const server = sessionServer({ failChunkAt: new Set([1]) });
    const calls = install(server.handler);
    await new GraphClient(fakeAuth()).uploadDriveItem('/me/drive/root:/big.bin', file(LARGE, 'big.bin'), {
      conflict: 'replace', maxSimpleSize: 4, chunkSize: 4,
    });
    expect(server.received.equals(LARGE)).toBe(true);
    expect(calls.some((c) => c.url === UPLOAD_URL && c.method === 'GET')).toBe(true);
    expect(calls.some((c) => c.method === 'DELETE')).toBe(false);
  });

  it('a session that keeps failing is cancelled (DELETE uploadUrl) and the error surfaces', async () => {
    const server = sessionServer({ failChunkAt: new Set(Array.from({ length: 50 }, (_, i) => i)) });
    const calls = install(server.handler);
    await expect(new GraphClient(fakeAuth()).uploadDriveItem('/me/drive/root:/big.bin', file(LARGE, 'big.bin'), {
      conflict: 'replace', maxSimpleSize: 4, chunkSize: 4,
    })).rejects.toThrow('fetch failed');
    const del = calls.filter((c) => c.method === 'DELETE');
    expect(del).toHaveLength(1);
    expect(del[0].url).toBe(UPLOAD_URL);
  });

  it('a rename upload that times out is not retried (a second copy could be created)', async () => {
    const calls = install(() => { throw abortError(); });
    await expect(new GraphClient(fakeAuth()).uploadDriveItem('/me/drive/root:/a.txt', file(SMALL, 'a.txt'), { conflict: 'rename' }))
      .rejects.toThrow(/may have been applied/);
    expect(calls).toHaveLength(1);
  });

  it('a replace upload is retried on 503', async () => {
    let n = 0;
    const calls = install(() => (n++ === 0 ? graphError(503, 'busy', { 'Retry-After': '0' }) : jsonResponse({ name: 'a.txt' })));
    await new GraphClient(fakeAuth()).uploadDriveItem('/me/drive/root:/a.txt', file(SMALL, 'a.txt'), { conflict: 'replace' });
    expect(calls).toHaveLength(2);
  });

  it('progress goes through info(): silent under --quiet', async () => {
    const server = sessionServer();
    install(server.handler);
    setQuiet(true);
    await new GraphClient(fakeAuth()).uploadDriveItem('/me/drive/root:/big.bin', file(LARGE, 'big.bin'), {
      conflict: 'replace', maxSimpleSize: 4, chunkSize: 4,
    });
    expect(stderr.join('')).toBe('');
  });

  it('dry-run: no request, preview of the content URL on stderr', async () => {
    const calls = install(() => jsonResponse({}));
    const res = await new GraphClient(fakeAuth(), true).uploadDriveItem('/me/drive/root:/a.txt', file(SMALL, 'a.txt'), { conflict: 'fail' });
    expect(calls).toHaveLength(0);
    expect(res.status).toBe(0);
    expect(stderr.join('')).toContain(`[dry-run] PUT ${BASE}/me/drive/root:/a.txt:/content?@microsoft.graph.conflictBehavior=fail`);
  });

  it('is blocked under --read-only before touching the network', async () => {
    const calls = install(() => jsonResponse({}));
    await expect(new GraphClient(fakeAuth(), false, 'v1.0', true).uploadDriveItem('/me/drive/root:/a.txt', file(SMALL, 'a.txt'), { conflict: 'fail' }))
      .rejects.toThrow('Read-only mode');
    expect(calls).toHaveLength(0);
  });
});

describe('GRAPH-3 putBinary', () => {
  it('PUTs the raw file to exactly /me/photo/$value', async () => {
    const calls = install(() => new Response(null, { status: 200 }));
    const local = file(SMALL, 'me.png');
    const res = await new GraphClient(fakeAuth()).putBinary('/me/photo/$value', local, 'image/png');
    expect(calls).toHaveLength(1);
    expect(calls[0].url).toBe(`${BASE}/me/photo/$value`);
    expect(calls[0].method).toBe('PUT');
    expect(calls[0].headers['content-type']).toBe('image/png');
    expect(Buffer.from(calls[0].body as Buffer).equals(SMALL)).toBe(true);
    expect(res.status).toBe(200);
  });

  it('surfaces a Graph error instead of resolving', async () => {
    install(() => graphError(400, 'Bad image'));
    await expect(new GraphClient(fakeAuth()).putBinary('/me/photo/$value', file(SMALL, 'me.png'), 'image/png')).rejects.toThrow('Bad image');
  });

  it('dry-run: no request', async () => {
    const calls = install(() => jsonResponse({}));
    const res = await new GraphClient(fakeAuth(), true).putBinary('/me/photo/$value', file(SMALL, 'me.png'), 'image/png');
    expect(calls).toHaveLength(0);
    expect(res.status).toBe(0);
  });
});
