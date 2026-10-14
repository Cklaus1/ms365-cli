// DR-T1 (DR-3, DR-4, DR-6, DR-7, DR-8): the dry-run preview path of GraphClient.
import { describe, it, expect, vi, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { GraphClient, redactForPreview } from '../src/graph-client.js';
import { previewCount, resetPreviewCount } from '../src/utils.js';

const auth = { getToken: vi.fn(async () => 't') };

function capture() {
  const out: string[] = [];
  const err: string[] = [];
  vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => { out.push(a.map(String).join(' ')); });
  vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => { err.push(a.map(String).join(' ')); });
  return { out, err };
}

afterEach(() => { vi.restoreAllMocks(); resetPreviewCount(); });

describe('redactForPreview', () => {
  it('redacts credential-like keys at any depth, case-insensitively', () => {
    const r = redactForPreview({
      password: 's3cret', a: { Client_Secret: 'x', nested: [{ accessToken: 'y', refresh_token: 'z' }] },
      Authorization: 'Bearer abc', cookie: 'c=1', subject: 'keep', tokenType: 'keep-me',
    });
    expect(r).toEqual({
      password: '<redacted>', a: { Client_Secret: '<redacted>', nested: [{ accessToken: '<redacted>', refresh_token: '<redacted>' }] },
      Authorization: '<redacted>', cookie: '<redacted>', subject: 'keep', tokenType: 'keep-me',
    });
  });

  it('truncates contentBytes to its decoded size', () => {
    const b64 = Buffer.alloc(3000, 1).toString('base64');
    expect(redactForPreview({ attachments: [{ name: 'a.bin', contentBytes: b64 }] }))
      .toEqual({ attachments: [{ name: 'a.bin', contentBytes: '<base64, 3000 bytes>' }] });
  });

  it('leaves non-objects unchanged and does not mutate the input', () => {
    const input = { password: 'p' };
    redactForPreview(input);
    expect(input.password).toBe('p');
    expect(redactForPreview('plain')).toBe('plain');
    expect(redactForPreview(null)).toBe(null);
  });
});

describe('dry-run request preview', () => {
  it('writes the trace to stderr only; the placeholder carries the parsed, redacted body', async () => {
    const { out, err } = capture();
    const client = new GraphClient(auth as never, true);
    const res = await client.request('/shares', {
      method: 'POST',
      headers: { 'If-Match': 'e1', Authorization: 'Bearer leak' },
      body: JSON.stringify({ type: 'view', password: 's3cret' }),
    });
    expect(out).toEqual([]);
    expect(err).toEqual([
      '[dry-run] POST https://graph.microsoft.com/v1.0/shares',
      '[dry-run] Body: {"type":"view","password":"<redacted>"}',
      '[dry-run] Headers: {"If-Match":"e1","Authorization":"<redacted>"}',
    ]);
    expect(res).toEqual({ status: 0, data: { dryRun: true, method: 'POST', url: 'https://graph.microsoft.com/v1.0/shares', body: { type: 'view', password: '<redacted>' } } });
    expect(err.join('\n')).not.toContain('s3cret');
    expect(auth.getToken).not.toHaveBeenCalled();
  });

  it('prints a non-JSON string body unchanged and a binary body as its size', async () => {
    const { err } = capture();
    const client = new GraphClient(auth as never, true);
    const a = await client.request('/x', { method: 'PUT', body: 'raw text' });
    const b = await client.request('/y', { method: 'PUT', body: new Uint8Array(7) });
    expect(err).toContain('[dry-run] Body: raw text');
    expect(err).toContain('[dry-run] Body: <7 bytes>');
    expect((a.data as { body: unknown }).body).toBe('raw text');
    expect((b.data as { body: unknown }).body).toBe('<7 bytes>');
  });

  it('counts previewed requests until reset', async () => {
    capture();
    const client = new GraphClient(auth as never, true);
    resetPreviewCount();
    await client.request('/a');
    await client.request('/b', { method: 'DELETE' });
    await client.batch([{ method: 'GET', url: '/c' }, { method: 'PATCH', url: '/d', body: { secret: 'x' } }]);
    expect(previewCount()).toBe(4);
    resetPreviewCount();
    expect(previewCount()).toBe(0);
  });

  it('batch: trace on stderr with redacted bodies', async () => {
    const { out, err } = capture();
    const client = new GraphClient(auth as never, true);
    await client.batch([{ method: 'patch', url: '/d', body: { clientSecret: 'x', n: 1 } }]);
    expect(out).toEqual([]);
    expect(err[0]).toBe('[dry-run] PATCH https://graph.microsoft.com/v1.0/d');
    expect(err[1]).toBe('[dry-run] Body: {"clientSecret":"<redacted>","n":1}');
  });

  it('isPreview is true only for dry-run placeholders', async () => {
    capture();
    const dry = new GraphClient(auth as never, true);
    const live = new GraphClient(auth as never);
    expect(dry.isPreview(await dry.request('/me'))).toBe(true);
    expect(dry.isPreview({ status: 200 })).toBe(false);
    expect(live.isPreview({ status: 0 })).toBe(false); // e.g. an unanswered batch item
  });

  it('paged reads yield no items under dry-run (no placeholder rows)', async () => {
    const { err } = capture();
    const client = new GraphClient(auth as never, true);
    const pages: unknown[][] = [];
    for await (const p of client.requestPagesStreaming('/me/messages')) pages.push(p);
    expect(pages).toEqual([[]]);
    expect((await client.requestAllPages('/me/events')).data).toEqual({ value: [] });
    expect(err.filter(l => / GET /.test(l))).toHaveLength(2);
  });

  it('uploadFile and uploadSession trace to stderr and count one request', async () => {
    const { out, err } = capture();
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-prev-'));
    const file = path.join(dir, 'f.bin');
    fs.writeFileSync(file, Buffer.alloc(10));
    const client = new GraphClient(auth as never, true);
    resetPreviewCount();
    await client.uploadFile('/me/drive/root:/f.bin:/content', file, 'application/octet-stream');
    await client.uploadSession('/me/messages/M/attachments/createUploadSession', { AttachmentItem: { name: 'f', token: 't' } }, file, 10);
    expect(out).toEqual([]);
    expect(err).toContain('[dry-run] PUT https://graph.microsoft.com/v1.0/me/drive/root:/f.bin:/content');
    expect(err).toContain('[dry-run] Body: {"AttachmentItem":{"name":"f","token":"<redacted>"}}');
    expect(previewCount()).toBe(2);
    fs.rmSync(dir, { recursive: true, force: true });
  });
});

// DR-T2 (DR-16, F23; DR-15 half): auth is a lazy thunk, never built under dry-run; the client ID is passed in.
describe('lazy auth and clientId', () => {
  it('never resolves the auth thunk under dry-run', async () => {
    capture();
    const thunk = vi.fn(() => { throw new Error('no client id'); });
    const client = new GraphClient(thunk as never, true);
    await client.request('/me/sendMail', { method: 'POST', body: '{}' });
    await client.batch([{ method: 'DELETE', url: '/x' }]);
    await client.uploadSession('/s', {}, '/nonexistent', 10);
    expect(thunk).not.toHaveBeenCalled();
  });

  it('resolves the thunk once, on the first real request', async () => {
    const orig = globalThis.fetch;
    globalThis.fetch = vi.fn(async () => new Response('{}', { status: 200, headers: { 'content-type': 'application/json' } })) as never;
    try {
      const thunk = vi.fn(() => auth);
      const client = new GraphClient(thunk as never);
      expect(thunk).not.toHaveBeenCalled();
      await client.request('/me');
      await client.request('/me');
      expect(thunk).toHaveBeenCalledTimes(1);
    } finally {
      globalThis.fetch = orig;
    }
  });

  it('exposes the resolved client ID without building auth', () => {
    const thunk = vi.fn();
    expect(new GraphClient(thunk as never, true, 'v1.0', false, { clientId: 'abc' }).clientId()).toBe('abc');
    expect(new GraphClient(thunk as never, true).clientId()).toBeUndefined();
    expect(thunk).not.toHaveBeenCalled();
  });
});
