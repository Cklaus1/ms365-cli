import { describe, it, expect, afterEach, vi } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Writable } from 'node:stream';
import { GraphClient } from '../src/graph-client.js';
import { saveOrStream } from '../src/files.js';
import { mockFetch, sequence, graphError, fakeAuth, type Handler } from './helpers/fetch-mock.js';

let restore: (() => void) | undefined;
afterEach(() => {
  restore?.();
  restore = undefined;
  vi.restoreAllMocks();
});

function install(handler: Handler) {
  const m = mockFetch(handler);
  restore = m.restore;
  vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
  return m.calls;
}

/** PDF-ish bytes including invalid UTF-8, NUL and CR/LF so any text decoding would corrupt them. */
const PDF = Buffer.concat([Buffer.from('%PDF-1.7\r\n'), Buffer.from([0x00, 0xff, 0xfe, 0x80, 0x0d, 0x0a, 0xc3]), Buffer.alloc(70_000, 7)]);

function tmp(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-bin-'));
}

function collector(): { stream: Writable; bytes: () => Buffer } {
  const chunks: Buffer[] = [];
  const stream = new Writable({ write(chunk: Buffer, _enc, cb) { chunks.push(Buffer.from(chunk)); cb(); } });
  return { stream, bytes: () => Buffer.concat(chunks) };
}

describe('GRAPH-2 requestBinary + saveOrStream', () => {
  it('an application/pdf body round-trips byte-identical to --out', async () => {
    install(() => new Response(PDF, { status: 200, headers: { 'content-type': 'application/pdf' } }));
    const res = await new GraphClient(fakeAuth()).requestBinary('/me/drive/items/1/content');
    expect(res.contentType).toBe('application/pdf');
    const out = path.join(tmp(), 'report.pdf');
    const saved = await saveOrStream(res.body, { out });
    expect(saved).toEqual({ path: out, bytes: PDF.length });
    expect(fs.readFileSync(out).equals(PDF)).toBe(true);
  });

  it('a JSON file is returned verbatim, annotations included (never parsed)', async () => {
    const text = '{"@odata.context":"x","a":1}\n';
    install(() => new Response(text, { status: 200, headers: { 'content-type': 'application/json' } }));
    const res = await new GraphClient(fakeAuth()).requestBinary('/me/drive/items/1/content');
    const sink = collector();
    const saved = await saveOrStream(res.body, {}, sink.stream);
    expect(sink.bytes().toString('utf8')).toBe(text);
    expect(saved).toEqual({ bytes: Buffer.byteLength(text) });
  });

  it('streams to stdout when no --out is given', async () => {
    install(() => new Response(PDF, { status: 200, headers: { 'content-type': 'application/pdf' } }));
    const res = await new GraphClient(fakeAuth()).requestBinary('/x');
    const sink = collector();
    await saveOrStream(res.body, {}, sink.stream);
    expect(sink.bytes().equals(PDF)).toBe(true);
  });

  it('refuses to overwrite an existing file without force, and leaves it intact', async () => {
    const out = path.join(tmp(), 'keep.txt');
    fs.writeFileSync(out, 'original');
    await expect(saveOrStream(Buffer.from('new'), { out })).rejects.toThrow(/exists.*--force/);
    expect(fs.readFileSync(out, 'utf8')).toBe('original');
    expect(fs.readdirSync(path.dirname(out))).toEqual(['keep.txt']);
  });

  it('overwrites with force', async () => {
    const out = path.join(tmp(), 'keep.txt');
    fs.writeFileSync(out, 'original');
    await saveOrStream(Buffer.from('new'), { out, force: true });
    expect(fs.readFileSync(out, 'utf8')).toBe('new');
  });

  it('a stream that fails midway leaves no partial file and keeps the existing one', async () => {
    const dir = tmp();
    const out = path.join(dir, 'f.bin');
    fs.writeFileSync(out, 'original');
    const failing = new ReadableStream<Uint8Array>({
      start(c) { c.enqueue(new Uint8Array([1, 2, 3])); c.error(new Error('connection reset')); },
    });
    await expect(saveOrStream(failing, { out, force: true })).rejects.toThrow('connection reset');
    expect(fs.readFileSync(out, 'utf8')).toBe('original');
    expect(fs.readdirSync(dir)).toEqual(['f.bin']);
  });

  it('dry-run: no fetch, preview on stderr, nothing written', async () => {
    const calls = install(() => new Response('x'));
    const res = await new GraphClient(fakeAuth(), true).requestBinary('/me/photo/$value');
    expect(calls).toHaveLength(0);
    expect(res.status).toBe(0);
    expect(res.body).toBeNull();
    const out = path.join(tmp(), 'p.jpg');
    expect(await saveOrStream(res.body, { out })).toEqual({ bytes: 0 });
    expect(fs.existsSync(out)).toBe(false);
  });

  it('follows the GET retry policy (503 retried)', async () => {
    const calls = install(sequence(() => graphError(503, 'busy', { 'Retry-After': '0' }), () => new Response(PDF, { headers: { 'content-type': 'application/pdf' } })));
    const res = await new GraphClient(fakeAuth()).requestBinary('/x');
    const sink = collector();
    await saveOrStream(res.body, {}, sink.stream);
    expect(calls).toHaveLength(2);
    expect(sink.bytes().equals(PDF)).toBe(true);
  });

  it('is allowed under --read-only', async () => {
    const calls = install(() => new Response(PDF, { headers: { 'content-type': 'application/pdf' } }));
    await new GraphClient(fakeAuth(), false, 'v1.0', true).requestBinary('/x');
    expect(calls).toHaveLength(1);
  });
});

describe('saveOrStream on filesystems without hard links (vfat/exFAT/FUSE)', () => {
  const noLink = (code: string) => Object.assign(new Error(`${code}: link not supported`), { code });

  it.each(['EPERM', 'ENOTSUP', 'ENOSYS'])('saves without --force when link fails with %s, leaving no temp file', async (code) => {
    vi.spyOn(fs, 'linkSync').mockImplementation(() => { throw noLink(code); });
    const dir = tmp();
    const out = path.join(dir, 'a.pdf');
    expect(await saveOrStream(PDF, { out })).toEqual({ path: out, bytes: PDF.byteLength });
    expect(fs.readFileSync(out).equals(PDF)).toBe(true);
    expect(fs.readdirSync(dir)).toEqual(['a.pdf']);
  });

  it('still refuses to overwrite a file that appeared during the download', async () => {
    const dir = tmp();
    const out = path.join(dir, 'a.txt');
    vi.spyOn(fs, 'linkSync').mockImplementation(() => {
      fs.writeFileSync(out, 'racer');
      throw noLink('EPERM');
    });
    await expect(saveOrStream(Buffer.from('new'), { out })).rejects.toThrow(/File exists.*--force/);
    expect(fs.readFileSync(out, 'utf8')).toBe('racer');
    expect(fs.readdirSync(dir)).toEqual(['a.txt']);
  });

  it('other link errors still fail the save', async () => {
    vi.spyOn(fs, 'linkSync').mockImplementation(() => { throw noLink('EIO'); });
    const dir = tmp();
    await expect(saveOrStream(Buffer.from('x'), { out: path.join(dir, 'a.txt') })).rejects.toThrow('EIO');
    expect(fs.readdirSync(dir)).toEqual([]);
  });
});
