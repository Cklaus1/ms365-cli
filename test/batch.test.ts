import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { GraphClient } from '../src/graph-client.js';
import { runCommand } from './helpers/cli.js';
import { registerMailCommands } from '../src/commands/mail.js';

type Sub = { id: string; method: string; url: string; body?: unknown };
const answer = (fn: (q: Sub) => { status: number; body?: unknown; headers?: Record<string, string> }) =>
  (c: { body?: unknown }) => ({ responses: (c.body as { requests: Sub[] }).requests.map(q => ({ id: q.id, ...fn(q) })) });

describe('GraphClient.batch (P-5, G-10, D19)', () => {
  let originalFetch: typeof globalThis.fetch;
  const auth = { getToken: vi.fn().mockResolvedValue('tok') };
  beforeEach(() => { originalFetch = globalThis.fetch; vi.spyOn(console, 'error').mockImplementation(() => {}); });
  afterEach(() => { globalThis.fetch = originalFetch; vi.restoreAllMocks(); });

  function mockBatch(handler: (q: Sub) => { status: number; body?: unknown; headers?: Record<string, string> }) {
    const batches: Sub[][] = [];
    globalThis.fetch = vi.fn(async (_u: string, init: RequestInit) => {
      const reqs = JSON.parse(init.body as string).requests as Sub[];
      batches.push(reqs);
      return new Response(JSON.stringify(answer(handler)({ body: { requests: reqs } })), { status: 200, headers: { 'content-type': 'application/json' } });
    }) as never;
    return batches;
  }

  it('chunks into batches of 20 and returns results in input order', async () => {
    const batches = mockBatch(q => ({ status: 200, body: { n: Number(q.id) } }));
    const reqs = Array.from({ length: 45 }, (_, i) => ({ method: 'GET', url: `/me/messages/${i}` }));
    const res = await new GraphClient(auth as never).batch(reqs);
    expect(batches.map(b => b.length)).toEqual([20, 20, 5]);
    expect(res.map(r => (r.body as { n: number }).n)).toEqual(Array.from({ length: 45 }, (_, i) => i));
  });

  it('retries 429 items in a new batch; never retries a 5xx POST; retries a 5xx PATCH', async () => {
    const seen: Record<string, number> = {};
    const batches = mockBatch(q => {
      seen[q.url] = (seen[q.url] ?? 0) + 1;
      if (q.url === '/throttled' && seen[q.url] === 1) return { status: 429, headers: { 'Retry-After': '0' } };
      if (q.url === '/post500' || (q.url === '/patch500' && seen[q.url] === 1)) return { status: 503 };
      return { status: 200, body: {} };
    });
    const res = await new GraphClient(auth as never).batch([
      { method: 'GET', url: '/throttled' },
      { method: 'POST', url: '/post500', body: {} },
      { method: 'PATCH', url: '/patch500', body: {} },
    ]);
    expect(res.map(r => r.status)).toEqual([200, 503, 200]);
    expect(seen).toEqual({ '/throttled': 2, '/post500': 1, '/patch500': 2 });
    expect(batches).toHaveLength(2);
    expect(batches[0][1].body).toEqual({});
  }, 15_000);

  it('read-only blocks batches with writes; dry-run sends nothing', async () => {
    await expect(new GraphClient(auth as never, false, 'v1.0', true).batch([{ method: 'PATCH', url: '/x', body: {} }])).rejects.toThrow(/Read-only/);
    const f = vi.fn(); globalThis.fetch = f as never;
    const log = vi.spyOn(console, 'error').mockImplementation(() => {});
    const r = await new GraphClient(auth as never, true).batch([{ method: 'GET', url: '/a' }]);
    expect(f).not.toHaveBeenCalled();
    expect(r[0].status).toBe(0);
    expect(log.mock.calls.flat().join('\n')).toContain('$batch');
  });
});

describe('mail batch-read / batch-move via $batch', () => {
  it('marks 100 messages read in ≤ 6 HTTP requests (spec acceptance criterion)', async () => {
    const ids = Array.from({ length: 100 }, (_, i) => `M${i}`).join(',');
    const r = await runCommand(registerMailCommands, ['mail', 'batch-read', ids], {
      routes: [['POST', '/$batch', answer(() => ({ status: 200, body: {} }))]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls.length).toBeLessThanOrEqual(6);
    expect(r.calls.every(c => c.path === '/$batch')).toBe(true);
    expect(r.stderr).toContain('Marked 100 message(s) as read.');
  });

  it('partial failure exits 1 and names the failed IDs with Graph messages', async () => {
    const r = await runCommand(registerMailCommands, ['mail', 'batch-move', 'GOOD,BAD', '--folder', 'AAMkFolderId'], {
      routes: [
        ['GET', '/me/mailFolders', () => ({ value: [] })],
        ['POST', '/$batch', answer(q => q.url.includes('BAD') ? { status: 404, body: { error: { message: 'Item not found' } } } : { status: 201, body: {} })],
      ],
    });
    expect(r.exitCode).toBe(1);
    expect(r.stderr).toContain('Moved 1 message(s).');
    expect(r.stderr).toMatch(/BAD\.\.\.: Item not found/);
  });
});
