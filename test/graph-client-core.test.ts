import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { GraphClient, isTrustedHost } from '../src/graph-client.js';
import { CliError, ExitCode } from '../src/errors.js';

const json = (body: unknown, status = 200, headers: Record<string, string> = {}) =>
  new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json', ...headers } });

describe('GraphClient core (T03)', () => {
  let originalFetch: typeof globalThis.fetch;
  let auth: { getToken: ReturnType<typeof vi.fn> };
  beforeEach(() => {
    originalFetch = globalThis.fetch;
    auth = { getToken: vi.fn().mockResolvedValue('tok') };
    vi.spyOn(console, 'error').mockImplementation(() => {});
  });
  afterEach(() => { globalThis.fetch = originalFetch; vi.restoreAllMocks(); });
  const client = () => new GraphClient(auth as never);

  // Critical-path invariant: every non-JSON body round-trips byte-for-byte.
  describe('binary round-trip (C-3)', () => {
    const bytes = Uint8Array.from({ length: 4096 }, (_, i) => (i * 37 + 11) % 256); // includes invalid UTF-8
    it.each([
      'application/pdf',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/zip',
      'application/gzip',
      'image/png',
      'video/mp4',
      'application/octet-stream',
    ])('auto mode: %s', async (ct) => {
      globalThis.fetch = vi.fn(async () => new Response(bytes, { status: 200, headers: { 'content-type': ct } })) as never;
      const r = await client().request('/me/drive/items/x/content');
      const d = r.data as { contentBytes: string; size: number };
      expect(Buffer.from(d.contentBytes, 'base64').equals(Buffer.from(bytes))).toBe(true);
      expect(d.size).toBe(bytes.length);
    });

    it('stream mode leaves the body unread', async () => {
      globalThis.fetch = vi.fn(async () => new Response(bytes, { status: 200, headers: { 'content-type': 'application/pdf', 'content-length': String(bytes.length) } })) as never;
      const r = await client().request('/x/content', { responseType: 'stream' });
      const buf = Buffer.from(await new Response(r.stream!).arrayBuffer());
      expect(buf.equals(Buffer.from(bytes))).toBe(true);
      expect((r.data as { size: number }).size).toBe(bytes.length);
    });

    it('text mode returns the raw string even for JSON-looking files', async () => {
      globalThis.fetch = vi.fn(async () => new Response('{"@odata.type":"keep"}', { status: 200, headers: { 'content-type': 'application/json' } })) as never;
      const r = await client().request('/x/content', { responseType: 'text' });
      expect(r.data).toBe('{"@odata.type":"keep"}');
    });

    it('auto mode keeps text/html as rawResponse', async () => {
      globalThis.fetch = vi.fn(async () => new Response('<html>hi</html>', { status: 200, headers: { 'content-type': 'text/html' } })) as never;
      const r = await client().request('/me/onenote/pages/1/content');
      expect((r.data as { rawResponse: string }).rawResponse).toBe('<html>hi</html>');
    });
  });

  describe('retry policy (C-7)', () => {
    const abortErr = () => Object.assign(new Error('aborted'), { name: 'AbortError' });

    it('never retries a POST after a timeout', async () => {
      const f = vi.fn(async () => { throw abortErr(); });
      globalThis.fetch = f as never;
      await expect(client().request('/me/sendMail', { method: 'POST', body: '{}' })).rejects.toThrow(/not retried/);
      expect(f).toHaveBeenCalledTimes(1);
    });

    it('never retries a POST after a 5xx', async () => {
      const f = vi.fn(async () => new Response('boom', { status: 503 }));
      globalThis.fetch = f as never;
      await expect(client().request('/me/sendMail', { method: 'POST', body: '{}' })).rejects.toThrow(/503/);
      expect(f).toHaveBeenCalledTimes(1);
    });

    it('retries a POST on 429', async () => {
      let n = 0;
      globalThis.fetch = vi.fn(async () => (++n === 1 ? new Response('', { status: 429, headers: { 'Retry-After': '0' } }) : json({ ok: 1 }))) as never;
      const r = await client().request('/me/sendMail', { method: 'POST', body: '{}' });
      expect(n).toBe(2);
      expect(r.status).toBe(200);
    });

    it('retries a GET after a timeout', async () => {
      let n = 0;
      globalThis.fetch = vi.fn(async () => { if (++n === 1) throw abortErr(); return json({ id: 1 }); }) as never;
      const r = await client().request('/me');
      expect(n).toBe(2);
      expect(r.status).toBe(200);
    }, 10_000);

    it('retries a POST on a pre-send connection error', async () => {
      let n = 0;
      globalThis.fetch = vi.fn(async () => {
        if (++n === 1) throw Object.assign(new TypeError('fetch failed'), { cause: { code: 'ECONNREFUSED' } });
        return json({});
      }) as never;
      await client().request('/me/sendMail', { method: 'POST', body: '{}' });
      expect(n).toBe(2);
    }, 10_000);

    it('final 429 maps to exit code 6', async () => {
      globalThis.fetch = vi.fn(async () => new Response('', { status: 429, headers: { 'Retry-After': '0' } })) as never;
      const err = await client().request('/me').catch(e => e);
      expect(err).toBeInstanceOf(CliError);
      expect((err as CliError).exitCode).toBe(ExitCode.THROTTLED);
    });
  });

  describe('pagination (C-8)', () => {
    it('sends the original headers on every page and refreshes on 401', async () => {
      const seen: Array<{ url: string; cl?: string; auth?: string }> = [];
      let n = 0;
      globalThis.fetch = vi.fn(async (url: string, init: RequestInit) => {
        const h = init.headers as Record<string, string>;
        seen.push({ url, cl: h.ConsistencyLevel, auth: h.Authorization });
        n++;
        if (n === 1) return json({ value: [1], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/groups?$skiptoken=a' });
        if (n === 2) return json({ error: { message: 'expired' } }, 401);
        return json({ value: [2] });
      }) as never;
      auth.getToken.mockResolvedValueOnce('t1').mockResolvedValueOnce('t1').mockResolvedValueOnce('t2');
      const r = await client().requestAllPages('/groups', { headers: { ConsistencyLevel: 'eventual' } });
      expect((r.data as { value: number[] }).value).toEqual([1, 2]);
      expect(seen.every(s => s.cl === 'eventual')).toBe(true);
      expect(seen[2].auth).toBe('Bearer t2');
    });

    it('refuses to send the token to a non-Graph absolute URL', async () => {
      await expect(client().request('https://evil.example.com/x')).rejects.toThrow(/non-Graph/);
    });
  });

  describe('OData metadata (C-9)', () => {
    it('keeps etag/type/count and returns nextLink/deltaLink/count as metadata', async () => {
      globalThis.fetch = vi.fn(async () => json({
        '@odata.context': 'ctx', '@odata.count': 7, '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/d',
        value: [{ '@odata.etag': 'W/"1"', '@odata.type': '#microsoft.graph.fileAttachment', id: 'a' }],
      })) as never;
      const r = await client().request('/x');
      const d = r.data as Record<string, unknown> & { value: Array<Record<string, unknown>> };
      expect(d['@odata.context']).toBeUndefined();
      expect(d['@odata.deltaLink']).toBeUndefined();
      expect(d['@odata.count']).toBe(7);
      expect(d.value[0]['@odata.etag']).toBe('W/"1"');
      expect(r.count).toBe(7);
      expect(r.deltaLink).toBe('https://graph.microsoft.com/v1.0/d');
      expect(r.headers).toBeDefined();
    });
  });

  describe('redirects (G-2 / D15)', () => {
    it('re-sends the token only to trusted hosts', async () => {
      const calls: Array<{ url: string; auth?: string }> = [];
      globalThis.fetch = vi.fn(async (url: string, init: RequestInit) => {
        calls.push({ url, auth: (init.headers as Record<string, string>).Authorization });
        if (url.includes('graph.microsoft.com/v1.0/a')) return new Response(null, { status: 302, headers: { location: 'https://contoso.sharepoint.com/f' } });
        if (url.includes('sharepoint.com')) return new Response(null, { status: 302, headers: { location: 'https://cdn.example.net/blob' } });
        return new Response(new Uint8Array([1, 2]), { status: 200, headers: { 'content-type': 'application/pdf' } });
      }) as never;
      const r = await client().request('/a');
      expect(calls.map(c => !!c.auth)).toEqual([true, true, false]);
      expect((r.data as { size: number }).size).toBe(2);
    });

    it('isTrustedHost', () => {
      expect(isTrustedHost('https://graph.microsoft.com/v1.0')).toBe(true);
      expect(isTrustedHost('https://x.sharepoint.com/a')).toBe(true);
      expect(isTrustedHost('https://x.svc.ms/a')).toBe(true);
      expect(isTrustedHost('https://sharepoint.com.evil.io/a')).toBe(false);
      expect(isTrustedHost('not a url')).toBe(false);
    });
  });
});
