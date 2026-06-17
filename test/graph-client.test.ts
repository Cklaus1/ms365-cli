import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

describe('GraphClient', () => {
  let GraphClient: typeof import('../src/graph-client.js').GraphClient;
  let mockAuth: any;

  beforeEach(async () => {
    const mod = await import('../src/graph-client.js');
    GraphClient = mod.GraphClient;
    mockAuth = {
      getToken: vi.fn().mockResolvedValue('mock-token'),
      initialize: vi.fn(),
    };
  });

  describe('buildUrl', () => {
    it('builds base URL with endpoint', () => {
      const client = new GraphClient(mockAuth, false, 'v1.0', false);
      expect(client.buildUrl('/me')).toBe('https://graph.microsoft.com/v1.0/me');
    });

    it('appends query params', () => {
      const client = new GraphClient(mockAuth);
      const url = client.buildUrl('/me/messages', { '$top': '10', '$select': 'id,subject' });
      expect(url).toContain('?');
      expect(url).toContain('$top=10');
      expect(url).toContain('$select=id%2Csubject');
    });

    it('preserves $ prefix on OData params', () => {
      const client = new GraphClient(mockAuth);
      const url = client.buildUrl('/me', { '$filter': "displayName eq 'Test'" });
      expect(url).toContain('$filter=');
    });

    it('encodes non-OData params normally', () => {
      const client = new GraphClient(mockAuth);
      const url = client.buildUrl('/me', { search: 'hello world' });
      expect(url).toContain('search=hello%20world');
    });

    it('uses beta API version when configured', () => {
      const client = new GraphClient(mockAuth, false, 'beta');
      expect(client.buildUrl('/me')).toBe('https://graph.microsoft.com/beta/me');
    });

    it('appends to existing query string with &', () => {
      const client = new GraphClient(mockAuth);
      const url = client.buildUrl("/me/drive/root/search(q='test')?existing=1", { '$top': '5' });
      expect(url).toContain('&$top=5');
      expect(url).not.toContain('?$top'); // should use & not ?
    });

    it('handles empty queryParams', () => {
      const client = new GraphClient(mockAuth);
      const url = client.buildUrl('/me', {});
      expect(url).toBe('https://graph.microsoft.com/v1.0/me');
    });
  });

  describe('read-only mode', () => {
    it('blocks POST requests', async () => {
      const client = new GraphClient(mockAuth, false, 'v1.0', true);
      await expect(
        client.request('/me/messages', { method: 'POST', body: '{}' })
      ).rejects.toThrow('Read-only mode');
    });

    it('blocks PATCH requests', async () => {
      const client = new GraphClient(mockAuth, false, 'v1.0', true);
      await expect(
        client.request('/me/messages/123', { method: 'PATCH', body: '{}' })
      ).rejects.toThrow('Read-only mode');
    });

    it('blocks DELETE requests', async () => {
      const client = new GraphClient(mockAuth, false, 'v1.0', true);
      await expect(
        client.request('/me/messages/123', { method: 'DELETE' })
      ).rejects.toThrow('Read-only mode');
    });

    it('allows GET requests', async () => {
      const client = new GraphClient(mockAuth, false, 'v1.0', true);
      const err = await client.request('/me').catch(e => e);
      expect(err.message).not.toContain('Read-only mode');
    });

    it('blocks file uploads before touching filesystem', async () => {
      const client = new GraphClient(mockAuth, false, 'v1.0', true);
      await expect(
        client.uploadFile('/me/drive/root:/test.txt:/content', '/nonexistent/path', 'text/plain')
      ).rejects.toThrow('Read-only mode');
    });
  });

  describe('dry-run mode', () => {
    it('returns dry-run response for POST', async () => {
      const client = new GraphClient(mockAuth, true);
      const spy = vi.spyOn(console, 'log').mockImplementation(() => {});
      const response = await client.request('/me/messages', {
        method: 'POST',
        body: JSON.stringify({ subject: 'test' }),
      });
      expect(response.status).toBe(0);
      expect((response.data as any).dryRun).toBe(true);
      expect((response.data as any).method).toBe('POST');
      spy.mockRestore();
    });

    it('returns dry-run response for GET (default method)', async () => {
      const client = new GraphClient(mockAuth, true);
      const spy = vi.spyOn(console, 'log').mockImplementation(() => {});
      const response = await client.request('/me');
      expect((response.data as any).dryRun).toBe(true);
      expect((response.data as any).method).toBe('GET');
      spy.mockRestore();
    });

    it('logs headers in dry-run when present', async () => {
      const client = new GraphClient(mockAuth, true);
      const spy = vi.spyOn(console, 'log').mockImplementation(() => {});
      await client.request('/me', {
        headers: { 'Prefer': 'outlook.body-content-type="text"' },
      });
      const headerLog = spy.mock.calls.find(c => String(c[0]).includes('Headers'));
      expect(headerLog).toBeDefined();
      spy.mockRestore();
    });

    it('dry-run blocks read-only check', async () => {
      // dry-run + read-only: dry-run should take precedence for non-GET
      const client = new GraphClient(mockAuth, true, 'v1.0', true);
      // read-only check runs before dry-run check, so this should throw
      await expect(
        client.request('/me/messages', { method: 'POST', body: '{}' })
      ).rejects.toThrow('Read-only mode');
    });
  });

  describe('requestPagesStreaming', () => {
    it('is an async generator', () => {
      const client = new GraphClient(mockAuth, true);
      const gen = client.requestPagesStreaming('/me/messages');
      expect(typeof gen[Symbol.asyncIterator]).toBe('function');
    });

    it('yields dry-run data for single page', async () => {
      const client = new GraphClient(mockAuth, true);
      const spy = vi.spyOn(console, 'log').mockImplementation(() => {});
      const pages: unknown[][] = [];
      for await (const page of client.requestPagesStreaming('/me/messages')) {
        pages.push(page);
      }
      expect(pages.length).toBe(1);
      expect(pages[0].length).toBe(1); // single dry-run response wrapped
      spy.mockRestore();
    });
  });

  describe('request with fetch mock', () => {
    let originalFetch: typeof globalThis.fetch;

    beforeEach(() => {
      originalFetch = globalThis.fetch;
    });

    afterEach(() => {
      globalThis.fetch = originalFetch;
    });

    it('retries on 429 with Retry-After header', async () => {
      let attempt = 0;
      globalThis.fetch = vi.fn(async () => {
        attempt++;
        if (attempt === 1) {
          return new Response('rate limited', {
            status: 429,
            headers: { 'Retry-After': '0' },
          });
        }
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        });
      }) as any;

      const client = new GraphClient(mockAuth);
      const errSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
      const response = await client.request('/me/messages');
      expect(attempt).toBe(2);
      expect(response.status).toBe(200);
      errSpy.mockRestore();
    });

    it('retries on 500 server error', async () => {
      let attempt = 0;
      globalThis.fetch = vi.fn(async () => {
        attempt++;
        if (attempt === 1) {
          return new Response('server error', { status: 500 });
        }
        return new Response(JSON.stringify({ id: '1' }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        });
      }) as any;

      const client = new GraphClient(mockAuth);
      const errSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
      const response = await client.request('/me');
      expect(attempt).toBe(2);
      expect(response.status).toBe(200);
      errSpy.mockRestore();
    });

    it('retries once on 401 with fresh token', async () => {
      let attempt = 0;
      globalThis.fetch = vi.fn(async () => {
        attempt++;
        if (attempt === 1) {
          return new Response(JSON.stringify({ error: { message: 'token expired' } }), {
            status: 401,
            headers: { 'content-type': 'application/json' },
          });
        }
        return new Response(JSON.stringify({ id: 'ok' }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        });
      }) as any;

      mockAuth.getToken = vi.fn()
        .mockResolvedValueOnce('old-token')
        .mockResolvedValueOnce('fresh-token');

      const client = new GraphClient(mockAuth);
      const response = await client.request('/me');
      expect(attempt).toBe(2);
      expect(response.status).toBe(200);
      // Token was force-refreshed
      expect(mockAuth.getToken).toHaveBeenCalledWith(true);
    });

    it('throws on 403 with permission hint', async () => {
      globalThis.fetch = vi.fn(async () => {
        return new Response(
          JSON.stringify({ error: { message: 'Insufficient scope permissions' } }),
          { status: 403, headers: { 'content-type': 'application/json' } },
        );
      }) as any;

      const client = new GraphClient(mockAuth);
      await expect(client.request('/me/teams')).rejects.toThrow('--org-mode');
    });

    it('handles binary content responses', async () => {
      const binaryData = new Uint8Array([0x89, 0x50, 0x4e, 0x47]); // PNG header
      globalThis.fetch = vi.fn(async () => {
        return new Response(binaryData, {
          status: 200,
          headers: { 'content-type': 'image/png' },
        });
      }) as any;

      const client = new GraphClient(mockAuth);
      const response = await client.request('/me/photo/$value');
      const data = response.data as any;
      expect(data.contentBytes).toBeDefined();
      expect(data.contentType).toBe('image/png');
      expect(data.size).toBe(4);
    });

    it('handles empty response body', async () => {
      globalThis.fetch = vi.fn(async () => {
        return new Response(null, { status: 200 });
      }) as any;

      const client = new GraphClient(mockAuth);
      const response = await client.request('/me/messages/123', { method: 'DELETE' });
      expect(response.status).toBe(200);
      expect((response.data as any).message).toBe('OK');
    });

    it('strips @odata properties from response', async () => {
      globalThis.fetch = vi.fn(async () => {
        return new Response(JSON.stringify({
          '@odata.context': 'https://graph.microsoft.com/...',
          '@odata.nextLink': 'https://graph.microsoft.com/...?skiptoken=abc',
          id: '1',
          value: [{ '@odata.type': '#message', id: 'm1' }],
        }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        });
      }) as any;

      const client = new GraphClient(mockAuth);
      const response = await client.request('/me/messages');
      const data = response.data as any;
      expect(data['@odata.context']).toBeUndefined();
      expect(data.id).toBe('1');
      expect(data.value[0]['@odata.type']).toBeUndefined();
      expect(data.value[0].id).toBe('m1');
      // nextLink extracted before stripping
      expect(response.nextLink).toBe('https://graph.microsoft.com/...?skiptoken=abc');
    });

    it('paginates with requestAllPages', async () => {
      let callCount = 0;
      globalThis.fetch = vi.fn(async (url: string) => {
        callCount++;
        if (callCount === 1) {
          return new Response(JSON.stringify({
            '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?skip=10',
            value: [{ id: '1' }, { id: '2' }],
          }), { status: 200, headers: { 'content-type': 'application/json' } });
        }
        return new Response(JSON.stringify({
          value: [{ id: '3' }],
        }), { status: 200, headers: { 'content-type': 'application/json' } });
      }) as any;

      const client = new GraphClient(mockAuth);
      const response = await client.requestAllPages('/me/messages');
      const data = response.data as any;
      expect(data.value).toHaveLength(3);
      expect(data.value.map((i: any) => i.id)).toEqual(['1', '2', '3']);
    });

    it('respects pageLimit in requestAllPages', async () => {
      globalThis.fetch = vi.fn(async () => {
        return new Response(JSON.stringify({
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next',
          value: [{ id: '1' }],
        }), { status: 200, headers: { 'content-type': 'application/json' } });
      }) as any;

      const client = new GraphClient(mockAuth);
      const errSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
      const response = await client.requestAllPages('/me/messages', {}, 1);
      const data = response.data as any;
      // Only 1 page fetched despite nextLink
      expect(data.value).toHaveLength(1);
      expect(errSpy).toHaveBeenCalledWith(expect.stringContaining('page limit'));
      errSpy.mockRestore();
    });
  });
});
