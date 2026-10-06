import { describe, it, expect, afterEach, vi } from 'vitest';
import { GraphClient } from '../src/graph-client.js';
import { setQuiet } from '../src/utils.js';
import { mockFetch, sequence, jsonResponse, fakeAuth, valuesOf, type Handler } from './helpers/fetch-mock.js';

let restore: (() => void) | undefined;
afterEach(() => {
  restore?.();
  restore = undefined;
  setQuiet(false);
  vi.restoreAllMocks();
});

function install(handler: Handler) {
  const m = mockFetch(handler);
  restore = m.restore;
  return m.calls;
}

describe('GRAPH-4 raw payload with OData metadata', () => {
  it('keeps @odata annotations in data and surfaces nextLink/deltaLink/count', async () => {
    install(() => jsonResponse({
      '@odata.context': 'ctx',
      '@odata.count': 42,
      '@odata.nextLink': 'https://graph.microsoft.com/v1.0/x?$skiptoken=2',
      '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/x?$deltatoken=9',
      value: [{ '@odata.etag': 'W/"abc"', '@odata.type': '#microsoft.graph.fileAttachment', id: 'm1' }],
    }));
    const res = await new GraphClient(fakeAuth()).request('/x');
    expect(res.nextLink).toBe('https://graph.microsoft.com/v1.0/x?$skiptoken=2');
    expect(res.deltaLink).toBe('https://graph.microsoft.com/v1.0/x?$deltatoken=9');
    expect(res.count).toBe(42);
    expect(valuesOf(res.data)[0]).toEqual({ '@odata.etag': 'W/"abc"', '@odata.type': '#microsoft.graph.fileAttachment', id: 'm1' });
  });

  it('requestAllPages keeps item annotations, count from the first page and deltaLink from the last', async () => {
    install(sequence(
      () => jsonResponse({ '@odata.count': 2, '@odata.nextLink': 'https://graph.microsoft.com/v1.0/d?p=2', value: [{ '@odata.etag': 'e1', id: '1' }] }),
      () => jsonResponse({ '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/d?token=t', value: [{ id: '2' }] }),
    ));
    const res = await new GraphClient(fakeAuth()).requestAllPages('/d');
    expect(valuesOf(res.data)).toEqual([{ '@odata.etag': 'e1', id: '1' }, { id: '2' }]);
    expect(res.count).toBe(2);
    expect(res.deltaLink).toBe('https://graph.microsoft.com/v1.0/d?token=t');
    expect(res.nextLink).toBeUndefined();
  });

  it('requestAllPages returns the unfollowed nextLink when the page limit stops it', async () => {
    vi.spyOn(console, 'error').mockImplementation(() => {});
    install(() => jsonResponse({ '@odata.nextLink': 'https://graph.microsoft.com/v1.0/d?p=next', value: [{ id: '1' }] }));
    const res = await new GraphClient(fakeAuth()).requestAllPages('/d', {}, 1);
    expect(valuesOf(res.data)).toHaveLength(1);
    expect(res.nextLink).toBe('https://graph.microsoft.com/v1.0/d?p=next');
  });
});

describe('CORE-4 dry-run preview goes to stderr', () => {
  function capture() {
    const out = vi.spyOn(process.stdout, 'write').mockImplementation(() => true);
    const log = vi.spyOn(console, 'log').mockImplementation(() => {});
    const err = vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
    return { out, log, err, stderr: () => err.mock.calls.map((c) => String(c[0])).join('') };
  }

  it('prints method, URL, body and headers on stderr only, even with --quiet', async () => {
    setQuiet(true);
    const io = capture();
    const res = await new GraphClient(fakeAuth(), true).request('/me/sendMail', {
      method: 'POST',
      body: JSON.stringify({ subject: 'hi' }),
      headers: { Prefer: 'outlook.timezone="UTC"' },
    });
    expect(io.out).not.toHaveBeenCalled();
    expect(io.log).not.toHaveBeenCalled();
    expect(io.stderr()).toContain('[dry-run] POST https://graph.microsoft.com/v1.0/me/sendMail');
    expect(io.stderr()).toContain('{"subject":"hi"}');
    expect(io.stderr()).toContain('outlook.timezone');
    expect(res.data).toEqual({ dryRun: true, method: 'POST', url: 'https://graph.microsoft.com/v1.0/me/sendMail', body: { subject: 'hi' } });
  });

  it('keeps a non-JSON body as a string', async () => {
    capture();
    const res = await new GraphClient(fakeAuth(), true).request('/x', { method: 'POST', body: 'plain' });
    expect(res.data).toMatchObject({ body: 'plain' });
  });
});
