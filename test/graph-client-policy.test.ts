import { describe, it, expect, afterEach, vi } from 'vitest';
import { GraphClient } from '../src/graph-client.js';
import { mockFetch, sequence, jsonResponse, graphError, abortError, fakeAuth, valuesOf, type Handler } from './helpers/fetch-mock.js';

let restore: (() => void) | undefined;
afterEach(() => {
  restore?.();
  restore = undefined;
  vi.restoreAllMocks();
});

function install(handler: Handler) {
  const m = mockFetch(handler);
  restore = m.restore;
  vi.spyOn(console, 'error').mockImplementation(() => {});
  return m.calls;
}

const busy = () => graphError(503, 'busy', { 'Retry-After': '0' });
const ok = () => jsonResponse({ id: 'ok' });

describe('GRAPH-1 retry policy', () => {
  it('POST + 503 is sent once and says it may have been applied', async () => {
    const calls = install(sequence(busy, ok));
    const err = await new GraphClient(fakeAuth()).request('/me/sendMail', { method: 'POST', body: '{}' }).catch((e: Error) => e);
    expect(calls).toHaveLength(1);
    expect(String(err)).toMatch(/may have been applied/);
  });

  it('PATCH + 500 is not retried', async () => {
    const calls = install(sequence(() => graphError(500), ok));
    await expect(new GraphClient(fakeAuth()).request('/me/events/1', { method: 'PATCH', body: '{}' })).rejects.toThrow(/500/);
    expect(calls).toHaveLength(1);
  });

  it('GET + 503 is attempted 3 times', async () => {
    const calls = install(sequence(busy));
    await expect(new GraphClient(fakeAuth()).request('/me')).rejects.toThrow(/503/);
    expect(calls).toHaveLength(3);
  });

  it('GET + 503 recovers on a later attempt', async () => {
    const calls = install(sequence(busy, busy, ok));
    const res = await new GraphClient(fakeAuth()).request('/me');
    expect(calls).toHaveLength(3);
    expect(res.data).toEqual({ id: 'ok' });
  });

  it.each(['PUT', 'DELETE'])('%s + 503 is retried', async (method) => {
    const calls = install(sequence(busy, ok));
    await new GraphClient(fakeAuth()).request('/me/x', { method });
    expect(calls).toHaveLength(2);
  });

  it('POST + 429 is retried', async () => {
    const calls = install(sequence(() => graphError(429, 'slow down', { 'Retry-After': '0' }), ok));
    const res = await new GraphClient(fakeAuth()).request('/me/sendMail', { method: 'POST', body: '{}' });
    expect(calls).toHaveLength(2);
    expect(calls[1].body).toBe('{}');
    expect(res.status).toBe(200);
  });

  it('POST timeout is not retried and says it may have been applied', async () => {
    const calls = install(() => { throw abortError(); });
    const err = await new GraphClient(fakeAuth()).request('/chats/1/messages', { method: 'POST', body: '{}' }).catch((e: Error) => e);
    expect(calls).toHaveLength(1);
    expect(String(err)).toMatch(/timed out/);
    expect(String(err)).toMatch(/may have been applied/);
  });

  it('GET timeout is retried', async () => {
    const calls = install(sequence(() => { throw abortError(); }, ok));
    const res = await new GraphClient(fakeAuth()).request('/me');
    expect(calls).toHaveLength(2);
    expect(res.data).toEqual({ id: 'ok' });
  });

  it('a 4xx error is not retried and carries the Graph message', async () => {
    const calls = install(sequence(() => graphError(404, 'Item not found')));
    await expect(new GraphClient(fakeAuth()).request('/me/x')).rejects.toThrow('Graph API error (404): Item not found');
    expect(calls).toHaveLength(1);
  });
});

describe('GRAPH-7 401 refresh', () => {
  it('forces a token refresh once and retries with the new token', async () => {
    const auth = fakeAuth();
    const calls = install(sequence(() => graphError(401, 'expired'), ok));
    const res = await new GraphClient(auth).request('/me');
    expect(res.data).toEqual({ id: 'ok' });
    expect(auth.getToken).toHaveBeenCalledWith(true);
    expect(calls.map((c) => c.headers.authorization)).toEqual(['Bearer token-1', 'Bearer token-2']);
  });

  it('a second 401 fails with the login hint (no refresh loop)', async () => {
    const auth = fakeAuth();
    const calls = install(sequence(() => graphError(401, 'expired')));
    await expect(new GraphClient(auth).request('/me')).rejects.toThrow('ms365 auth login');
    expect(calls).toHaveLength(2);
  });

  it('a 401 on POST is retried (rejected before processing)', async () => {
    const calls = install(sequence(() => graphError(401), ok));
    await new GraphClient(fakeAuth()).request('/me/sendMail', { method: 'POST', body: '{}' });
    expect(calls).toHaveLength(2);
  });
});

describe('GRAPH-5 one request core for nextLink pages', () => {
  const page1 = () => jsonResponse({ '@odata.nextLink': 'https://graph.microsoft.com/v1.0/groups?$skiptoken=2', value: [{ id: '1' }] });
  const page2 = () => jsonResponse({ value: [{ id: '2' }] });

  it('propagates caller headers and auth to every page', async () => {
    const calls = install(sequence(page1, page2));
    const res = await new GraphClient(fakeAuth()).requestAllPages('/groups', { headers: { ConsistencyLevel: 'eventual' } });
    expect(valuesOf(res.data)).toEqual([{ id: '1' }, { id: '2' }]);
    expect(calls).toHaveLength(2);
    expect(calls[1].url).toBe('https://graph.microsoft.com/v1.0/groups?$skiptoken=2');
    expect(calls[1].headers.consistencylevel).toBe('eventual');
    expect(calls[1].headers.authorization).toBe('Bearer token-1');
  });

  it('refreshes the token on a 401 for a nextLink page', async () => {
    const auth = fakeAuth();
    const calls = install(sequence(page1, () => graphError(401), page2));
    const res = await new GraphClient(auth).requestAllPages('/groups');
    expect(valuesOf(res.data)).toHaveLength(2);
    expect(auth.getToken).toHaveBeenCalledWith(true);
    expect(calls[2].headers.authorization).toBe('Bearer token-2');
  });

  it('reports a parsed Graph error for a failing nextLink page', async () => {
    install(sequence(page1, () => graphError(400, 'Bad skiptoken')));
    await expect(new GraphClient(fakeAuth()).requestAllPages('/groups')).rejects.toThrow('Graph API error (400): Bad skiptoken');
  });
});
