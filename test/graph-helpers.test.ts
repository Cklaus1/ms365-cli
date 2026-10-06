import { describe, it, expect, afterEach, vi } from 'vitest';
import { GraphClient } from '../src/graph-client.js';
import { odataQuote, searchQuote, encodeDrivePath, drivePath, preferTimezone } from '../src/graph-helpers.js';
import { mockFetch, jsonResponse, fakeAuth, type Handler } from './helpers/fetch-mock.js';

const BASE = 'https://graph.microsoft.com/v1.0';

let restore: (() => void) | undefined;
afterEach(() => {
  restore?.();
  restore = undefined;
  vi.restoreAllMocks();
});

function install(handler: Handler) {
  const m = mockFetch(handler);
  restore = m.restore;
  return m.calls;
}

describe('odataQuote (MAIL-8)', () => {
  it.each([
    ['plain', "'plain'"],
    ["O'Brien", "'O''Brien'"],
    ["''", "''''''"],
    ['', "''"],
  ])('%j -> %s', (input, expected) => {
    expect(odataQuote(input)).toBe(expected);
  });

  it('survives buildUrl: the server decodes the escaped literal intact', () => {
    const url = new GraphClient(fakeAuth()).buildUrl('/users', { $filter: `displayName eq ${odataQuote("O'Brien & Co")}` });
    expect(new URL(url).searchParams.get('$filter')).toBe("displayName eq 'O''Brien & Co'");
  });
});

describe('searchQuote (MAIL-8 / GROUPS-1)', () => {
  it.each([
    ['budget', '"budget"'],
    ['say "hi"', '"say \\"hi\\""'],
    ['back\\slash', '"back\\\\slash"'],
  ])('%j -> %s', (input, expected) => {
    expect(searchQuote(input)).toBe(expected);
  });

  it('a $search value with quotes reaches the server as one phrase', () => {
    const url = new GraphClient(fakeAuth()).buildUrl('/me/messages', { $search: searchQuote('a "b" c') });
    expect(new URL(url).searchParams.get('$search')).toBe('"a \\"b\\" c"');
  });
});

describe('encodeDrivePath / drivePath (DRIVE-1)', () => {
  it.each([
    ['/Docs/notes#v2.md', 'Docs/notes%23v2.md'],
    ['Docs/what?.txt', 'Docs/what%3F.txt'],
    ['/100%25.txt', '100%2525.txt'],
    ['/My Files/a b.pdf', 'My%20Files/a%20b.pdf'],
    ['//Docs///x/', 'Docs/x'],
    ['/Ünïcode/ä.txt', '%C3%9Cn%C3%AFcode/%C3%A4.txt'],
    ['/', ''],
    ['', ''],
  ])('%j -> %j', (input, expected) => {
    expect(encodeDrivePath(input)).toBe(expected);
  });

  it('builds item and child addresses; root has no colon path', () => {
    expect(drivePath('/me/drive/root', '/Docs/notes#v2.md', '/content')).toBe('/me/drive/root:/Docs/notes%23v2.md:/content');
    expect(drivePath('/me/drive/root', '/Docs/', '/children')).toBe('/me/drive/root:/Docs:/children');
    expect(drivePath('/me/drive/root', '/', '/children')).toBe('/me/drive/root/children');
    expect(drivePath('/me/drive/root', '/Docs/a.pdf')).toBe('/me/drive/root:/Docs/a.pdf');
    expect(drivePath('/me/drive/root', '/')).toBe('/me/drive/root');
  });

  it('# ? and % survive the trip through fetch (no truncation, no double decoding)', async () => {
    const calls = install(() => jsonResponse({ id: '1' }));
    await new GraphClient(fakeAuth()).request(drivePath('/me/drive/root', '/Docs/notes#v2?%.md', '/content'), { method: 'PUT', body: 'x' });
    const sent = new URL(calls[0].url);
    expect(sent.hash).toBe('');
    expect(sent.search).toBe('');
    expect(sent.pathname).toBe('/v1.0/me/drive/root:/Docs/notes%23v2%3F%25.md:/content');
  });

  it('dry-run prints the encoded colon path', async () => {
    const err: string[] = [];
    vi.spyOn(process.stderr, 'write').mockImplementation((c: string | Uint8Array) => { err.push(String(c)); return true; });
    await new GraphClient(fakeAuth(), true).request(drivePath('/me/drive/root', '/Docs/notes#v2.md', '/content'), { method: 'PUT', body: 'x' });
    expect(err.join('')).toContain(`[dry-run] PUT ${BASE}/me/drive/root:/Docs/notes%23v2.md:/content`);
  });
});

describe('per-request API version (CAL-1)', () => {
  it('apiVersion beta targets /beta for that request only', async () => {
    const calls = install(() => jsonResponse({ value: [] }));
    const client = new GraphClient(fakeAuth());
    await client.request('/me/findRooms', { apiVersion: 'beta' });
    await client.request('/me/calendars');
    expect(calls.map((c) => c.url)).toEqual([
      'https://graph.microsoft.com/beta/me/findRooms',
      `${BASE}/me/calendars`,
    ]);
  });

  it('applies to every page of requestAllPages and to dry-run', async () => {
    const err: string[] = [];
    vi.spyOn(process.stderr, 'write').mockImplementation((c: string | Uint8Array) => { err.push(String(c)); return true; });
    await new GraphClient(fakeAuth(), true).requestAllPages('/me/findRoomLists', { apiVersion: 'beta' });
    expect(err.join('')).toContain('[dry-run] GET https://graph.microsoft.com/beta/me/findRoomLists');
  });
});

describe('preferTimezone (WF-1)', () => {
  it('produces the outlook.timezone Prefer header', () => {
    expect(preferTimezone({}, 'Europe/Berlin')).toEqual({ Prefer: 'outlook.timezone="Europe/Berlin"' });
  });

  it('defaults to the local IANA zone', () => {
    const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
    expect(preferTimezone()).toEqual({ Prefer: `outlook.timezone="${tz}"` });
  });

  it('keeps other headers and existing Prefer values (comma-joined), without mutating the input', () => {
    const input = { Prefer: 'outlook.body-content-type="text"', 'X-A': '1' };
    expect(preferTimezone(input, 'UTC')).toEqual({
      Prefer: 'outlook.body-content-type="text", outlook.timezone="UTC"',
      'X-A': '1',
    });
    expect(input.Prefer).toBe('outlook.body-content-type="text"');
  });

  it('rejects a non-IANA zone instead of sending a header Graph ignores', () => {
    expect(() => preferTimezone({}, 'Not/AZone')).toThrow(/time zone/i);
  });

  it('is sent on every page', async () => {
    const calls = install((_c, i) => jsonResponse(i === 0
      ? { value: [{ id: '1' }], '@odata.nextLink': `${BASE}/me/calendarView?$skiptoken=2` }
      : { value: [{ id: '2' }] }));
    await new GraphClient(fakeAuth()).requestAllPages('/me/calendarView', { headers: preferTimezone({}, 'Asia/Tokyo') });
    expect(calls.map((c) => c.headers.prefer)).toEqual(['outlook.timezone="Asia/Tokyo"', 'outlook.timezone="Asia/Tokyo"']);
  });
});
