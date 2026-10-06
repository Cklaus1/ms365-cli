import { describe, it, expect } from 'vitest';
import { runCli } from '../e2e/run-cli.js';

const LA = 'America/Los_Angeles';
const ENV = {
  MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000',
  MS365_CLI_TOKEN: 'fake',
  TZ: LA,
};

describe('calendar rooms (CAL-1)', () => {
  it.each([['rooms', 'findRooms'], ['room-lists', 'findRoomLists']])('cal %s calls the beta endpoint', async (cmd, fn) => {
    const r = await runCli(['calendar', cmd, '-o', 'json'], {
      env: ENV,
      routes: [{ match: `/beta/me/${fn}$`, method: 'GET', body: { value: [{ name: 'Room 1', address: 'r1@x.test' }] } }],
    });
    expect(r.code).toBe(0);
    expect(r.requests.map((q) => q.url)).toEqual([`https://graph.microsoft.com/beta/me/${fn}`]);
    expect(JSON.parse(r.stdout).value[0].name).toBe('Room 1');
  });
});

describe('calendar decline (CAL-2)', () => {
  it('rejects --no-notify with --propose-start before any request', async () => {
    const r = await runCli(['calendar', 'decline', 'E1', '--no-notify', '--propose-start', '2026-10-07T09:00:00', '--propose-end', '2026-10-07T10:00:00'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toContain('--no-notify');
  });

  it('rejects --propose-end without --propose-start', async () => {
    const r = await runCli(['calendar', 'decline', 'E1', '--propose-end', '2026-10-07T10:00:00'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });

  it('proposes the new time in the local zone (or --timezone), not hard-coded UTC', async () => {
    const route = { match: '/me/events/E1/decline$', method: 'POST', status: 202, body: '' };
    const local = await runCli(['calendar', 'decline', 'E1', '--propose-start', '2026-10-07T09:00:00', '--propose-end', '2026-10-07T10:00:00'], { env: ENV, routes: [route] });
    expect(local.code).toBe(0);
    expect(JSON.parse(local.requests[0].body!)).toEqual({
      comment: '',
      sendResponse: true,
      proposedNewTime: {
        start: { dateTime: '2026-10-07T09:00:00', timeZone: LA },
        end: { dateTime: '2026-10-07T10:00:00', timeZone: LA },
      },
    });
    const explicit = await runCli(['calendar', 'decline', 'E1', '--propose-start', '2026-10-07T09:00:00', '--propose-end', '2026-10-07T10:00:00', '--timezone', 'Europe/Berlin'], { env: ENV, routes: [route] });
    expect(JSON.parse(explicit.requests[0].body!).proposedNewTime.start.timeZone).toBe('Europe/Berlin');
  });
});

describe('calendar update (CAL-2, TODO-1)', () => {
  it('no fields: no request, exit 1', async () => {
    const r = await runCli(['calendar', 'update', 'E1'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toMatch(/nothing to update/i);
  });

  it('--timezone alone: no request, exit 1', async () => {
    const r = await runCli(['calendar', 'update', 'E1', '--timezone', 'Europe/Berlin'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });

  it('update-calendar with no fields: no request, exit 1', async () => {
    const r = await runCli(['calendar', 'update-calendar', 'C1'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });

  it('sends only the given fields', async () => {
    const r = await runCli(['calendar', 'update', 'E1', '--subject', 'New', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/me/events/E1$', method: 'PATCH', body: { id: 'E1', subject: 'New' } }],
    });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.requests[0].body!)).toEqual({ subject: 'New' });
  });
});

describe('calendar reads (WF-1, PAGE-1)', () => {
  it('cal view sends --top, the local Prefer zone, and honours --page-limit', async () => {
    const r = await runCli(['calendar', 'view', '--start', '2026-10-06T00:00:00', '--end', '2026-10-07T00:00:00', '--top', '5', '--all', '--page-limit', '2', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/me/calendarView', method: 'GET', body: { value: [{ id: 'A' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=n' } }],
    });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(2);
    const first = new URL(r.requests[0].url);
    expect(first.searchParams.get('$top')).toBe('5');
    expect(first.searchParams.get('startDateTime')).toBe('2026-10-06T00:00:00');
    for (const q of r.requests) expect(q.headers.prefer).toBe(`outlook.timezone="${LA}"`);
    expect(JSON.parse(r.stdout)).toEqual([{ id: 'A' }, { id: 'A' }]);
  });

  it('cal events and cal get ask for local-zone times', async () => {
    const routes = [
      { match: '/me/events/E1', method: 'GET', body: { id: 'E1' } },
      { match: '/me/events', method: 'GET', body: { value: [] } },
    ];
    const list = await runCli(['calendar', 'events', '-o', 'json'], { env: ENV, routes });
    expect(list.code).toBe(0);
    expect(JSON.parse(list.stdout)).toEqual({ value: [] });
    expect(new URL(list.requests[0].url).searchParams.get('$top')).toBe('10');
    expect(list.requests[0].headers.prefer).toBe(`outlook.timezone="${LA}"`);
    const get = await runCli(['calendar', 'get', 'E1', '-o', 'json'], { env: ENV, routes });
    expect(get.code).toBe(0);
    expect(get.requests[0].headers.prefer).toBe(`outlook.timezone="${LA}"`);
  });

  it('rejects a non-integer --top before any request', async () => {
    const r = await runCli(['calendar', 'view', '--start', 'a', '--end', 'b', '--top', 'x'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });
});

describe('calendar addresses (CAL-2)', () => {
  it('create rejects an invalid attendee before any request', async () => {
    const r = await runCli(['calendar', 'create', '--subject', 's', '--start', '2026-10-07T09:00:00', '--end', '2026-10-07T10:00:00', '--attendees', 'a@b.cd,not-an-email'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toContain('not-an-email');
  });

  it('free-busy rejects an invalid address before any request', async () => {
    const r = await runCli(['calendar', 'free-busy', '--emails', 'bogus', '--start', '2026-10-07T09:00:00', '--end', '2026-10-07T10:00:00'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });

  it('create sends trimmed attendees', async () => {
    const r = await runCli(['calendar', 'create', '--subject', 's', '--start', '2026-10-07T09:00:00', '--end', '2026-10-07T10:00:00', '--attendees', 'a@b.cd, c@d.ef', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/me/events$', method: 'POST', status: 201, body: { id: 'E9' } }],
    });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.requests[0].body!).attendees).toEqual([
      { emailAddress: { address: 'a@b.cd' }, type: 'required' },
      { emailAddress: { address: 'c@d.ef' }, type: 'required' },
    ]);
  });
});
