import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { runCli, tempDir } from '../e2e/run-cli.js';

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

  it.each([['JST-9'], ['UTC+3'], ['']])('TZ=%j without an IANA name falls back to UTC (events, decline, workflow)', async (tz) => {
    const env = { ...ENV, TZ: tz };
    const list = await runCli(['calendar', 'events', '-o', 'json'], { env, routes: [{ match: '/me/events', method: 'GET', body: { value: [] } }] });
    expect(list.stderr).toBe('');
    expect(list.code).toBe(0);
    expect(list.requests[0].headers.prefer).toBe('outlook.timezone="UTC"');

    const decline = await runCli(['calendar', 'decline', 'E1', '--propose-start', '2026-10-07T09:00:00', '--propose-end', '2026-10-07T10:00:00'], {
      env, routes: [{ match: '/me/events/E1/decline$', method: 'POST', status: 202, body: '' }],
    });
    expect(decline.code).toBe(0);
    // No zone name: the local wall-clock time is converted to the UTC instant (CAL-4).
    const instant = new Date(Date.UTC(2026, 9, 7, 9) - ({ 'JST-9': 9, 'UTC+3': -3, '': 0 }[tz]! * 3600_000)).toISOString().slice(0, 19);
    expect(JSON.parse(decline.requests[0].body!).proposedNewTime.start).toEqual({ dateTime: instant, timeZone: 'UTC' });

    const digest = await runCli(['workflow', 'digest', '-o', 'json'], {
      env,
      routes: [
        { match: '/me/calendarView', method: 'GET', body: { value: [] } },
        { match: '/me/mailFolders/inbox\\?', method: 'GET', body: { unreadItemCount: 0 } },
        { match: '/me/todo/lists(\\?|$)', method: 'GET', body: { value: [] } },
      ],
    });
    expect(digest.stderr).toBe('');
    expect(digest.code).toBe(0);
    expect(digest.requests[0].headers.prefer).toBe('outlook.timezone="UTC"');
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

describe('calendar dry-run (DRY-1)', () => {
  it.each([['accept', 'Event accepted'], ['decline', 'Event declined'], ['tentative', 'Tentatively accepted']])('--dry-run %s prints only the preview', async (cmd, message) => {
    const r = await runCli(['--dry-run', 'calendar', cmd, 'E1'], { env: ENV });
    expect(r.code).toBe(0);
    expect(r.stderr).toContain('[dry-run] POST https://graph.microsoft.com/v1.0/me/events/E1/');
    expect(r.stderr).not.toContain(message);
  });

  it('a real accept still confirms on stderr', async () => {
    const r = await runCli(['calendar', 'accept', 'E1'], { env: ENV, routes: [{ match: '/me/events/E1/accept$', method: 'POST', status: 202, body: '' }] });
    expect(r.code).toBe(0);
    expect(r.stderr).toContain('Event accepted.');
  });
});

describe('calendar free-busy --interval (CAL-3)', () => {
  const args = ['calendar', 'free-busy', '--emails', 'a@b.cd', '--start', '2026-10-07T09:00:00', '--end', '2026-10-07T10:00:00', '--timezone', 'UTC', '-o', 'json'];
  const route = { match: '/me/calendar/getSchedule$', method: 'POST', body: { value: [] } };

  it.each([['abc'], ['15x'], ['0'], ['1441']])('--interval %s exits 1 before any request', async (bad) => {
    const r = await runCli([...args, '--interval', bad], { env: ENV, routes: [route] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toContain('Expected an integer from 5 to 1440.');
  });

  it('sends --interval as a number, default 30', async () => {
    const given = await runCli([...args, '--interval', '15'], { env: ENV, routes: [route] });
    expect(given.code).toBe(0);
    expect(JSON.parse(given.requests[0].body!).availabilityViewInterval).toBe(15);
    const dflt = await runCli(args, { env: ENV, routes: [route] });
    expect(dflt.code).toBe(0);
    expect(JSON.parse(dflt.requests[0].body!).availabilityViewInterval).toBe(30);
  });
});

describe('calendar time zones (CAL-4)', () => {
  const create = ['calendar', 'create', '--subject', 's', '--start', '2026-10-07T09:00', '--end', '2026-10-07T10:00', '-o', 'json'];
  const createRoute = { match: '/me/events$', method: 'POST', status: 201, body: { id: 'E9' } };
  const sent = async (args: string[], tz: string, route: { match: string; method: string; status?: number; body: unknown }) => {
    const r = await runCli(args, { env: { ...ENV, TZ: tz }, routes: [route] });
    expect(r.stderr).toBe('');
    expect(r.code).toBe(0);
    return JSON.parse(r.requests[0].body!);
  };

  it.each([
    ['America/New_York', '2026-10-07T09:00', '2026-10-07T10:00', 'America/New_York'],
    ['Asia/Kolkata', '2026-10-07T09:00', '2026-10-07T10:00', 'Asia/Kolkata'],
    ['Europe/Amsterdam', '2026-10-07T07:00:00', '2026-10-07T08:00:00', 'UTC'],
    ['JST-9', '2026-10-07T00:00:00', '2026-10-07T01:00:00', 'UTC'],
  ])('create with TZ=%s sends a zone Graph documents', async (tz, start, end, zone) => {
    const body = await sent(create, tz, createRoute);
    expect(body.start).toEqual({ dateTime: start, timeZone: zone });
    expect(body.end).toEqual({ dateTime: end, timeZone: zone });
  });

  it('an explicit --timezone wins (create, update, free-busy)', async () => {
    const tz = 'Europe/Amsterdam';
    const created = await sent([...create, '--timezone', 'W. Europe Standard Time'], tz, createRoute);
    expect(created.start).toEqual({ dateTime: '2026-10-07T09:00', timeZone: 'W. Europe Standard Time' });
    const updated = await sent(['calendar', 'update', 'E1', '--start', '2026-10-07T09:00', '--timezone', 'Europe/Berlin', '-o', 'json'], tz, { match: '/me/events/E1$', method: 'PATCH', body: { id: 'E1' } });
    expect(updated.start).toEqual({ dateTime: '2026-10-07T09:00', timeZone: 'Europe/Berlin' });
    const fb = await sent(['calendar', 'free-busy', '--emails', 'a@b.cd', '--start', '2026-10-07T09:00', '--end', '2026-10-07T10:00', '--timezone', 'UTC', '-o', 'json'], tz, { match: '/me/calendar/getSchedule$', method: 'POST', body: { value: [] } });
    expect(fb.startTime).toEqual({ dateTime: '2026-10-07T09:00', timeZone: 'UTC' });
  });

  it('update and free-busy default to the local zone, not UTC', async () => {
    const updated = await sent(['calendar', 'update', 'E1', '--start', '2026-10-07T09:00', '--end', '2026-10-07T10:00', '-o', 'json'], 'America/New_York', { match: '/me/events/E1$', method: 'PATCH', body: { id: 'E1' } });
    expect(updated.start).toEqual({ dateTime: '2026-10-07T09:00', timeZone: 'America/New_York' });
    expect(updated.end).toEqual({ dateTime: '2026-10-07T10:00', timeZone: 'America/New_York' });
    const fb = await sent(['calendar', 'free-busy', '--emails', 'a@b.cd', '--start', '2026-10-07T09:00', '--end', '2026-10-07T10:00', '-o', 'json'], 'Europe/Amsterdam', { match: '/me/calendar/getSchedule$', method: 'POST', body: { value: [] } });
    expect(fb.startTime).toEqual({ dateTime: '2026-10-07T07:00:00', timeZone: 'UTC' });
    expect(fb.endTime).toEqual({ dateTime: '2026-10-07T08:00:00', timeZone: 'UTC' });
  });

  it('help shows no machine-specific default zone', async () => {
    for (const cmd of ['create', 'update', 'free-busy', 'decline']) {
      const tokyo = await runCli(['calendar', cmd, '--help'], { env: { ...ENV, TZ: 'Asia/Tokyo' } });
      const ny = await runCli(['calendar', cmd, '--help'], { env: { ...ENV, TZ: 'America/New_York' } });
      expect(tokyo.stdout).toBe(ny.stdout);
      expect(tokyo.stdout.replace(/\s+/g, ' ')).toContain('(default: local time zone)');
      expect(tokyo.stdout).not.toMatch(/default: "?UTC/);
    }
  });

  it('generated calendar skill is the same in every zone', async () => {
    const skill = async (tz: string) => {
      const out = tempDir('ms365-skills-');
      const r = await runCli(['generate-skills', '--output-dir', out], { env: { ...ENV, TZ: tz } });
      expect(r.code).toBe(0);
      return fs.readFileSync(path.join(out, 'ms365-calendar', 'SKILL.md'), 'utf8');
    };
    const tokyo = await skill('Asia/Tokyo');
    expect(tokyo).toBe(await skill('America/New_York'));
    expect(tokyo).not.toContain('Asia/Tokyo');
  });
});

describe('calendar colors (CHOICE-1)', () => {
  it.each([
    ['create-calendar', '--name', 'n'],
    ['update-calendar', 'C1'],
  ])('%s rejects an unknown --color before any request', async (...args) => {
    const r = await runCli(['calendar', ...args, '--color', 'purple'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toContain('Allowed choices are auto, lightBlue, lightGreen, lightOrange, lightGray, lightYellow, lightTeal, lightPink, lightBrown, lightRed, maxColor');
  });

  it('update-calendar sends a listed --color', async () => {
    const r = await runCli(['calendar', 'update-calendar', 'C1', '--color', 'lightTeal', '-o', 'json'], { env: ENV, routes: [{ match: '/me/calendars/C1$', method: 'PATCH', body: { id: 'C1' } }] });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.requests[0].body!)).toEqual({ color: 'lightTeal' });
  });
});
