import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { runCommand, jsonResponse, type RecordedCall } from '../helpers/cli.js';
import { registerCalendarCommands } from '../../src/commands/calendar.js';

const TZ = 'America/Los_Angeles';
const env = { MS365_CLI_TIMEZONE: TZ };
const PREFER = `outlook.timezone="${TZ}"`;

const query = (call: RecordedCall) => new URLSearchParams(call.path.split('?')[1] ?? '');
const writes = (calls: RecordedCall[]) => calls.filter(c => c.method !== 'GET');

beforeEach(() => {
  vi.useFakeTimers({ toFake: ['Date'] });
  vi.setSystemTime(new Date('2026-10-14T17:00:00Z')); // Wed 10:00 PDT
});
afterEach(() => { vi.useRealTimers(); });

describe('cal create (M-9, M-29)', () => {
  it('defaults to the resolved local zone instead of UTC', async () => {
    const r = await runCommand(registerCalendarCommands,
      ['cal', 'create', '--subject', 'S', '--start', '2026-10-20T09:00', '--end', '2026-10-20T09:30'], { env });
    expect(r.exitCode).toBe(0);
    const body = r.calls[0].body as Record<string, { dateTime: string; timeZone: string }>;
    expect(body.start).toEqual({ dateTime: '2026-10-20T09:00:00', timeZone: TZ });
    expect(body.end).toEqual({ dateTime: '2026-10-20T09:30:00', timeZone: TZ });
  });

  it('accepts natural times and an explicit --timezone', async () => {
    const r = await runCommand(registerCalendarCommands,
      ['cal', 'create', '--subject', 'S', '--start', 'tomorrow 15:00', '--end', 'tomorrow 16:00', '--timezone', 'America/New_York'], { env });
    expect(r.exitCode).toBe(0);
    const body = r.calls[0].body as Record<string, { dateTime: string; timeZone: string }>;
    expect(body.start).toEqual({ dateTime: '2026-10-15T15:00:00', timeZone: 'America/New_York' });
    expect(body.end).toEqual({ dateTime: '2026-10-15T16:00:00', timeZone: 'America/New_York' });
  });

  it('rejects an invalid attendee before sending anything', async () => {
    const r = await runCommand(registerCalendarCommands,
      ['cal', 'create', '--subject', 'S', '--start', '2026-10-20T09:00', '--end', '2026-10-20T09:30', '--attendees', 'a@x.com,not-an-email'], { env });
    expect(r.exitCode).toBe(1);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toContain('not-an-email');
  });
});

describe('cal update (M-9)', () => {
  it('reuses the event\'s original time zone when --timezone is omitted', async () => {
    const r = await runCommand(registerCalendarCommands, ['cal', 'update', 'E1', '--start', '2026-10-20T09:00'], {
      env,
      routes: [['GET', '/me/events/E1', () => ({ originalStartTimeZone: 'America/New_York', start: { dateTime: '2026-10-20T13:00:00', timeZone: 'UTC' } })]],
    });
    expect(r.exitCode).toBe(0);
    const patch = writes(r.calls)[0];
    expect(patch.method).toBe('PATCH');
    expect((patch.body as Record<string, unknown>).start).toEqual({ dateTime: '2026-10-20T09:00:00', timeZone: 'America/New_York' });
  });

  it('passes an ISO time through unchanged for a Windows-named event zone', async () => {
    const r = await runCommand(registerCalendarCommands, ['cal', 'update', 'E1', '--end', '2026-10-20T10:00'], {
      env,
      routes: [['GET', '/me/events/E1', () => ({ originalStartTimeZone: 'Eastern Standard Time' })]],
    });
    expect(r.exitCode).toBe(0);
    expect((writes(r.calls)[0].body as Record<string, unknown>).end).toEqual({ dateTime: '2026-10-20T10:00', timeZone: 'Eastern Standard Time' });
  });

  it('uses --timezone without reading the event', async () => {
    const r = await runCommand(registerCalendarCommands, ['cal', 'update', 'E1', '--start', '2026-10-20T09:00', '--timezone', 'Europe/London'], { env });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(1);
    expect((r.calls[0].body as Record<string, unknown>).start).toEqual({ dateTime: '2026-10-20T09:00:00', timeZone: 'Europe/London' });
  });

  it('does not read the event when no time changes', async () => {
    const r = await runCommand(registerCalendarCommands, ['cal', 'update', 'E1', '--subject', 'New'], { env });
    expect(r.exitCode).toBe(0);
    expect(r.calls.map(c => c.method)).toEqual(['PATCH']);
  });
});

describe('cal free-busy (M-9)', () => {
  it('defaults to the local zone', async () => {
    const r = await runCommand(registerCalendarCommands,
      ['cal', 'free-busy', '--emails', 'a@x.com', '--start', '2026-10-20T09:00', '--end', '2026-10-20T17:00'], { env });
    expect(r.exitCode).toBe(0);
    const body = r.calls[0].body as Record<string, { dateTime: string; timeZone: string }>;
    expect(body.startTime).toEqual({ dateTime: '2026-10-20T09:00:00', timeZone: TZ });
    expect(body.endTime.timeZone).toBe(TZ);
  });
});

describe('cal view (M-26, Prefer)', () => {
  it('passes --page-limit through and sends the Prefer header with local-zone bounds', async () => {
    let page = 0;
    const r = await runCommand(registerCalendarCommands,
      ['cal', 'view', '--start', '2026-10-20', '--end', '2026-10-21', '--all', '--page-limit', '2'], {
        env,
        routes: [['GET', /calendarView/, () => ({ value: [{ id: String(++page) }], '@odata.nextLink': `https://graph.microsoft.com/v1.0/me/calendarView?page=${page + 1}` })]],
      });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(2);
    expect(r.calls[0].headers.Prefer).toBe(PREFER);
    expect(query(r.calls[0]).get('startDateTime')).toBe('2026-10-20T07:00:00.000Z');
    expect(query(r.calls[0]).get('endDateTime')).toBe('2026-10-21T07:00:00.000Z');
  });

  it('sends Prefer on events and get', async () => {
    const r1 = await runCommand(registerCalendarCommands, ['cal', 'events'], { env });
    expect(r1.calls[0].headers.Prefer).toBe(PREFER);
    const r2 = await runCommand(registerCalendarCommands, ['cal', 'get', 'E1', '--timezone', 'Europe/London'], { env });
    expect(r2.calls[0].headers.Prefer).toBe('outlook.timezone="Europe/London"');
  });
});

describe('cal decline (M-23)', () => {
  it('rejects --no-notify with a proposed time before any request', async () => {
    const r = await runCommand(registerCalendarCommands,
      ['cal', 'decline', 'E1', '--no-notify', '--propose-start', '2026-10-20T09:00', '--propose-end', '2026-10-20T10:00'], { env });
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('sends proposed times in the resolved zone', async () => {
    const r = await runCommand(registerCalendarCommands,
      ['cal', 'decline', 'E1', '--propose-start', '2026-10-20T09:00', '--propose-end', '2026-10-20T10:00'], { env });
    expect(r.exitCode).toBe(0);
    const body = r.calls[0].body as { sendResponse: boolean; proposedNewTime: Record<string, unknown> };
    expect(body.sendResponse).toBe(true);
    expect(body.proposedNewTime.start).toEqual({ dateTime: '2026-10-20T09:00:00', timeZone: TZ });
    expect(body.proposedNewTime.end).toEqual({ dateTime: '2026-10-20T10:00:00', timeZone: TZ });
  });

  it('requires --propose-end with --propose-start (usage)', async () => {
    const r = await runCommand(registerCalendarCommands, ['cal', 'decline', 'E1', '--propose-start', '2026-10-20T09:00'], { env });
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});

describe('cal rooms / room-lists (M-18)', () => {
  it('uses the v1.0 places endpoints', async () => {
    const r1 = await runCommand(registerCalendarCommands, ['cal', 'rooms'], { env });
    expect(r1.calls[0].path).toBe('/places/microsoft.graph.room');
    const r2 = await runCommand(registerCalendarCommands, ['cal', 'room-lists'], { env });
    expect(r2.calls[0].path).toBe('/places/microsoft.graph.roomlist');
  });

  it('adds an org-mode fix hint on 403', async () => {
    const r = await runCommand(registerCalendarCommands, ['cal', 'rooms'], {
      env,
      routes: [['GET', '/places', () => jsonResponse({ error: { code: 'ErrorAccessDenied', message: 'Access is denied.' } }, 403)]],
    });
    expect(r.exitCode).toBe(4);
    expect(r.stderr).toContain('requires --org-mode login with Place.Read.All');
  });
});
