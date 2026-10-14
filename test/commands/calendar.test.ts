import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { runCommand, jsonResponse, type RecordedCall, type Route } from '../helpers/cli.js';
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
    expect(r.exitCode).toBe(2); // V-11: invalid input is a usage error
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

describe('cal create --teams (PM-4)', () => {
  it('marks the event as a Teams online meeting', async () => {
    const r = await runCommand(registerCalendarCommands,
      ['cal', 'create', '--subject', 'S', '--start', '2026-10-20T09:00', '--end', '2026-10-20T09:30', '--teams'], { env });
    expect(r.exitCode).toBe(0);
    const body = r.calls[0].body as Record<string, unknown>;
    expect(body.isOnlineMeeting).toBe(true);
    expect(body.onlineMeetingProvider).toBe('teamsForBusiness');
  });

  it('leaves online-meeting fields out without --teams', async () => {
    const r = await runCommand(registerCalendarCommands,
      ['cal', 'create', '--subject', 'S', '--start', '2026-10-20T09:00', '--end', '2026-10-20T09:30'], { env });
    expect(r.calls[0].body).not.toHaveProperty('isOnlineMeeting');
  });
});

describe('cal find-time (PM-4)', () => {
  const suggestions = () => ({
    emptySuggestionsReason: '',
    meetingTimeSuggestions: [{
      confidence: 100,
      organizerAvailability: 'free',
      meetingTimeSlot: {
        start: { dateTime: '2026-10-15T16:00:00.0000000', timeZone: 'UTC' },
        end: { dateTime: '2026-10-15T16:30:00.0000000', timeZone: 'UTC' },
      },
      attendeeAvailability: [{ attendee: { type: 'required', emailAddress: { address: 'a@x.com' } }, availability: 'free' }],
    }],
  });

  it('posts findMeetingTimes and prints suggestions in the user zone', async () => {
    const r = await runCommand(registerCalendarCommands,
      ['cal', 'find-time', '--attendees', 'a@x.com, b@x.com', '--duration', '30m', '--start', '2026-10-15T09:00', '--end', '2026-10-15T17:00'], {
        env, routes: [['POST', '/me/findMeetingTimes', suggestions]],
      });
    expect(r.exitCode).toBe(0);
    const call = r.calls[0];
    expect(call.headers.Prefer).toBe(PREFER);
    expect(call.body).toEqual({
      attendees: [
        { type: 'required', emailAddress: { address: 'a@x.com' } },
        { type: 'required', emailAddress: { address: 'b@x.com' } },
      ],
      timeConstraint: {
        activityDomain: 'work',
        timeSlots: [{ start: { dateTime: '2026-10-15T09:00:00', timeZone: TZ }, end: { dateTime: '2026-10-15T17:00:00', timeZone: TZ } }],
      },
      meetingDuration: 'PT30M',
    });
    expect(JSON.parse(r.stdout)).toEqual({ value: [{
      start: '2026-10-15T09:00:00-07:00',
      end: '2026-10-15T09:30:00-07:00',
      confidence: 100,
      organizerAvailability: 'free',
      attendeeAvailability: [{ attendee: 'a@x.com', availability: 'free' }],
    }] });
  });

  it('converts hour and mixed durations; defaults to the next 7 days', async () => {
    const r = await runCommand(registerCalendarCommands, ['cal', 'find-time', '--attendees', 'a@x.com', '--duration', '1h30m'], { env });
    const body = r.calls[0].body as { meetingDuration: string; timeConstraint: { timeSlots: Array<{ start: { dateTime: string }; end: { dateTime: string } }> } };
    expect(body.meetingDuration).toBe('PT1H30M');
    expect(body.timeConstraint.timeSlots[0].start.dateTime).toBe('2026-10-14T10:00:00');
    expect(body.timeConstraint.timeSlots[0].end.dateTime).toBe('2026-10-21T10:00:00');
  });

  it('reports why there are no suggestions', async () => {
    const r = await runCommand(registerCalendarCommands, ['cal', 'find-time', '--attendees', 'a@x.com', '--duration', '1h'], {
      env, routes: [['POST', '/me/findMeetingTimes', () => ({ emptySuggestionsReason: 'attendeesUnavailable', meetingTimeSuggestions: [] })]],
    });
    expect(JSON.parse(r.stdout)).toEqual({ value: [], emptySuggestionsReason: 'attendeesUnavailable' });
  });

  it('rejects a bad duration or attendee before sending', async () => {
    const bad = await runCommand(registerCalendarCommands, ['cal', 'find-time', '--attendees', 'a@x.com', '--duration', 'soon'], { env });
    expect(bad.exitCode).toBe(2);
    expect(bad.calls).toHaveLength(0);
    const badMail = await runCommand(registerCalendarCommands, ['cal', 'find-time', '--attendees', 'nope', '--duration', '1h'], { env });
    expect(badMail.exitCode).not.toBe(0);
    expect(badMail.calls).toHaveLength(0);
  });
});

describe('collection commands page (Q-8, Q-20)', () => {
  const twoPages = (prefix: string): Route => {
    const next = `https://graph.microsoft.com/v1.0${prefix}?$skiptoken=p2`;
    return ['GET', prefix, (call) => (call.url.includes('skiptoken') ? { value: [{ id: 'b' }] } : { value: [{ id: 'a' }], '@odata.nextLink': next })];
  };
  const cases: Array<[string[], string]> = [
    [['cal', 'list'], '/me/calendars'],
    [['cal', 'rooms'], '/places/microsoft.graph.room'],
    [['cal', 'room-lists'], '/places/microsoft.graph.roomlist'],
  ];

  for (const [argv, prefix] of cases) {
    it(`${argv.join(' ')}: page 1 carries nextCursor; --all fetches both pages`, async () => {
      const r1 = await runCommand(registerCalendarCommands, [...argv, '-o', 'json'], { env, routes: [twoPages(prefix)] });
      expect(r1.exitCode).toBe(0);
      expect(JSON.parse(r1.stdout).nextCursor).toBeTruthy();
      const r2 = await runCommand(registerCalendarCommands, [...argv, '--all', '-o', 'json'], { env, routes: [twoPages(prefix)] });
      expect(r2.exitCode).toBe(0);
      expect(JSON.parse(r2.stdout).value.map((v: { id: string }) => v.id)).toEqual(['a', 'b']);
    });
  }

  it('cal room-lists keeps the Place.Read.All hint on a 403', async () => {
    const r = await runCommand(registerCalendarCommands, ['cal', 'room-lists', '--all'], {
      env,
      routes: [['GET', '/places', () => jsonResponse({ error: { code: 'ErrorAccessDenied', message: 'Access is denied.' } }, 403)]],
    });
    expect(r.exitCode).toBe(4);
    expect(r.stderr).toContain('requires --org-mode login with Place.Read.All');
  });
});

describe('cal free-busy validation (V-12)', () => {
  const base = ['cal', 'free-busy', '--start', '2026-10-20T09:00', '--end', '2026-10-20T17:00'];

  it('rejects an invalid --emails entry before any request (exit 2)', async () => {
    const r = await runCommand(registerCalendarCommands, [...base, '--emails', 'a@x.com,notanemail'], { env });
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toContain('notanemail');
  });

  it('sends trimmed, validated addresses and an integer interval', async () => {
    const r = await runCommand(registerCalendarCommands, [...base, '--emails', 'a@x.com, b@y.com', '--interval', '15'], { env });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toMatchObject({ schedules: ['a@x.com', 'b@y.com'], availabilityViewInterval: 15 });
  });

  for (const bad of ['abc', '4', '1441', '7.5']) {
    it(`rejects --interval ${bad} (exit 2)`, async () => {
      const r = await runCommand(registerCalendarCommands, [...base, '--emails', 'a@x.com', '--interval', bad], { env });
      expect(r.exitCode).toBe(2);
      expect(r.calls).toHaveLength(0);
    });
  }
});

describe('cal update --dry-run zone marker (DR-17, F5)', () => {
  it('previews the zone GET and marks the PATCH zone as <from GET>', async () => {
    const r = await runCommand(registerCalendarCommands, ['cal', 'update', 'E1', '--start', '2026-10-20T09:00', '--end', '2026-10-20T10:00', '-o', 'json'], { env, dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain('GET https://graph.microsoft.com/v1.0/me/events/E1?$select=originalStartTimeZone');
    const body = JSON.parse(r.stdout).body as Record<string, { dateTime: string; timeZone: string }>;
    expect(body.start).toEqual({ dateTime: '2026-10-20T09:00:00', timeZone: '<from GET>' });
    expect(body.end.timeZone).toBe('<from GET>');
    expect(r.stdout).not.toContain(TZ);
  });

  it('with --timezone there is no GET and no marker', async () => {
    const r = await runCommand(registerCalendarCommands, ['cal', 'update', 'E1', '--start', '2026-10-20T09:00', '--timezone', 'Europe/London', '-o', 'json'], { env, dryRun: true });
    expect(r.stderr).not.toContain('$select=originalStartTimeZone');
    expect(JSON.parse(r.stdout).body.start.timeZone).toBe('Europe/London');
  });
});

describe('cal status lines under --dry-run (DR-1)', () => {
  const cases: string[][] = [
    ['cal', 'delete', 'E1'],
    ['cal', 'accept', 'E1'],
    ['cal', 'decline', 'E1'],
    ['cal', 'decline', 'E1', '--propose-start', '2026-10-20T09:00', '--propose-end', '2026-10-20T10:00'],
    ['cal', 'tentative', 'E1'],
  ];
  for (const argv of cases) {
    it(`${argv.slice(0, 2).join(' ')}${argv.length > 4 ? ' (proposed)' : ''}: no confirmation line, one trailer`, async () => {
      const r = await runCommand(registerCalendarCommands, argv, { env, dryRun: true });
      expect(r.exitCode).toBe(0);
      expect(r.stderr).not.toMatch(/Event (deleted|accepted|declined)|Tentatively accepted/);
      expect(r.stderr).toContain('request(s) previewed; nothing was sent.');
    });
  }

  it('prints the confirmation on a real run', async () => {
    const r = await runCommand(registerCalendarCommands, ['cal', 'accept', 'E1'], { env });
    expect(r.stderr).toContain('Event accepted.');
  });
});

describe('cal enumerated flags (V-19)', () => {
  it('create-calendar --color rejects an unknown colour (exit 2)', async () => {
    const r = await runCommand(registerCalendarCommands, ['cal', 'create-calendar', '--name', 'N', '--color', 'purple'], { env });
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toContain('lightBlue');
  });

  it('update-calendar --color rejects an unknown colour and accepts a listed one', async () => {
    const bad = await runCommand(registerCalendarCommands, ['cal', 'update-calendar', 'C1', '--color', 'purple'], { env });
    expect(bad.exitCode).toBe(2);
    const ok = await runCommand(registerCalendarCommands, ['cal', 'update-calendar', 'C1', '--color', 'lightTeal'], { env });
    expect(ok.exitCode).toBe(0);
    expect(ok.calls[0].body).toEqual({ color: 'lightTeal' });
  });
});
