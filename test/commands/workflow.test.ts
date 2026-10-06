import { describe, it, expect } from 'vitest';
import { runCli, type FakeRoute } from '../e2e/run-cli.js';
import { eventInstant } from '../../src/commands/workflow.js';

const LA = 'America/Los_Angeles';
const ENV = {
  MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000',
  MS365_CLI_TOKEN: 'fake',
  TZ: LA,
};

/** Today's date (YYYY-MM-DD) as seen in Los Angeles, offset by `days`. */
function laDay(days = 0): string {
  const today = new Intl.DateTimeFormat('en-CA', { timeZone: LA }).format(new Date());
  const d = new Date(`${today}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

const utcEvent = (subject: string, start: string, end: string, extra: Record<string, unknown> = {}) => ({
  '@odata.etag': 'W/"x"',
  subject,
  start: { dateTime: `${start}.0000000`, timeZone: 'UTC' },
  end: { dateTime: `${end}.0000000`, timeZone: 'UTC' },
  isAllDay: false,
  ...extra,
});

const unread = (n: number): FakeRoute => ({ match: '/me/mailFolders/inbox\\?', method: 'GET', body: { unreadItemCount: n } });

describe('eventInstant (WF-1)', () => {
  it('reads Graph wall-clock times in the zone they were returned in', () => {
    expect(eventInstant({ dateTime: '2026-10-06T16:00:00.0000000', timeZone: 'UTC' }).toISOString()).toBe('2026-10-06T16:00:00.000Z');
    expect(eventInstant({ dateTime: '2026-10-06T09:00:00.0000000', timeZone: LA }).toISOString()).toBe('2026-10-06T16:00:00.000Z');
    expect(eventInstant({ dateTime: '2026-01-06T09:00:00.0000000', timeZone: LA }).toISOString()).toBe('2026-01-06T17:00:00.000Z');
    expect(eventInstant({ dateTime: '2026-10-06T18:00:00.0000000', timeZone: 'Europe/Berlin' }).toISOString()).toBe('2026-10-06T16:00:00.000Z');
    expect(eventInstant({ dateTime: '2026-10-06T18:00:00+02:00', timeZone: 'whatever' }).toISOString()).toBe('2026-10-06T16:00:00.000Z');
  });
});

describe('workflow standup (WF-1, WF-2)', () => {
  it('asks for local-zone times over the local day and shows UTC payload times in local time', async () => {
    const r = await runCli(['workflow', 'standup'], {
      env: ENV,
      routes: [
        { match: '/me/calendarView', method: 'GET', body: { value: [utcEvent('Planning', '2026-10-06T16:00:00', '2026-10-06T16:30:00')] } },
        unread(42),
      ],
    });
    expect(r.stderr).toBe('');
    expect(r.code).toBe(0);
    const view = r.requests.find((q) => q.url.includes('/me/calendarView'))!;
    expect(view.headers.prefer).toBe(`outlook.timezone="${LA}"`);
    const url = new URL(view.url);
    const start = new Date(url.searchParams.get('startDateTime')!);
    const end = new Date(url.searchParams.get('endDateTime')!);
    // Local (LA) midnight is 07:00 or 08:00 UTC, never 00:00 UTC.
    expect([7, 8]).toContain(start.getUTCHours());
    expect(start.getUTCMinutes()).toBe(0);
    expect(new Intl.DateTimeFormat('en-CA', { timeZone: LA }).format(start)).toBe(laDay());
    expect((end.getTime() - start.getTime()) / 3_600_000).toBeGreaterThanOrEqual(23);
    expect(r.stdout).toContain('09:00 AM–09:30 AM  Planning');
    expect(r.stdout).toContain('Unread Emails: 42');
  });

  it('-o json reports the real inbox unread count and no OData annotations', async () => {
    const r = await runCli(['workflow', 'standup', '-o', 'json'], {
      env: ENV,
      routes: [
        { match: '/me/calendarView', method: 'GET', body: { value: [utcEvent('Planning', '2026-10-06T16:00:00', '2026-10-06T16:30:00')] } },
        unread(1234),
      ],
    });
    expect(r.code).toBe(0);
    const out = JSON.parse(r.stdout);
    expect(out.unreadEmails).toBe(1234);
    expect(out.events[0]).not.toHaveProperty('@odata.etag');
    const inbox = r.requests.find((q) => q.url.includes('/me/mailFolders/inbox'))!;
    expect(new URL(inbox.url).searchParams.get('$select')).toBe('unreadItemCount');
    expect(r.requests.some((q) => q.url.includes('/me/messages'))).toBe(false);
  });
});

describe('workflow digest (WF-1, WF-2, WF-3)', () => {
  const lists: FakeRoute = { match: '/me/todo/lists(\\?|$)', method: 'GET', body: { value: [{ id: 'L1', displayName: 'Tasks' }] } };

  it('buckets UTC payload events by the local (LA) day', async () => {
    const r = await runCli(['workflow', 'digest'], {
      env: ENV,
      routes: [
        { match: '/me/calendarView', method: 'GET', body: { value: [
          utcEvent('Late sync', '2026-10-07T02:00:00', '2026-10-07T02:30:00'),
          utcEvent('Morning', '2026-10-07T16:00:00', '2026-10-07T17:00:00'),
          { subject: 'Holiday', start: { dateTime: '2026-10-08T00:00:00.0000000', timeZone: 'UTC' }, end: { dateTime: '2026-10-09T00:00:00.0000000', timeZone: 'UTC' }, isAllDay: true },
        ] } },
        unread(7),
        lists,
        { match: '/me/todo/lists/L1/tasks', method: 'GET', body: { value: [] } },
      ],
    });
    expect(r.stderr).toBe('');
    expect(r.code).toBe(0);
    const view = r.requests.find((q) => q.url.includes('/me/calendarView'))!;
    expect(view.headers.prefer).toBe(`outlook.timezone="${LA}"`);
    const out = r.stdout;
    const tue = out.indexOf('Tue, Oct 6:');
    const wed = out.indexOf('Wed, Oct 7:');
    const thu = out.indexOf('Thu, Oct 8:');
    expect(tue).toBeGreaterThan(-1);
    expect(wed).toBeGreaterThan(tue);
    expect(thu).toBeGreaterThan(wed);
    expect(out.slice(tue, wed)).toContain('07:00 PM–07:30 PM  Late sync');
    expect(out.slice(wed, thu)).toContain('09:00 AM–10:00 AM  Morning');
    expect(out.slice(thu)).toContain('[all day] Holiday');
    expect(out).toContain('Unread Email: 7');
  });

  it('does not call due-today overdue, keeps past tasks out of this week, and reads every page', async () => {
    const due = (day: string) => ({ dateTime: `${day}T00:00:00.0000000`, timeZone: 'UTC' });
    const r = await runCli(['workflow', 'digest', '-o', 'json'], {
      env: ENV,
      routes: [
        { match: '/me/calendarView', method: 'GET', body: { value: [] } },
        unread(0),
        lists,
        { match: '/me/todo/lists/L1/tasks\\?.*skiptoken=2', method: 'GET', body: { value: [
          { title: 'Due today', dueDateTime: due(laDay()) },
        ] } },
        { match: '/me/todo/lists/L1/tasks', method: 'GET', body: {
          value: [
            { title: 'Old', dueDateTime: due(laDay(-10)) },
            { title: 'Far', dueDateTime: due(laDay(30)) },
            { title: 'No date' },
          ],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/todo/lists/L1/tasks?$skiptoken=2',
        } },
      ],
    });
    expect(r.stderr).toBe('');
    expect(r.code).toBe(0);
    const out = JSON.parse(r.stdout);
    expect(out.overdueTasks).toEqual([{ title: 'Old', due: laDay(-10) }]);
    expect(out.tasksDueThisWeek).toBe(1);
  });
});

describe('workflow meeting-prep (WF-3, CORE-4)', () => {
  it('skips all-day and in-progress events', async () => {
    const iso = (offsetMin: number) => new Date(Date.now() + offsetMin * 60_000).toISOString().slice(0, 19);
    const r = await runCli(['workflow', 'meeting-prep', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/me/calendarView', method: 'GET', body: { value: [
        { subject: 'All day', isAllDay: true, start: { dateTime: `${laDay()}T00:00:00.0000000`, timeZone: 'UTC' }, end: { dateTime: `${laDay(1)}T00:00:00.0000000`, timeZone: 'UTC' } },
        utcEvent('In progress', iso(-30), iso(30)),
        utcEvent('Next', iso(60), iso(90)),
      ] } }],
    });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout).subject).toBe('Next');
  });

  it('no upcoming meeting: message on stderr, -o json stdout still parses', async () => {
    const r = await runCli(['workflow', 'meeting-prep', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/me/calendarView', method: 'GET', body: { value: [] } }],
    });
    expect(r.code).toBe(0);
    expect(() => JSON.parse(r.stdout)).not.toThrow();
    expect(r.stderr).toContain('No upcoming meetings');
  });
});

describe('workflow email-to-task (MISC-1, WF-3)', () => {
  it('uses the default To Do list and honours -o', async () => {
    const r = await runCli(['workflow', 'email-to-task', 'M1', '-o', 'yaml'], {
      env: ENV,
      routes: [
        { match: '/me/messages/M1', method: 'GET', body: { subject: 'Hi', from: { emailAddress: { name: 'A', address: 'a@b.cd' } } } },
        { match: '/me/todo/lists(\\?|$)', method: 'GET', body: { value: [{ id: 'DEF', wellknownListName: 'defaultList' }] } },
        { match: '/me/todo/lists/DEF/tasks', method: 'POST', status: 201, body: { id: 'T1', title: '[Email] Hi' } },
      ],
    });
    expect(r.stderr).toBe('');
    expect(r.code).toBe(0);
    const listReq = r.requests.find((q) => q.url.includes('/me/todo/lists?'))!;
    expect(new URL(listReq.url).searchParams.get('$filter')).toBe("wellknownListName eq 'defaultList'");
    expect(r.stdout).toContain('id: T1');
  });

  it('rejects an invalid --importance before any request', async () => {
    const r = await runCli(['workflow', 'email-to-task', 'M1', '--importance', 'urgent'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });
});

describe('workflow focus-time (MISC-1)', () => {
  it.each([['--work-start', 'abc'], ['--work-end', '25:00'], ['--work-start', '9:75']])('rejects %s %s before any request', async (flag, value) => {
    const r = await runCli(['workflow', 'focus-time', flag, value], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
    expect(r.stdout).toBe('');
  });

  it('rejects an end before the start', async () => {
    const r = await runCli(['workflow', 'focus-time', '--work-start', '17:00', '--work-end', '09:00'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });

  it('a multi-day event blocks the whole working day', async () => {
    const r = await runCli(['workflow', 'focus-time', '--work-start', '00:00', '--work-end', '24:00', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/me/calendarView', method: 'GET', body: { value: [
        utcEvent('Offsite', `${laDay(-1)}T12:00:00`, `${laDay(2)}T12:00:00`),
      ] } }],
    });
    expect(r.code).toBe(0);
    const out = JSON.parse(r.stdout);
    expect(out.gaps).toEqual([]);
    expect(out.totalFocusMinutes).toBe(0);
  });
});

describe('workflow file-announce (WF-3, CORE-4)', () => {
  it('posts escaped HTML, not Markdown, and prints the message as data', async () => {
    const r = await runCli(['--org-mode', 'workflow', 'file-announce', 'D1', 'I1', '--team', 'T', '--channel', 'C', '-o', 'json'], {
      env: ENV,
      routes: [
        { match: '/drives/D1/items/I1', method: 'GET', body: { name: 'a<b>&c.txt', webUrl: 'https://x.test/a?b=1&c=2', size: 2048, lastModifiedDateTime: '2026-10-06T00:00:00Z' } },
        { match: '/teams/T/channels/C/messages', method: 'POST', status: 201, body: { id: 'MSG1' } },
      ],
    });
    expect(r.code).toBe(0);
    const post = r.requests.find((q) => q.method === 'POST')!;
    const body = JSON.parse(post.body!).body;
    expect(body.contentType).toBe('html');
    expect(body.content).not.toContain('**');
    expect(body.content).not.toContain('a<b>');
    expect(body.content).toContain('<a href="https://x.test/a?b=1&amp;c=2">a&lt;b&gt;&amp;c.txt</a>');
    expect(JSON.parse(r.stdout).id).toBe('MSG1');
    expect(r.stderr).toContain('Announced');
  });
});
