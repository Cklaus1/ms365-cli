import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { runCommand, type RecordedCall } from '../helpers/cli.js';
import { registerWorkflowCommands } from '../../src/commands/workflow.js';

const TZ = 'America/Los_Angeles';
const env = { MS365_CLI_TIMEZONE: TZ };
const PREFER = `outlook.timezone="${TZ}"`;

/** Graph returns UTC unless the Prefer header asks for the user's zone. */
const at = (call: RecordedCall, laWall: string, utcWall: string) =>
  call.headers.Prefer === PREFER
    ? { dateTime: `${laWall}.0000000`, timeZone: TZ }
    : { dateTime: `${utcWall}.0000000`, timeZone: 'UTC' };

const query = (call: RecordedCall) => new URLSearchParams(call.path.split('?')[1] ?? '');

beforeEach(() => { vi.useFakeTimers({ toFake: ['Date'] }); });
afterEach(() => { vi.useRealTimers(); });

describe('workflow standup', () => {
  it('uses the local day, sends Prefer, prints local times, reads the real unread count (M-7, M-8, M-16)', async () => {
    vi.setSystemTime(new Date('2026-10-15T01:00:00Z')); // Wed Oct 14, 18:00 PDT
    const r = await runCommand(registerWorkflowCommands, ['wf', 'standup'], {
      env,
      routes: [
        ['GET', '/me/calendarView', (c) => ({ value: [{
          subject: 'Sync', isAllDay: false,
          start: at(c, '2026-10-14T09:00:00', '2026-10-14T16:00:00'),
          end: at(c, '2026-10-14T09:30:00', '2026-10-14T16:30:00'),
        }] })],
        ['GET', '/me/mailFolders/inbox', () => ({ unreadItemCount: 250 })],
        ['GET', '/me/messages', () => ({ value: [{ id: 'x' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next' })],
      ],
    });
    expect(r.exitCode).toBe(0);
    const cv = r.calls.find(c => c.path.startsWith('/me/calendarView'))!;
    expect(cv.headers.Prefer).toBe(PREFER);
    expect(query(cv).get('startDateTime')).toBe('2026-10-14T07:00:00.000Z');
    expect(query(cv).get('endDateTime')).toBe('2026-10-15T07:00:00.000Z');
    expect(r.stdout).toContain('09:00 AM–09:30 AM  Sync');
    expect(r.stdout).toContain('Wednesday, October 14, 2026');
    expect(r.stdout).toContain('Unread Emails: 250');
    const inbox = r.calls.find(c => c.path.startsWith('/me/mailFolders/inbox'))!;
    expect(query(inbox).get('$select')).toBe('unreadItemCount');
  });
});

describe('workflow meeting-prep', () => {
  it('skips all-day, already-started and declined events (M-21)', async () => {
    vi.setSystemTime(new Date('2026-10-14T17:00:00Z')); // 10:00 PDT
    const ev = (c: RecordedCall, subject: string, la: [string, string], utc: [string, string], extra: Record<string, unknown> = {}) => ({
      subject, isAllDay: false, responseStatus: { response: 'accepted' },
      start: at(c, la[0], utc[0]), end: at(c, la[1], utc[1]), ...extra,
    });
    const r = await runCommand(registerWorkflowCommands, ['wf', 'meeting-prep'], {
      env,
      routes: [['GET', '/me/calendarView', (c) => ({ value: [
        ev(c, 'Holiday', ['2026-10-14T00:00:00', '2026-10-15T00:00:00'], ['2026-10-14T00:00:00', '2026-10-15T00:00:00'], { isAllDay: true }),
        ev(c, 'Started', ['2026-10-14T09:30:00', '2026-10-14T10:30:00'], ['2026-10-14T16:30:00', '2026-10-14T17:30:00']),
        ev(c, 'Declined', ['2026-10-14T11:00:00', '2026-10-14T11:30:00'], ['2026-10-14T18:00:00', '2026-10-14T18:30:00'], { responseStatus: { response: 'declined' } }),
        ev(c, 'Real', ['2026-10-14T13:00:00', '2026-10-14T14:00:00'], ['2026-10-14T20:00:00', '2026-10-14T21:00:00']),
      ] })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].headers.Prefer).toBe(PREFER);
    expect(r.stdout).toContain('Subject:   Real');
    expect(r.stdout).toContain('Wed, Oct 14, 01:00 PM – 02:00 PM');
  });
});

describe('workflow email-to-task', () => {
  it('sends --due as midnight in the user zone (M-10)', async () => {
    const r = await runCommand(registerWorkflowCommands, ['wf', 'email-to-task', 'M1', '--list', 'L1', '--due', '2026-10-20'], {
      env,
      routes: [
        ['GET', '/me/messages/M1', () => ({ subject: 'Hi', from: { emailAddress: { name: 'A', address: 'a@x.com' } } })],
        ['POST', '/me/todo/lists/L1/tasks', () => ({ id: 'T1', title: '[Email] Hi' })],
      ],
    });
    expect(r.exitCode).toBe(0);
    const post = r.calls.find(c => c.method === 'POST')!;
    expect((post.body as Record<string, unknown>).dueDateTime).toEqual({ dateTime: '2026-10-20T00:00:00', timeZone: TZ });
  });
});

describe('workflow digest', () => {
  it('uses the local Monday–Sunday week, Prefer, local day grouping and unreadItemCount (M-7, M-8, M-16)', async () => {
    vi.setSystemTime(new Date('2026-10-15T01:00:00Z')); // Wed Oct 14, 18:00 PDT (Thu in UTC)
    const r = await runCommand(registerWorkflowCommands, ['wf', 'digest'], {
      env,
      routes: [
        ['GET', '/me/calendarView', (c) => ({ value: [{
          subject: 'Late', isAllDay: false,
          start: at(c, '2026-10-13T20:00:00', '2026-10-14T03:00:00'),
          end: at(c, '2026-10-13T21:00:00', '2026-10-14T04:00:00'),
        }] })],
        ['GET', '/me/mailFolders/inbox', () => ({ unreadItemCount: 250 })],
        ['GET', '/me/messages', () => ({ value: [{ id: 'x' }] })],
        ['GET', '/me/todo/lists', () => ({ value: [] })],
      ],
    });
    expect(r.exitCode).toBe(0);
    const cv = r.calls.find(c => c.path.startsWith('/me/calendarView'))!;
    expect(cv.headers.Prefer).toBe(PREFER);
    expect(query(cv).get('startDateTime')).toBe('2026-10-12T07:00:00.000Z');
    expect(query(cv).get('endDateTime')).toBe('2026-10-19T07:00:00.000Z');
    expect(r.stdout).toContain('Unread Email: 250');
    expect(r.stdout).toMatch(/Tue, Oct 13:\n\s+08:00 PM–09:00 PM  Late/);
  });
});

describe('workflow focus-time', () => {
  it('computes gaps in the user zone (M-7, M-8)', async () => {
    vi.setSystemTime(new Date('2026-10-14T15:00:00Z')); // 08:00 PDT
    const r = await runCommand(registerWorkflowCommands, ['wf', 'focus-time', '-o', 'json'], {
      env,
      routes: [['GET', '/me/calendarView', (c) => ({ value: [{
        subject: 'Standup', isAllDay: false, showAs: 'busy',
        start: at(c, '2026-10-14T09:00:00', '2026-10-14T16:00:00'),
        end: at(c, '2026-10-14T10:00:00', '2026-10-14T17:00:00'),
      }] })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].headers.Prefer).toBe(PREFER);
    expect(JSON.parse(r.stdout).gaps).toEqual([{ start: '10:00', end: '17:00', minutes: 420 }]);
  });
});

describe('workflow file-announce', () => {
  it('posts escaped HTML and reports status on stderr (M-22, C-19)', async () => {
    const r = await runCommand(registerWorkflowCommands, ['wf', 'file-announce', 'D1', 'I1', '--team', 'T', '--channel', 'C'], {
      routes: [
        ['GET', '/drives/D1/items/I1', () => ({ name: 'Q<3> & "R"', webUrl: 'https://x.sharepoint.com/a?b=1&c="2"', size: 2048, lastModifiedDateTime: '2026-10-14T00:00:00Z' })],
        ['POST', '/teams/T/channels/C/messages', () => ({ id: 'm' })],
      ],
    });
    expect(r.exitCode).toBe(0);
    const content = ((r.calls[1].body as Record<string, Record<string, string>>).body).content;
    expect(content).not.toContain('**');
    expect(content).not.toContain('](');
    expect(content).toContain('<b>Q&lt;3&gt; &amp; &quot;R&quot;</b>');
    expect(content).toContain('<a href="https://x.sharepoint.com/a?b=1&amp;c=&quot;2&quot;">Q&lt;3&gt; &amp; &quot;R&quot;</a>');
    expect(r.stdout).toBe('');
    expect(r.stderr).toContain('Announced');
  });

  it('escapes a custom --message', async () => {
    const r = await runCommand(registerWorkflowCommands, ['wf', 'file-announce', 'D1', 'I1', '--team', 'T', '--channel', 'C', '--message', 'see <script>'], {
      routes: [['GET', '/drives/D1/items/I1', () => ({ name: 'a.txt', webUrl: 'https://x/a.txt', size: 1 })]],
    });
    const content = ((r.calls[1].body as Record<string, Record<string, string>>).body).content;
    expect(content).toContain('see &lt;script&gt;');
  });
});

describe('workflow under --dry-run (DR-9…DR-12, F6)', () => {
  const traced = (stderr: string) => stderr.split('\n').filter(l => /^\[dry-run\] (GET|POST) /.test(l));

  it.each([
    [['wf', 'standup'], 2],
    [['wf', 'digest'], 3],
    [['wf', 'focus-time'], 1],
    [['wf', 'meeting-prep'], 1],
    [['wf', 'meeting-prep', '--event-id', 'E1'], 1],
  ])('%j prints the trace only and no report (text)', async (argv, gets) => {
    const r = await runCommand(registerWorkflowCommands, argv, { dryRun: true, env });
    expect(r.exitCode).toBe(0);
    expect(r.stdout).toBe('');
    expect(traced(r.stderr)).toHaveLength(gets);
    expect(r.stderr).not.toMatch(/No upcoming meetings|No focus time/);
    expect(r.stderr).toContain(`${gets} request(s) previewed; nothing was sent.`);
  });

  it.each([
    [['wf', 'standup']], [['wf', 'digest']], [['wf', 'focus-time']], [['wf', 'meeting-prep']],
    [['wf', 'meeting-prep', '--event-id', 'E1']],
  ])('%j -o json prints exactly {"dryRun":true}', async (argv) => {
    const r = await runCommand(registerWorkflowCommands, [...argv, '-o', 'json'], { dryRun: true, env });
    expect(r.exitCode).toBe(0);
    expect(JSON.parse(r.stdout)).toEqual({ dryRun: true });
  });

  it('email-to-task without --list previews the POST to the <first-list> placeholder, body marked <from GET> (DR-10, DR-11)', async () => {
    const r = await runCommand(registerWorkflowCommands, ['wf', 'email-to-task', 'M1', '--due', '2026-10-20', '-o', 'json'], { dryRun: true, env });
    expect(r.exitCode).toBe(0);
    expect(traced(r.stderr)).toEqual([
      expect.stringContaining('GET https://graph.microsoft.com/v1.0/me/messages/M1'),
      expect.stringContaining('GET https://graph.microsoft.com/v1.0/me/todo/lists'),
      expect.stringContaining(`POST https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent('<first-list>')}/tasks`),
    ]);
    const out = JSON.parse(r.stdout);
    expect(out.dryRun).toBe(true);
    expect(out.body.title).toBe('[Email] <from GET>');
    expect(out.body.body.content).toContain('Date: <from GET>');
    expect(out.body.body.content).not.toContain('undefined');
    expect(r.stderr).not.toContain('Task created');
  });

  it('file-announce marks read values <from GET> and prints no "Announced" line (DR-11, DR-1)', async () => {
    const r = await runCommand(registerWorkflowCommands, ['wf', 'file-announce', 'D1', 'I1', '--team', 'T', '--channel', 'C'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    const post = r.stderr.split('\n').find(l => l.startsWith('[dry-run] Body:') || l.includes('"content"'))!;
    expect(post).toContain('<b><from GET></b>');
    expect(post).toContain('Size: <from GET>');
    expect(post).not.toContain('undefined');
    expect(r.stderr).not.toContain('Announced');
  });
});

describe('workflow validation (VT7: V-13, V-19, V-20, V-21)', () => {
  it.each([
    ['--work-start', 'abc'], ['--work-end', '25'], ['--work-start', '9:75'], ['--work-start', '-1'],
  ])('focus-time %s %s exits 2 before any request', async (flag, value) => {
    const r = await runCommand(registerWorkflowCommands, ['wf', 'focus-time', flag, value], { env });
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('focus-time rejects --work-start not before --work-end (exit 2), also under dry-run', async () => {
    const r = await runCommand(registerWorkflowCommands, ['wf', 'focus-time', '--work-start', '17', '--work-end', '9'], { env, dryRun: true });
    expect(r.exitCode).toBe(2);
    expect(r.stderr).not.toContain('[dry-run] GET');
  });

  it('focus-time accepts HH:MM work hours', async () => {
    vi.setSystemTime(new Date('2026-10-14T15:00:00Z')); // 08:00 PDT
    const r = await runCommand(registerWorkflowCommands, ['wf', 'focus-time', '--work-start', '08:30', '--work-end', '12:15', '-o', 'json'], {
      env, routes: [['GET', '/me/calendarView', () => ({ value: [] })]],
    });
    expect(r.exitCode).toBe(0);
    expect(JSON.parse(r.stdout).gaps).toEqual([{ start: '8:30', end: '12:15', minutes: 225 }]);
  });

  it('email-to-task --due 2026-02-30 exits 2 before any request (V-13)', async () => {
    const r = await runCommand(registerWorkflowCommands, ['wf', 'email-to-task', 'M1', '--due', '2026-02-30'], { env });
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('email-to-task --importance outside low|normal|high exits 2 (V-19)', async () => {
    const r = await runCommand(registerWorkflowCommands, ['wf', 'email-to-task', 'M1', '--importance', 'urgent'], { env });
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('email-to-task with no To Do lists exits 5 not_found (V-18)', async () => {
    const r = await runCommand(registerWorkflowCommands, ['wf', 'email-to-task', 'M1'], {
      env, routes: [['GET', '/me/messages/M1', () => ({ subject: 'Hi' })], ['GET', '/me/todo/lists', () => ({ value: [] })]],
    });
    expect(r.exitCode).toBe(5);
    expect(r.calls.some(c => c.method === 'POST')).toBe(false);
  });

  const routes = [
    ['GET', '/me/messages/M1', () => ({ subject: 'Hi' })],
    ['POST', '/me/todo/lists/L1/tasks', () => ({ id: 'T1', title: '[Email] Hi' })],
  ] as never;

  it('email-to-task -o yaml prints YAML, not prose (V-21)', async () => {
    const r = await runCommand(registerWorkflowCommands, ['wf', 'email-to-task', 'M1', '--list', 'L1', '-o', 'yaml'], { env, routes });
    expect(r.exitCode).toBe(0);
    expect(r.stdout).toContain('id: T1');
    expect(r.stdout).not.toContain('Task created');
  });

  it('email-to-task -o text keeps the human status line on stderr (V-21, DR-1)', async () => {
    const r = await runCommand(registerWorkflowCommands, ['wf', 'email-to-task', 'M1', '--list', 'L1', '-o', 'text'], { env, routes });
    expect(r.exitCode).toBe(0);
    expect(r.stdout).toBe('');
    expect(r.stderr).toContain('Task created: "[Email] Hi"');
    expect(r.stderr).toContain('ID: T1');
  });
});
