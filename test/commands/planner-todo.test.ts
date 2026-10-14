import { describe, it, expect } from 'vitest';
import { runCommand } from '../helpers/cli.js';
import { registerTodoCommands } from '../../src/commands/todo.js';
import { registerPlannerCommands } from '../../src/commands/planner.js';

const LA = { MS365_CLI_TIMEZONE: 'America/Los_Angeles' };

describe('todo due dates (M-10)', () => {
  it('create --due sends local midnight in the resolved time zone', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'create', 'AAMkLIST1', '--title', 'x', '--due', '2026-10-20', '-o', 'json'], { env: LA });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toMatchObject({ dueDateTime: { dateTime: '2026-10-20T00:00:00', timeZone: 'America/Los_Angeles' } });
  });

  it('update --due accepts natural times and --timezone', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'update', 'AAMkLIST1', 'T1', '--due', '2026-10-21T15:30', '--timezone', 'Europe/Berlin', '-o', 'json'], { env: LA });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0]).toMatchObject({ method: 'PATCH', body: { dueDateTime: { dateTime: '2026-10-21T00:00:00', timeZone: 'Europe/Berlin' } } });
  });

  it('rejects an unparseable --due with exit 2', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'create', 'AAMkLIST1', '--title', 'x', '--due', 'someday']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});

describe('planner due dates (M-10)', () => {
  it('create-task --due sends the UTC instant of local midnight', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'create-task', '--plan', 'xqQg5FS2LkCp935s-FIFm2QAFkHM', '--title', 'x', '--due', '2026-10-20', '-o', 'json'], { env: LA });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toMatchObject({ dueDateTime: '2026-10-20T07:00:00.000Z' });
  });

  it('update-task --due honours --timezone', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'update-task', 'T1', '--due', '2026-10-20', '--timezone', 'Asia/Tokyo', '--etag', 'W/"e"', '-o', 'json'], { env: LA });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toMatchObject({ dueDateTime: '2026-10-19T15:00:00.000Z' });
  });
});

describe('planner ETag auto-fetch (M-14)', () => {
  it('reads @odata.etag through the client and sends it as If-Match', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'update-task', 'T1', '--title', 'y', '-o', 'json'], {
      apiVersion: 'beta',
      routes: [['GET', '/planner/tasks/T1', () => ({ id: 'T1', '@odata.etag': 'W/"abc"' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls.map(c => c.method)).toEqual(['GET', 'PATCH']);
    expect(r.calls.every(c => c.url.startsWith('https://graph.microsoft.com/beta/'))).toBe(true);
    expect(r.calls[1].headers['If-Match']).toBe('W/"abc"');
  });

  it('falls back to the ETag response header', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'update-plan', 'xqQg5FS2LkCp935s-FIFm2QAFkHM', '--title', 'y', '-o', 'json'], {
      routes: [['GET', '/planner/plans/xqQg5FS2LkCp935s-FIFm2QAFkHM', () => new Response('{"id":"xqQg5FS2LkCp935s-FIFm2QAFkHM"}', { status: 200, headers: { 'content-type': 'application/json', ETag: 'W/"hdr"' } })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[1].headers['If-Match']).toBe('W/"hdr"');
  });

  it('--dry-run previews the PATCH with a placeholder ETag', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'update-task', 'T1', '--title', 'y', '-o', 'json'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toContain('[dry-run] PATCH https://graph.microsoft.com/v1.0/planner/tasks/T1');
    expect(r.stderr).toContain('"If-Match":"W/\\"dry-run\\""');
  });

  it('--dry-run delete-task previews the DELETE', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'delete-task', 'T1'], { dryRun: true, tty: true, answer: 'y' });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain('[dry-run] DELETE https://graph.microsoft.com/v1.0/planner/tasks/T1');
  });

  it('missing ETag fails with a clear message', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'update-bucket', 'hsOf2dhOJkqyYYZEtdzDe2QAIUCR', '--name', 'n'], {
      routes: [['GET', '/planner/buckets/hsOf2dhOJkqyYYZEtdzDe2QAIUCR', () => ({ id: 'hsOf2dhOJkqyYYZEtdzDe2QAIUCR' })]],
    });
    expect(r.exitCode).toBe(1);
    expect(r.stderr).toContain('--etag');
    expect(r.calls).toHaveLength(1);
  });
});

describe('create-plan container (M-27)', () => {
  it('sends container.url instead of owner', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'create-plan', '--title', 'Q4', '--group', 'G1', '-o', 'json']);
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toEqual({ title: 'Q4', container: { url: 'https://graph.microsoft.com/v1.0/groups/G1' } });
  });
});

describe('Planner Premium hint (G-9)', () => {
  it('adds the hint to a 404 on a plan, keeping exit 5', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'plan', 'prEmIuMpLaNiD0123456789abcde'], {
      routes: [['GET', '/planner/plans/prEmIuMpLaNiD0123456789abcde', () => new Response('{"error":{"code":"NotFound","message":"not found"}}', { status: 404 })]],
    });
    expect(r.exitCode).toBe(5);
    expect(r.stderr).toContain('not found');
    expect(r.stderr).toContain("If this is a Premium plan (Project for the web), Graph's Planner API can't reach it.");
  });

  it('adds the hint to a 403 on a task, keeping exit 4', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'task', 'T9'], {
      routes: [['GET', '/planner/tasks/T9', () => new Response('{"error":{"message":"forbidden"}}', { status: 403 })]],
    });
    expect(r.exitCode).toBe(4);
    expect(r.stderr).toContain('If this is a Premium plan');
  });

  it('does not add the hint to other errors', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'plan', 'xqQg5FS2LkCp935s-FIFm2QAFkHM'], {
      routes: [['GET', '/planner/plans/xqQg5FS2LkCp935s-FIFm2QAFkHM', () => new Response('{"error":{"message":"bad"}}', { status: 400 })]],
    });
    expect(r.exitCode).toBe(1);
    expect(r.stderr).not.toContain('Premium');
  });
});

describe('planner usage errors (V-8, V-9, V-10)', () => {
  it.each(['30', '50abc', '', '1e2'])('update-task --percent %j exits 2 with no Graph call', async (pct) => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'update-task', 'T1', '--percent', pct]);
    expect(r.exitCode).toBe(2);
    expect(r.stdout).toBe('');
    expect(r.stderr).toContain('--percent must be 0, 50, or 100');
    expect(r.calls).toHaveLength(0);
  });

  it('update-task --percent 50 sends percentComplete 50', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'update-task', 'T1', '--percent', '50', '--etag', 'W/"e"', '-o', 'json']);
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toEqual({ percentComplete: 50 });
  });

  it.each(['{bad', '{"foo":1}', '[]', '"x"', '{"category1":1}', '{"category26":"x"}', '{"category0":"x"}'])(
    'update-plan-details --categories %s exits 2 before resolving the plan', async (cats) => {
      const r = await runCommand(registerPlannerCommands, ['planner', 'update-plan-details', 'My Plan', '--categories', cats]);
      expect(r.exitCode).toBe(2);
      expect(r.stdout).toBe('');
      expect(r.stderr).toContain('--categories must be a JSON object like {"category1":"Red"}');
      expect(r.calls).toHaveLength(0);
    });

  it('update-plan-details --categories accepts string and null labels', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'update-plan-details', 'xqQg5FS2LkCp935s-FIFm2QAFkHM', '--categories', '{"category1":"Red","category25":null}', '--etag', 'W/"e"', '-o', 'json']);
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toEqual({ categoryDescriptions: { category1: 'Red', category25: null } });
  });

  it('a 404 on a task carries a fix starting "If this is a Premium plan"', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'task', 'WRONGID'], {
      routes: [['GET', '/planner/tasks/WRONGID', () => new Response('{"error":{"code":"NotFound","message":"not found"}}', { status: 404 })]],
    });
    expect(r.exitCode).toBe(5);
    expect(r.stderr).toMatch(/If this is a Premium plan/);
    expect(r.stderr).not.toContain("aren't available");
  });
});

describe('todo lists pages (Q-8)', () => {
  const next = 'https://graph.microsoft.com/v1.0/me/todo/lists?$skiptoken=p2';
  const routes: Parameters<typeof runCommand>[2] = {
    routes: [['GET', '/me/todo/lists', (call) => (call.url.includes('skiptoken') ? { value: [{ id: 'b' }] } : { value: [{ id: 'a' }], '@odata.nextLink': next })]],
  };

  it('page 1 carries nextCursor; --all fetches both pages', async () => {
    const r1 = await runCommand(registerTodoCommands, ['todo', 'lists', '-o', 'json'], routes);
    expect(r1.exitCode).toBe(0);
    expect(JSON.parse(r1.stdout).nextCursor).toBeTruthy();
    const r2 = await runCommand(registerTodoCommands, ['todo', 'lists', '--all', '-o', 'json'], routes);
    expect(JSON.parse(r2.stdout).value.map((v: { id: string }) => v.id)).toEqual(['a', 'b']);
  });
});

describe('todo status lines under --dry-run (DR-1)', () => {
  for (const argv of [['todo', 'delete', 'AAMkLIST1', 'T1'], ['todo', 'check-delete', 'AAMkLIST1', 'T1', 'I1']]) {
    it(`${argv[1]}: no confirmation line, one trailer`, async () => {
      const r = await runCommand(registerTodoCommands, argv, { dryRun: true });
      expect(r.exitCode).toBe(0);
      expect(r.stderr).not.toMatch(/deleted\./);
      expect(r.stderr).toContain('request(s) previewed; nothing was sent.');
    });
  }
});

describe('todo enumerated flags (V-19)', () => {
  it('create --importance rejects an unknown level (exit 2) and lists the allowed values', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'create', 'AAMkLIST1', '--title', 'x', '--importance', 'urgent']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toContain('normal');
  });

  it('create still defaults --importance to normal', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'create', 'AAMkLIST1', '--title', 'x', '-o', 'json']);
    expect(r.calls[0].body).toMatchObject({ importance: 'normal' });
  });

  it('update --status / --importance reject unknown values before any request', async () => {
    const r1 = await runCommand(registerTodoCommands, ['todo', 'update', 'AAMkLIST1', 'T1', '--status', 'done']);
    expect(r1.exitCode).toBe(2);
    expect(r1.calls).toHaveLength(0);
    expect(r1.stderr).toContain('waitingOnOthers');
    const r2 = await runCommand(registerTodoCommands, ['todo', 'update', 'AAMkLIST1', 'T1', '--importance', 'urgent']);
    expect(r2.exitCode).toBe(2);
    const ok = await runCommand(registerTodoCommands, ['todo', 'update', 'AAMkLIST1', 'T1', '--status', 'completed', '--importance', 'high', '-o', 'json']);
    expect(ok.exitCode).toBe(0);
    expect(ok.calls[0].body).toEqual({ status: 'completed', importance: 'high' });
  });
});
