import { describe, it, expect } from 'vitest';
import { runCommand } from '../helpers/cli.js';
import { registerTodoCommands } from '../../src/commands/todo.js';
import { registerPlannerCommands } from '../../src/commands/planner.js';

const LA = { MS365_CLI_TIMEZONE: 'America/Los_Angeles' };

describe('todo due dates (M-10)', () => {
  it('create --due sends local midnight in the resolved time zone', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'create', 'L1', '--title', 'x', '--due', '2026-10-20', '-o', 'json'], { env: LA });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toMatchObject({ dueDateTime: { dateTime: '2026-10-20T00:00:00', timeZone: 'America/Los_Angeles' } });
  });

  it('update --due accepts natural times and --timezone', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'update', 'L1', 'T1', '--due', '2026-10-21T15:30', '--timezone', 'Europe/Berlin', '-o', 'json'], { env: LA });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0]).toMatchObject({ method: 'PATCH', body: { dueDateTime: { dateTime: '2026-10-21T00:00:00', timeZone: 'Europe/Berlin' } } });
  });

  it('rejects an unparseable --due with exit 2', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'create', 'L1', '--title', 'x', '--due', 'someday']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});

describe('planner due dates (M-10)', () => {
  it('create-task --due sends the UTC instant of local midnight', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'create-task', '--plan', 'P1', '--title', 'x', '--due', '2026-10-20', '-o', 'json'], { env: LA });
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
    const r = await runCommand(registerPlannerCommands, ['planner', 'update-plan', 'P1', '--title', 'y', '-o', 'json'], {
      routes: [['GET', '/planner/plans/P1', () => new Response('{"id":"P1"}', { status: 200, headers: { 'content-type': 'application/json', ETag: 'W/"hdr"' } })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[1].headers['If-Match']).toBe('W/"hdr"');
  });

  it('--dry-run previews the PATCH with a placeholder ETag', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'update-task', 'T1', '--title', 'y', '-o', 'json'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(0);
    expect(r.stdout).toContain('[dry-run] PATCH https://graph.microsoft.com/v1.0/planner/tasks/T1');
    expect(r.stdout).toContain('"If-Match":"W/\\"dry-run\\""');
  });

  it('--dry-run delete-task previews the DELETE', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'delete-task', 'T1'], { dryRun: true, tty: true, answer: 'y' });
    expect(r.exitCode).toBe(0);
    expect(r.stdout).toContain('[dry-run] DELETE https://graph.microsoft.com/v1.0/planner/tasks/T1');
  });

  it('missing ETag fails with a clear message', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'update-bucket', 'B1', '--name', 'n'], {
      routes: [['GET', '/planner/buckets/B1', () => ({ id: 'B1' })]],
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
    const r = await runCommand(registerPlannerCommands, ['planner', 'plan', 'PREMIUM'], {
      routes: [['GET', '/planner/plans/PREMIUM', () => new Response('{"error":{"code":"NotFound","message":"not found"}}', { status: 404 })]],
    });
    expect(r.exitCode).toBe(5);
    expect(r.stderr).toContain('not found');
    expect(r.stderr).toContain("Premium plans (Project for the web) aren't available via the Graph Planner API");
  });

  it('adds the hint to a 403 on a task, keeping exit 4', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'task', 'T9'], {
      routes: [['GET', '/planner/tasks/T9', () => new Response('{"error":{"message":"forbidden"}}', { status: 403 })]],
    });
    expect(r.exitCode).toBe(4);
    expect(r.stderr).toContain('Premium plans');
  });

  it('does not add the hint to other errors', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'plan', 'P1'], {
      routes: [['GET', '/planner/plans/P1', () => new Response('{"error":{"message":"bad"}}', { status: 400 })]],
    });
    expect(r.exitCode).toBe(1);
    expect(r.stderr).not.toContain('Premium');
  });
});
