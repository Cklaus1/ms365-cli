import { describe, it, expect } from 'vitest';
import { runCli } from '../e2e/run-cli.js';

const ENV = {
  MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000',
  MS365_CLI_TOKEN: 'fake',
};

describe('todo option validation (TODO-1)', () => {
  it.each([
    ['create', 'L1', '--title', 't', '--importance', 'urgent'],
    ['update', 'L1', 'T1', '--importance', 'HIGH'],
    ['update', 'L1', 'T1', '--status', 'done'],
  ])('rejects a bad choice before any request: %s', async (...args) => {
    const r = await runCli(['todo', ...args], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toMatch(/Allowed choices/);
  });

  it('update with no fields: no request, exit 1', async () => {
    const r = await runCli(['todo', 'update', 'L1', 'T1'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toMatch(/nothing to update/i);
  });

  it('update sends the valid choices', async () => {
    const r = await runCli(['todo', 'update', 'L1', 'T1', '--status', 'completed', '--importance', 'high', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/me/todo/lists/L1/tasks/T1$', method: 'PATCH', body: { id: 'T1' } }],
    });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.requests[0].body!)).toEqual({ status: 'completed', importance: 'high' });
  });

  it.each([
    ['create', 'L1', '--title', 't', '--due', '2026-02-30'],
    ['update', 'L1', 'T1', '--due', '2026-04-31'],
    ['update', 'L1', 'T1', '--due', '2026-12'],
  ])('rejects a --due that is not a real YYYY-MM-DD day, before any request: %s', async (...args) => {
    const r = await runCli(['todo', ...args], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toContain('Use YYYY-MM-DD format');
  });
});

describe('todo lists (PAGE-1, CORE-4)', () => {
  it('tasks honours --top and --page-limit; empty -o json parses', async () => {
    const r = await runCli(['todo', 'tasks', 'L1', '--top', '3', '--all', '--page-limit', '1', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/me/todo/lists/L1/tasks', method: 'GET', body: { value: [], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/todo/lists/L1/tasks?$skiptoken=n' } }],
    });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(1);
    expect(new URL(r.requests[0].url).searchParams.get('$top')).toBe('3');
    expect(JSON.parse(r.stdout)).toEqual([]);
  });
});

describe('todo due dates (TZ-1)', () => {
  it.each([
    ['Asia/Tokyo', 'Asia/Tokyo'],
    ['Europe/Amsterdam', 'UTC'],
  ])('TZ=%s: --due is local midnight in %s (create, update)', async (tz, zone) => {
    const env = { ...ENV, TZ: tz };
    const due = { dateTime: '2026-10-07T00:00:00', timeZone: zone };
    const created = await runCli(['todo', 'create', 'L1', '--title', 't', '--due', '2026-10-07', '-o', 'json'], {
      env, routes: [{ match: '/me/todo/lists/L1/tasks$', method: 'POST', status: 201, body: { id: 'T1' } }],
    });
    expect(created.code).toBe(0);
    expect(JSON.parse(created.requests[0].body!).dueDateTime).toEqual(due);
    const updated = await runCli(['todo', 'update', 'L1', 'T1', '--due', '2026-10-07', '-o', 'json'], {
      env, routes: [{ match: '/me/todo/lists/L1/tasks/T1$', method: 'PATCH', body: { id: 'T1' } }],
    });
    expect(updated.code).toBe(0);
    expect(JSON.parse(updated.requests[0].body!)).toEqual({ dueDateTime: due });
  });
});
