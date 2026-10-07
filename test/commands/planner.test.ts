import { describe, it, expect } from 'vitest';
import { runCli } from '../e2e/run-cli.js';

const CLIENT = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000' };
const ENV = { ...CLIENT, MS365_CLI_TOKEN: 'fake' };
const TASK = 'https://graph.microsoft.com/v1.0/planner/tasks/T1';

describe('PLANNER-1 ETag lookup goes through the client', () => {
  it('update-task --dry-run with no login/cache previews the PATCH with zero fetch calls', async () => {
    const r = await runCli(['--dry-run', 'planner', 'update-task', 'T1', '--title', 'x'], { env: CLIENT, routes: [] });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toContain(`[dry-run] PATCH ${TASK}`);
    expect(r.stderr).toContain('"If-Match":"<dry-run>"');
    expect(r.stderr).not.toContain('[dry-run] GET');
  });

  it('--read-only update-task and delete-task fail with zero fetch calls', async () => {
    const routes = [{ match: '.', body: { '@odata.etag': 'W/"1"' } }];
    const upd = await runCli(['--read-only', 'planner', 'update-task', 'T1', '--title', 'x'], { env: ENV, routes });
    expect(upd.code).toBe(1);
    expect(upd.stderr).toMatch(/read-only/i);
    expect(upd.requests).toHaveLength(0);

    const del = await runCli(['--read-only', 'planner', 'delete-task', 'T1'], { env: ENV, routes });
    expect(del.code).toBe(1);
    expect(del.stderr).toMatch(/read-only/i);
    expect(del.requests).toHaveLength(0);
  });

  it('fetches @odata.etag with the client and sends it as If-Match', async () => {
    const r = await runCli(['planner', 'update-task', 'T1', '--percent', '50', '-o', 'json'], {
      env: ENV,
      routes: [
        { match: '/planner/tasks/T1$', method: 'GET', body: { '@odata.etag': 'W/"abc"', id: 'T1' } },
        { match: '/planner/tasks/T1$', method: 'PATCH', status: 204 },
      ],
    });
    expect(r.code).toBe(0);
    expect(r.requests.map((q) => `${q.method} ${q.url}`)).toEqual([`GET ${TASK}`, `PATCH ${TASK}`]);
    expect(r.requests[1].headers['if-match']).toBe('W/"abc"');
    expect(r.requests[0].headers.authorization).toBe('Bearer fake');
    expect(JSON.parse(r.requests[1].body!)).toEqual({ percentComplete: 50 });
    expect(() => JSON.parse(r.stdout)).not.toThrow();
  });

  it('an explicit --etag skips the lookup', async () => {
    const r = await runCli(['planner', 'update-bucket', 'B1', '--name', 'n', '--etag', 'W/"e"', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/planner/buckets/B1$', method: 'PATCH', status: 204 }],
    });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(1);
    expect(r.requests[0].headers['if-match']).toBe('W/"e"');
  });
});

describe('TODO-1 planner: no-field updates send nothing', () => {
  for (const args of [
    ['update-task', 'T1'],
    ['update-plan', 'P1'],
    ['update-bucket', 'B1'],
    ['update-plan-details', 'P1'],
    ['update-task-details', 'T1'],
  ]) {
    it(`${args[0]} without fields → nothing to update`, async () => {
      const r = await runCli(['planner', ...args, '--etag', 'W/"e"'], { env: ENV, routes: [{ match: '.', body: {} }] });
      expect(r.code).toBe(1);
      expect(r.stderr).toMatch(/nothing to update/i);
      expect(r.requests).toHaveLength(0);
    });
  }
});

describe('PLANNER-2', () => {
  it('create-plan sends container, not the deprecated owner', async () => {
    const r = await runCli(['planner', 'create-plan', '--title', 'Q4', '--group', 'G1', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/planner/plans$', method: 'POST', status: 201, body: { id: 'P1', title: 'Q4' } }],
    });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.requests[0].body!)).toEqual({ title: 'Q4', container: { containerId: 'G1', type: 'group' } });
    expect(JSON.parse(r.stdout)).toEqual({ id: 'P1', title: 'Q4' });
  });

  it('--due is the start of that day in the local time zone, not UTC midnight', async () => {
    const r = await runCli(['--dry-run', 'planner', 'create-task', '--plan', 'P1', '--title', 't', '--due', '2026-10-07', '-o', 'json'], {
      env: { ...CLIENT, TZ: 'America/New_York' },
      routes: [],
    });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout).body.dueDateTime).toBe('2026-10-07T04:00:00.000Z');
  });

  it.each([
    ['create-task', '--plan', 'P1', '--title', 't'],
    ['update-task', 'T1', '--etag', 'W/"e"'],
  ])('--due rejects an impossible calendar date (planner %s) with no request', async (...args) => {
    for (const due of ['2026-02-30', '2026-04-31', '2026-13-01']) {
      const r = await runCli(['planner', ...args, '--due', due], { env: { ...ENV, TZ: 'America/New_York' }, routes: [] });
      expect(r.code).toBe(1);
      expect(r.requests).toHaveLength(0);
      expect(r.stderr).toContain(`Invalid date: "${due}". Use YYYY-MM-DD format.`);
    }
  });

  it('--categories with invalid JSON fails with a clear message and no request', async () => {
    const r = await runCli(['planner', 'update-plan-details', 'P1', '--etag', 'W/"e"', '--categories', '{bad'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('--categories');
    expect(r.requests).toHaveLength(0);
  });
});

describe('PAGE-2 planner collections drain every page', () => {
  const cases: Array<[string[], string]> = [
    [['plans'], '/me/planner/plans'],
    [['my-tasks'], '/me/planner/tasks'],
    [['tasks', 'P1'], '/planner/plans/P1/tasks'],
    [['buckets', 'P1'], '/planner/plans/P1/buckets'],
    [['bucket-tasks', 'B1'], '/planner/buckets/B1/tasks'],
  ];
  for (const [args, endpoint] of cases) {
    it(`${args[0]} prints one combined JSON array from two pages`, async () => {
      const r = await runCli(['planner', ...args, '-o', 'json'], {
        env: ENV,
        routes: [
          { match: 'skiptoken=2', body: { value: [{ id: 'b' }] } },
          { match: endpoint, body: { value: [{ id: 'a' }], '@odata.nextLink': `https://graph.microsoft.com/v1.0${endpoint}?$skiptoken=2` } },
        ],
      });
      expect(r.code).toBe(0);
      expect(JSON.parse(r.stdout)).toEqual([{ id: 'a' }, { id: 'b' }]);
      expect(r.requests).toHaveLength(2);
    });
  }
});
