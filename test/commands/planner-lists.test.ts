import { describe, it, expect, afterEach } from 'vitest';
import { runCommand } from '../helpers/cli.js';
import { registerPlannerCommands } from '../../src/commands/planner.js';
import { setCursor, decodeCursor } from '../../src/formatter.js';

const PLAN = 'xqQg5FS2LkCp935s-FIFm2QAFkHM';
const BUCKET = 'hsOf2dhOJkqyYYZEtdzDe2QAIUCR';

afterEach(() => { setCursor(null); });

/** Each list command and the collection path it reads. */
const LISTS: Array<[argv: string[], path: string]> = [
  [['planner', 'plans'], '/me/planner/plans'],
  [['planner', 'my-tasks'], '/me/planner/tasks'],
  [['planner', 'tasks', PLAN], `/planner/plans/${PLAN}/tasks`],
  [['planner', 'buckets', PLAN], `/planner/plans/${PLAN}/buckets`],
  [['planner', 'bucket-tasks', BUCKET], `/planner/buckets/${BUCKET}/tasks`],
];

const nextLinkFor = (path: string) => `https://graph.microsoft.com/v1.0${path}?$skiptoken=p2`;

describe('planner list commands page (Q-8, AC-2)', () => {
  it.each(LISTS)('%j without flags prints page 1 with nextCursor', async (argv, path) => {
    const r = await runCommand(registerPlannerCommands, [...argv, '-o', 'json'], {
      routes: [['GET', path, () => ({ value: [{ id: 'a' }], '@odata.nextLink': nextLinkFor(path) })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(1);
    const out = JSON.parse(r.stdout);
    expect(out.value).toEqual([{ id: 'a' }]);
    expect(decodeCursor(out.nextCursor)).toBe(nextLinkFor(path));
  });

  it.each(LISTS)('%j --all prints both pages in one envelope', async (argv, path) => {
    const r = await runCommand(registerPlannerCommands, [...argv, '--all', '-o', 'json'], {
      routes: [
        ['GET', /skiptoken=p2/, () => ({ value: [{ id: 'b' }] })],
        ['GET', path, () => ({ value: [{ id: 'a' }], '@odata.nextLink': nextLinkFor(path) })],
      ],
    });
    expect(r.exitCode).toBe(0);
    expect(JSON.parse(r.stdout)).toEqual({ value: [{ id: 'a' }, { id: 'b' }] });
  });

  it.each(LISTS)('%j honours --cursor', async (argv, path) => {
    setCursor(Buffer.from(nextLinkFor(path)).toString('base64url'));
    const r = await runCommand(registerPlannerCommands, [...argv, '-o', 'json'], {
      routes: [['GET', /skiptoken=p2/, () => ({ value: [{ id: 'b' }] })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls.map(c => c.url)).toEqual([nextLinkFor(path)]);
    expect(JSON.parse(r.stdout).value).toEqual([{ id: 'b' }]);
  });

  it('table output prints nextCursor on stderr', async () => {
    const path = '/me/planner/tasks';
    const r = await runCommand(registerPlannerCommands, ['planner', 'my-tasks', '-o', 'table'], {
      routes: [['GET', path, () => ({ value: [{ id: 'a' }], '@odata.nextLink': nextLinkFor(path) })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain('nextCursor:');
  });
});

describe('planner list commands keep the Premium hint (Q-8 constraint 1)', () => {
  it.each([['tasks'], ['buckets']])('planner %s on a 403 still carries the Premium fix', async (cmd) => {
    const r = await runCommand(registerPlannerCommands, ['planner', cmd, PLAN], {
      routes: [['GET', `/planner/plans/${PLAN}/${cmd}`, () => new Response('{"error":{"message":"forbidden"}}', { status: 403 })]],
    });
    expect(r.exitCode).toBe(4);
    expect(r.stderr).toContain('If this is a Premium plan');
  });

  it('planner tasks --all on a 404 still carries the Premium fix', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'tasks', PLAN, '--all'], {
      routes: [['GET', `/planner/plans/${PLAN}/tasks`, () => new Response('{"error":{"code":"NotFound","message":"nf"}}', { status: 404 })]],
    });
    expect(r.exitCode).toBe(5);
    expect(r.stderr).toContain('If this is a Premium plan');
  });
});

describe('planner delete confirmations use success() (DR-1)', () => {
  it.each([
    [['planner', 'delete-task', 'T1', '--etag', 'W/"e"'], 'Task deleted.'],
    [['planner', 'delete-plan', PLAN, '--etag', 'W/"e"'], 'Plan deleted.'],
    [['planner', 'delete-bucket', BUCKET, '--etag', 'W/"e"'], 'Bucket deleted.'],
  ])('%j prints its confirmation live and stays silent under --dry-run', async (argv, msg) => {
    const live = await runCommand(registerPlannerCommands, argv, { tty: true, answer: 'y', routes: [['DELETE', /\/planner\//, () => new Response(null, { status: 204 })]] });
    expect(live.exitCode).toBe(0);
    expect(live.stderr).toContain(msg);

    const dry = await runCommand(registerPlannerCommands, argv, { dryRun: true, tty: true, answer: 'y' });
    expect(dry.exitCode).toBe(0);
    expect(dry.stderr).toContain('[dry-run] DELETE');
    expect(dry.stderr).not.toContain(msg);
  });
});
