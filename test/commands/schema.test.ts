import { describe, it, expect, beforeAll } from 'vitest';
import { runCli } from '../e2e/run-cli.js';
import { buildCommandTree, findCommand } from '../../src/command-tree.js';
import { introspectionProgram } from '../../src/commands/registry.js';

describe('SCHEMA-1 schema --list and unknown endpoints', () => {
  it('schema --list exits 0 and prints a JSON list of endpoints', async () => {
    const r = await runCli(['schema', '--list']);
    expect(r.code).toBe(0);
    const list = JSON.parse(r.stdout) as Array<{ endpoint: string; method: string; path: string }>;
    expect(list.find(e => e.endpoint === 'mail.list')).toMatchObject({ method: 'GET', path: '/me/messages' });
  });

  it('schema --list honours -o', async () => {
    const r = await runCli(['schema', '--list', '-o', 'csv']);
    expect(r.code).toBe(0);
    expect(r.stdout.split('\n')[0]).toContain('endpoint');
  });

  it('schema <endpoint> prints the endpoint schema as JSON', async () => {
    const r = await runCli(['schema', 'mail.send']);
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout)).toMatchObject({ httpMethod: 'POST', path: '/me/sendMail' });
  });

  it('an unknown endpoint exits 1 with suggestions from the same service', async () => {
    const r = await runCli(['schema', 'mail.lst']);
    expect(r.code).toBe(1);
    expect(r.stdout).toBe('');
    expect(r.stderr).toContain('Unknown endpoint: mail.lst');
    expect(r.stderr).toMatch(/Did you mean: .*mail\.list/);
    expect(r.stderr).toContain('ms365 schema --list');
  });

  it('a misspelled service still gets close suggestions', async () => {
    const r = await runCli(['schema', 'contcts.list']);
    expect(r.code).toBe(1);
    expect(r.stderr).toMatch(/Did you mean: contacts\.list/);
  });

  it('no endpoint and no --list exits 1', async () => {
    const r = await runCli(['schema']);
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('schema --list');
  });
});

/** Endpoints with hand-written parameter/response detail before SCHEMA-2. */
const HAND_WRITTEN = [
  'mail.list', 'mail.get', 'mail.send', 'calendar.events', 'calendar.view', 'calendar.create', 'drive.list', 'drive.files',
  'drive.upload', 'todo.lists', 'todo.tasks', 'todo.create', 'contacts.list', 'teams.list', 'chat.send', 'planner.tasks',
  'search.query', 'mail.triage', 'mail.download-attachment', 'mail.create-rule', 'calendar.rooms', 'calendar.room-lists',
  'drive.cat', 'drive.cp', 'excel.format-range', 'excel.sort-range', 'sharepoint.delta', 'planner.create-plan', 'onenote.page',
  'chat.reply', 'user.set-ooo', 'user.set-presence', 'user.clear-presence', 'user.download-photo', 'groups.distribution-lists',
];

describe('SCHEMA-2 catalog generated from the command tree', () => {
  const tree = buildCommandTree(introspectionProgram());
  let rows: Array<{ endpoint: string; method: string; path: string; description: string }>;

  beforeAll(async () => {
    const r = await runCli(['schema', '--list']);
    expect(r.code).toBe(0);
    rows = JSON.parse(r.stdout);
  });

  it('--list has one row per annotated request of every Graph-calling leaf, and nothing else', () => {
    const expected = tree.leaves.flatMap(l => Array.isArray(l.graph)
      ? l.graph.map(g => ({ endpoint: l.path, method: g.method, path: g.path }))
      : []);
    expect(rows.map(({ endpoint, method, path }) => ({ endpoint, method, path }))).toEqual(expected);
    expect(rows.filter(r => findCommand(tree, r.endpoint)?.path !== r.endpoint)).toEqual([]);
    expect(rows.some(r => r.endpoint.startsWith('auth.') || ['schema', 'generate-skills', 'serve'].includes(r.endpoint))).toBe(false);
  });

  it('keeps the hand-written detail, whose request is one the command makes', async () => {
    const results = await Promise.all(HAND_WRITTEN.map(key => runCli(['schema', key])));
    const wrong = results.flatMap((r, i) => {
      const key = HAND_WRITTEN[i];
      if (r.code !== 0) return [`${key}: exit ${r.code}`];
      const out = JSON.parse(r.stdout);
      const leaf = findCommand(tree, key)!;
      const made = Array.isArray(leaf.graph) && leaf.graph.some(g => g.method === out.httpMethod && g.path === out.path);
      return made && out.scopes && out.parameters ? [] : [`${key}: ${out.httpMethod} ${out.path}`];
    });
    expect(wrong).toEqual([]);
  }, 30_000);

  it('a single-request command without hand-written detail shows its request', async () => {
    const r = await runCli(['schema', 'teams.get']);
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout)).toEqual({ httpMethod: 'GET', path: '/teams/{team-id}', description: 'Get team details' });
  });

  it('a multi-request command lists every request in order', async () => {
    const r = await runCli(['schema', 'planner.update-task']);
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout).requests).toEqual([
      { httpMethod: 'GET', path: '/planner/tasks/{task-id}' },
      { httpMethod: 'PATCH', path: '/planner/tasks/{task-id}' },
    ]);
  });

  it('accepts group and command aliases', async () => {
    const r = await runCli(['schema', 'cal.view']);
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout)).toMatchObject({ httpMethod: 'GET', path: '/me/calendarView' });
  });

  it('commands without Graph calls have no entry', async () => {
    const r = await runCli(['schema', 'auth.login']);
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('Unknown endpoint: auth.login');
  });

  it('chat send/reply describe the created message printed on stdout', async () => {
    for (const key of ['chat.send', 'chat.reply']) {
      const r = await runCli(['schema', key]);
      expect(JSON.parse(r.stdout).response).toBe('chatMessage (created message, printed on stdout)');
    }
  });
});
