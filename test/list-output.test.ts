import { describe, it, expect, afterEach } from 'vitest';
import { formatOutput, streamPages, setFields, setCursor, encodeCursor, decodeCursor, projectItem } from '../src/formatter.js';
import { runCommand } from './helpers/cli.js';
import { registerTodoCommands } from '../src/commands/todo.js';
import { CliError, ExitCode } from '../src/errors.js';

afterEach(() => { setFields(null); setCursor(null); });

const items = [
  { id: '1', subject: 'a', from: { emailAddress: { address: 'x@c.com', name: 'X' } } },
  { id: '2', subject: 'b', from: { emailAddress: { address: 'y@c.com', name: 'Y' } } },
];

describe('ndjson and --fields (P-6, D18)', () => {
  it('ndjson prints one compact JSON object per line', () => {
    const out = formatOutput({ value: items }, 'ndjson');
    expect(out.split('\n').map(l => JSON.parse(l).id)).toEqual(['1', '2']);
  });

  it('--fields projects nested dot paths and keeps structure', () => {
    setFields(['id', 'from.emailAddress.address']);
    expect(JSON.parse(formatOutput({ value: items }, 'json')).value[0]).toEqual({ id: '1', from: { emailAddress: { address: 'x@c.com' } } });
    expect(formatOutput({ value: items }, 'table')).toContain('from.emailAddress.address');
    expect(formatOutput({ value: items }, 'table')).not.toContain('subject');
    expect(projectItem({ id: 9, other: 1 })).toEqual({ id: 9 });
  });

  it('streamed json with wrapper equals the buffered {value: [...]} shape', async () => {
    async function* pages() { yield [items[0]]; yield [items[1]]; }
    let out = '';
    await streamPages(pages(), 'json', (c) => { out += c; }, { jsonWrapper: true });
    expect(JSON.parse(out)).toEqual({ value: items });
  });

  it('streamed ndjson emits lines as pages arrive', async () => {
    const seen: string[] = [];
    async function* pages() { yield [items[0]]; seen.push('page2'); yield [items[1]]; }
    const lines: string[] = [];
    await streamPages(pages(), 'ndjson', (c) => { lines.push(`${seen.length}:${c}`); });
    expect(lines[0].startsWith('0:')).toBe(true); // first item written before page 2 was requested
    expect(lines).toHaveLength(2);
  });
});

describe('--cursor (§6 pagination)', () => {
  const nextLink = 'https://graph.microsoft.com/v1.0/me/todo/lists/L/tasks?$skiptoken=abc';

  it('a single page prints nextCursor on stderr, and --cursor fetches exactly that link', async () => {
    const r1 = await runCommand(registerTodoCommands, ['todo', 'tasks', 'L', '-o', 'json'], {
      routes: [['GET', '/me/todo/lists/L/tasks', () => ({ value: [{ id: 't1' }], '@odata.nextLink': nextLink })]],
    });
    expect(r1.exitCode).toBe(0);
    const cursor = JSON.parse(r1.stdout).nextCursor as string;
    expect(decodeCursor(cursor)).toBe(nextLink);
    expect(JSON.parse(r1.stdout).nextCursor).toBe(cursor);

    setCursor(cursor);
    const r2 = await runCommand(registerTodoCommands, ['todo', 'tasks', 'L', '-o', 'json'], {
      routes: [['GET', /skiptoken=abc/, () => ({ value: [{ id: 't2' }] })]],
    });
    expect(r2.calls.map(c => c.url)).toEqual([nextLink]);
    expect(JSON.parse(r2.stdout).value[0].id).toBe('t2');
    expect(r2.stderr).not.toContain('nextCursor');
  });

  it('rejects cursors that do not point at Graph (exit 2)', () => {
    try { decodeCursor(encodeCursor('https://evil.example.com/x')); expect.fail(); } catch (e) {
      expect((e as CliError).exitCode).toBe(ExitCode.USAGE);
    }
  });

  it('--all streams every page as ndjson', async () => {
    let n = 0;
    const r = await runCommand(registerTodoCommands, ['todo', 'tasks', 'L', '--all', '-o', 'ndjson'], {
      routes: [['GET', /tasks/, () => (++n === 1 ? { value: [{ id: 'a' }], '@odata.nextLink': nextLink } : { value: [{ id: 'b' }] })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.stdout.trim().split('\n').map(l => JSON.parse(l).id)).toEqual(['a', 'b']);
  });
});

describe('consistent collection envelope (D27)', () => {
  it('table output keeps the cursor on stderr', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'tasks', 'L', '-o', 'table'], {
      routes: [['GET', '/me/todo/lists/L/tasks', () => ({ value: [{ id: 't1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/x?$skiptoken=1' })]],
    });
    expect(r.stdout).not.toContain('nextCursor');
    expect(r.stderr).toMatch(/nextCursor: \S+/);
  });
});
