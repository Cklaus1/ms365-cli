import { describe, it, expect, vi, afterEach } from 'vitest';
import { execFileSync, spawnSync } from 'child_process';
import { PassThrough, Readable } from 'stream';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { Command } from 'commander';
import { registerAuthCommands } from '../src/commands/auth.js';
import { registerMailCommands } from '../src/commands/mail.js';
import { registerCalendarCommands } from '../src/commands/calendar.js';
import { registerDriveCommands } from '../src/commands/drive.js';
import { registerTodoCommands } from '../src/commands/todo.js';
import { registerTeamsCommands } from '../src/commands/teams.js';
import { registerUserCommands } from '../src/commands/user.js';
import { registerSetupCommand } from '../src/commands/setup.js';
import { registerWatchCommands } from '../src/commands/watch.js';
import { registerDescribeCommand } from '../src/commands/describe.js';
import { registerMcpCommands } from '../src/commands/mcp.js';
import { registerAuditTrailCommands } from '../src/commands/audit-trail.js';
import { buildTools, serve, spawnRunner, type Runner, type RunResult, type ServeOptions } from '../src/mcp-server.js';
import { runCommand } from './helpers/cli.js';
import type { GraphClient } from '../src/graph-client.js';
import type { AuthManager } from '../src/auth.js';

const noClient = (() => { throw new Error('no Graph in mcp tests'); }) as () => GraphClient;
const noAuth = (() => { throw new Error('no auth in mcp tests'); }) as () => AuthManager;

function program(): Command {
  const p = new Command().name('ms365').version('0.1.0').option('--read-only').option('--error-format <f>');
  registerAuthCommands(p, noAuth);
  registerMailCommands(p, noClient);
  registerCalendarCommands(p, noClient);
  registerDriveCommands(p, noClient);
  registerTodoCommands(p, noClient);
  registerTeamsCommands(p, noClient);
  registerUserCommands(p, noClient);
  registerSetupCommand(p, noAuth);
  registerWatchCommands(p, noClient);
  registerAuditTrailCommands(p, noClient);
  registerDescribeCommand(p);
  registerMcpCommands(p);
  return p;
}

/** A runner that records argv and answers with a canned result. */
function fakeRunner(result: Partial<RunResult> = {}) {
  const calls: string[][] = [];
  const run: Runner = async (argv) => {
    calls.push(argv);
    return { exitCode: 0, stdout: '{"ok":true}\n', stderr: '', ...result };
  };
  return { calls, run };
}

async function session(lines: Array<string | object>, opts: { runner?: Runner; allow?: string[]; readOnly?: boolean; globals?: string[] } = {}) {
  const out: any[] = [];
  const tools = buildTools(program(), { allow: opts.allow, readOnly: opts.readOnly });
  const input = Readable.from(lines.map(l => (typeof l === 'string' ? l : JSON.stringify(l)) + '\n'));
  await serve(input, (line) => out.push(JSON.parse(line)), {
    tools, version: '0.1.0', runner: opts.runner ?? fakeRunner().run, readOnly: opts.readOnly, globals: opts.globals,
  });
  return out;
}

const req = (id: number, method: string, params?: unknown) => ({ jsonrpc: '2.0', id, method, ...(params === undefined ? {} : { params }) });
const call = (id: number, name: string, args: Record<string, unknown> = {}) => req(id, 'tools/call', { name, arguments: args });

async function toolList(opts: { allow?: string[]; readOnly?: boolean } = {}) {
  const [res] = await session([req(1, 'tools/list')], opts);
  return res.result.tools as Array<{ name: string; description: string; inputSchema: any; annotations: any }>;
}

describe('mcp protocol', () => {
  it('initialize echoes a supported protocol version and announces tools', async () => {
    const [res] = await session([req(1, 'initialize', { protocolVersion: '2025-03-26', capabilities: {}, clientInfo: { name: 't', version: '1' } })]);
    expect(res).toEqual({
      jsonrpc: '2.0', id: 1,
      result: { protocolVersion: '2025-03-26', capabilities: { tools: { listChanged: false }, resources: { subscribe: true, listChanged: false } }, serverInfo: { name: 'ms365', version: '0.1.0' } },
    });
  });

  it('initialize answers with 2025-06-18 for an unknown version', async () => {
    const [res] = await session([req(1, 'initialize', { protocolVersion: '1999-01-01' })]);
    expect(res.result.protocolVersion).toBe('2025-06-18');
  });

  it('never answers notifications; answers ping', async () => {
    const out = await session([
      { jsonrpc: '2.0', method: 'notifications/initialized' },
      { jsonrpc: '2.0', method: 'no/such/notification' },
      req(2, 'ping'),
    ]);
    expect(out).toEqual([{ jsonrpc: '2.0', id: 2, result: {} }]);
  });

  it('maps malformed JSON, unknown methods and bad params to JSON-RPC errors', async () => {
    const out = await session(['{not json', req(2, 'nope'), req(3, 'tools/call', { name: 'no_such_tool' }), req(4, 'tools/call', 'x')]);
    expect(out[0]).toMatchObject({ id: null, error: { code: -32700 } });
    expect(out.find(r => r.id === 2).error.code).toBe(-32601);
    expect(out.find(r => r.id === 3).error.code).toBe(-32602);
    expect(out.find(r => r.id === 4).error.code).toBe(-32602);
  });
});

describe('mcp tools/list', () => {
  it('exposes one tool per leaf command, minus mcp, sign-in and interactive commands', async () => {
    const tools = await toolList();
    const names = tools.map(t => t.name);
    expect(tools.length).toBeGreaterThanOrEqual(50);
    expect(names).toContain('mail_list');
    expect(names).toContain('auth_profile_list');
    expect(names).toContain('watch');
    for (const gone of ['mcp_serve', 'auth_login', 'auth_setup', 'auth_logout', 'auth_profile_delete']) expect(names).not.toContain(gone);
  });

  it('hides every audit-trail leaf: an admin oversight action, like auth grant-scopes (AO-14, AC 11)', async () => {
    const names = (await toolList()).map(t => t.name);
    expect(names.filter(n => n.startsWith('audit_trail'))).toEqual([]);
    // All five leaf paths are excluded, including those added later (apps, access, import).
    const p = program();
    const at = p.commands.find(c => c.name() === 'audit-trail')!;
    for (const leaf of ['apps', 'access', 'import']) if (!at.commands.some(c => c.name() === leaf)) at.command(leaf).action(() => {});
    expect(buildTools(p).map(t => t.name).filter(n => n.startsWith('audit_trail'))).toEqual([]);
  });

  it('mail_send: required options, boolean flags, repeatable options and catalog annotations', async () => {
    const t = (await toolList()).find(x => x.name === 'mail_send')!;
    expect(t.inputSchema.type).toBe('object');
    expect(t.inputSchema.additionalProperties).toBe(false);
    expect(t.inputSchema.required).toEqual(expect.arrayContaining(['to', 'subject', 'body']));
    expect(t.inputSchema.properties.html).toMatchObject({ type: 'boolean' });
    expect(t.inputSchema.properties.to).toMatchObject({ type: 'string' });
    expect(t.inputSchema.properties.attach).toMatchObject({ type: 'array', items: { type: 'string' } });
    expect(t.annotations).toMatchObject({ readOnlyHint: false, destructiveHint: false, idempotentHint: false });
    expect(t.description).toMatch(/^Send an email/);
  });

  it('positional arguments are required string properties; -o is an enum defaulting to json', async () => {
    const tools = await toolList();
    const get = tools.find(x => x.name === 'mail_get')!;
    expect(get.inputSchema.properties.messageId).toMatchObject({ type: 'string' });
    expect(get.inputSchema.required).toEqual(['messageId']);
    expect(get.inputSchema.properties.output).toMatchObject({ type: 'string', default: 'json', enum: expect.arrayContaining(['json', 'table']) });
    expect(get.annotations).toMatchObject({ readOnlyHint: true, destructiveHint: false, idempotentHint: true });
    const list = tools.find(x => x.name === 'mail_list')!;
    expect(list.inputSchema.properties.pageLimit).toMatchObject({ type: 'string' });
    expect(list.inputSchema.properties.top).toMatchObject({ type: 'string', default: '10' });
    expect(list.inputSchema.properties.all).toMatchObject({ type: 'boolean' });
    const del = tools.find(x => x.name === 'mail_delete')!;
    expect(del.annotations.destructiveHint).toBe(true);
    expect(del.description).toMatch(/destructive/i);
  });

  it('--read-only exposes only read tools', async () => {
    const names = (await toolList({ readOnly: true })).map(t => t.name);
    expect(names).toContain('mail_list');
    for (const w of ['mail_send', 'mail_delete', 'mail_flag', 'calendar_create']) expect(names).not.toContain(w);
  });

  it('--allow restricts tools to matching patterns', async () => {
    const names = (await toolList({ allow: ['mail_*', 'calendar_list'] })).map(t => t.name);
    expect(names).toContain('mail_list');
    expect(names).toContain('mail_send');
    expect(names).toContain('calendar_list');
    expect(names).not.toContain('calendar_create');
    expect(names.every(n => n.startsWith('mail_') || n === 'calendar_list')).toBe(true);
  });
});

describe('mcp tools/call', () => {
  it('builds argv from positionals, options and booleans, with json output and json errors', async () => {
    const r = fakeRunner();
    const [res] = await session([call(1, 'mail_get', { messageId: 'AAMk', unique: true, select: 'id,subject' })], { runner: r.run });
    expect(r.calls).toEqual([['--error-format', 'json', 'mail', 'get', '--select=id,subject', '--unique', '-o', 'json', '--', 'AAMk']]);
    expect(res.result).toEqual({ content: [{ type: 'text', text: '{"ok":true}\n' }], structuredContent: { ok: true }, isError: false });
  });

  it('honours an explicit output, repeats array options and omits false flags', async () => {
    const r = fakeRunner({ stdout: 'a table\n' });
    const [res] = await session([call(1, 'mail_send', { to: 'a@b.c', subject: 's', body: 'b', html: false, attach: ['x.txt', 'y.txt'] })], { runner: r.run });
    expect(r.calls[0]).toEqual(['--error-format', 'json', 'mail', 'send', '--to=a@b.c', '--subject=s', '--body=b', '--attach=x.txt', '--attach=y.txt']);
    expect(res.result.structuredContent).toBeUndefined();
    const r2 = fakeRunner({ stdout: 'a table\n' });
    await session([call(1, 'mail_list', { output: 'table' })], { runner: r2.run });
    expect(r2.calls[0]).toEqual(['--error-format', 'json', 'mail', 'list', '-o', 'table']);
  });

  it('rejects unknown, missing and mistyped arguments with -32602', async () => {
    const r = fakeRunner();
    const out = await session([
      call(1, 'mail_get', { messageId: 'x', bogus: 1 }),
      call(2, 'mail_get', {}),
      call(3, 'mail_get', { messageId: 5 }),
      call(4, 'mail_list', { output: 'xml' }),
      call(5, 'mail_list', {}),
    ], { runner: r.run });
    for (const id of [1, 2, 3, 4]) expect(out.find(o => o.id === id).error.code).toBe(-32602);
    expect(out.find(o => o.id === 5).result.isError).toBe(false);
    expect(r.calls).toHaveLength(1);
  });

  it('maps a non-zero exit to isError with the JSON error line from stderr', async () => {
    const errLine = '{"error":{"code":"not_found","exitCode":5,"message":"gone"}}';
    const r = fakeRunner({ exitCode: 5, stdout: '', stderr: `noise\n${errLine}\n` });
    const [res] = await session([call(1, 'mail_get', { messageId: 'x' })], { runner: r.run });
    expect(res.result.isError).toBe(true);
    expect(res.result.content[0].text).toBe(errLine);
  });

  it('a refused confirmation (exit 7) tells the user to run it in a terminal', async () => {
    const r = fakeRunner({ exitCode: 7, stdout: '', stderr: '{"error":{"code":"confirmation_refused","exitCode":7,"message":"Run in a terminal."}}\n' });
    const [res] = await session([call(1, 'mail_delete', { messageId: 'x' })], { runner: r.run });
    expect(res.result.isError).toBe(true);
    expect(res.result.content[0].text).toMatch(/terminal: ms365 mail delete -- x/);
  });

  it('--read-only passes --read-only to every call and hides write tools', async () => {
    const r = fakeRunner();
    const out = await session([call(1, 'mail_list'), call(2, 'mail_send', { to: 'a', subject: 's', body: 'b' })], { runner: r.run, readOnly: true });
    expect(r.calls).toEqual([['--error-format', 'json', '--read-only', 'mail', 'list', '-o', 'json']]);
    expect(out.find(o => o.id === 2).error.code).toBe(-32602);
  });

  it('a stderr tail stands in when the failure has no JSON error line', async () => {
    const r = fakeRunner({ exitCode: 1, stdout: '', stderr: 'boom\n' });
    const [res] = await session([call(1, 'mail_list')], { runner: r.run });
    expect(res.result).toMatchObject({ isError: true, content: [{ type: 'text', text: 'boom' }] });
  });
});

describe('mcp tool shaping (X-6, X-15)', () => {
  it('watch has no follow-only options; drive_share has no passwordStdin; drive_copy has no wait/timeout', async () => {
    const tools = await toolList();
    const watch = tools.find(t => t.name === 'watch')!.inputSchema.properties;
    for (const gone of ['follow', 'interval', 'maxFailures', 'maxPasses']) expect(watch).not.toHaveProperty(gone);
    expect(watch).toHaveProperty('name');
    expect(watch).toHaveProperty('reset');
    const share = tools.find(t => t.name === 'drive_share')!.inputSchema.properties;
    expect(share).not.toHaveProperty('passwordStdin');
    expect(share).toHaveProperty('type');
    const copy = tools.find(t => t.name === 'drive_copy')!.inputSchema.properties;
    for (const gone of ['wait', 'timeout']) expect(copy).not.toHaveProperty(gone);
    expect(copy).toHaveProperty('name');
  });

  it('a hidden option is an unknown argument', async () => {
    const r = fakeRunner();
    const [res] = await session([call(1, 'watch', { resource: 'mail', follow: true })], { runner: r.run });
    expect(res.error.code).toBe(-32602);
    expect(r.calls).toHaveLength(0);
  });

  it('spawnRunner marks the child with MS365_CLI_MCP=1', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-mcp-env-'));
    const entry = path.join(dir, 'echo-env.mjs');
    fs.writeFileSync(entry, 'process.stdout.write(String(process.env.MS365_CLI_MCP));\n');
    const r = await spawnRunner(entry)([]);
    expect(r).toMatchObject({ exitCode: 0, stdout: '1' });
  });
});

describe('mcp dry-run trace (core-1)', () => {
  const trace = '[dry-run] POST https://graph.microsoft.com/v1.0/me/sendMail\n[dry-run] body: {"message":{}}\n';
  it('adds the stderr [dry-run] lines as a second text item', async () => {
    const r = fakeRunner({ stdout: '{"dryRun":true}\n', stderr: `noise\n${trace}Dry run: 1 request previewed.\n` });
    const [res] = await session([call(1, 'mail_send', { to: 'a@b.c', subject: 's', body: 'b' })], { runner: r.run, globals: ['--dry-run'] });
    expect(r.calls[0]).toContain('--dry-run');
    expect(res.result.isError).toBe(false);
    expect(res.result.content).toEqual([
      { type: 'text', text: '{"dryRun":true}\n' },
      { type: 'text', text: '[dry-run] POST https://graph.microsoft.com/v1.0/me/sendMail\n[dry-run] body: {"message":{}}' },
    ]);
  });

  it('keeps isError on a failed dry-run call and still shows the trace', async () => {
    const errLine = '{"error":{"code":"usage","exitCode":2,"message":"bad"}}';
    const r = fakeRunner({ exitCode: 2, stdout: '', stderr: `${trace}${errLine}\n` });
    const [res] = await session([call(1, 'mail_list')], { runner: r.run, globals: ['--dry-run'] });
    expect(res.result.isError).toBe(true);
    expect(res.result.content[0].text).toBe(errLine);
    expect(res.result.content[1].text).toMatch(/^\[dry-run\] POST/);
  });

  it('a live call has one content item', async () => {
    const [res] = await session([call(1, 'mail_list')], { runner: fakeRunner({ stderr: 'note\n' }).run });
    expect(res.result.content).toHaveLength(1);
  });
});

describe('mcp watch resources (X-8)', () => {
  afterEach(() => { vi.useRealTimers(); });

  /** A server on an open stdin; `send` writes a message, `flush` lets the promise queue drain. */
  function live(runner: Runner, extra: Partial<ServeOptions> = {}) {
    const out: any[] = [];
    const logs: string[] = [];
    const input = new PassThrough();
    const done = serve(input, (line) => out.push(JSON.parse(line)), {
      tools: buildTools(program()), version: '0.1.0', runner, log: (l) => logs.push(l), watchIntervalMs: 60_000, ...extra,
    });
    const flush = async () => { for (let i = 0; i < 20; i++) await vi.advanceTimersByTimeAsync(0); };
    const send = async (msg: object) => { input.write(JSON.stringify(msg) + '\n'); await flush(); };
    const end = async () => { input.end(); await done; };
    return { out, logs, send, flush, end };
  }

  /** A runner whose watch passes answer from a queue (then empty). */
  function queued(results: Array<Partial<RunResult>>) {
    const calls: string[][] = [];
    const run: Runner = async (argv) => { calls.push(argv); return { exitCode: 0, stdout: '', stderr: '', ...(results.shift() ?? {}) }; };
    return { calls, run };
  }

  const updated = (out: any[]) => out.filter(m => m.method === 'notifications/resources/updated');

  it('resources/list names one resource per watch resource', async () => {
    vi.useFakeTimers();
    const s = live(queued([]).run);
    await s.send(req(1, 'resources/list'));
    await s.end();
    expect(s.out[0].result.resources.map((r: any) => r.uri)).toEqual(['ms365://watch/mail', 'ms365://watch/events', 'ms365://watch/files', 'ms365://watch/contacts']);
    expect(s.out[0].result.resources[0]).toMatchObject({ name: expect.any(String), mimeType: 'application/x-ndjson' });
  });

  it('subscribe polls (first pass --reset), notifies on changes, read drains the buffer', async () => {
    vi.useFakeTimers();
    const r = queued([{ stdout: '' }, { stdout: '{"id":"1","change":"created"}\n{"id":"2","change":"updated"}\n' }]);
    const s = live(r.run);
    await s.send(req(1, 'resources/subscribe', { uri: 'ms365://watch/mail' }));
    expect(s.out.find(m => m.id === 1).result).toEqual({});
    expect(r.calls).toEqual([['--error-format', 'json', 'watch', '--name=mcp-mail', '--reset', '--', 'mail']]);
    expect(updated(s.out)).toHaveLength(0);
    await vi.advanceTimersByTimeAsync(60_000); await s.flush();
    expect(r.calls[1]).toEqual(['--error-format', 'json', 'watch', '--name=mcp-mail', '--', 'mail']);
    expect(updated(s.out)).toEqual([{ jsonrpc: '2.0', method: 'notifications/resources/updated', params: { uri: 'ms365://watch/mail' } }]);
    await s.send(req(2, 'resources/read', { uri: 'ms365://watch/mail' }));
    expect(s.out.find(m => m.id === 2).result.contents).toEqual([{ uri: 'ms365://watch/mail', mimeType: 'application/x-ndjson', text: '{"id":"1","change":"created"}\n{"id":"2","change":"updated"}\n' }]);
    await s.send(req(3, 'resources/read', { uri: 'ms365://watch/mail' }));
    expect(s.out.find(m => m.id === 3).result.contents[0].text).toBe('');
    await s.end();
  });

  it('unsubscribe stops the loop; re-subscribing in the same session does not --reset again (F10)', async () => {
    vi.useFakeTimers();
    const r = queued([]);
    const s = live(r.run);
    await s.send(req(1, 'resources/subscribe', { uri: 'ms365://watch/files' }));
    await s.send(req(2, 'resources/unsubscribe', { uri: 'ms365://watch/files' }));
    await vi.advanceTimersByTimeAsync(300_000); await s.flush();
    expect(r.calls).toHaveLength(1);
    await s.send(req(3, 'resources/subscribe', { uri: 'ms365://watch/files' }));
    expect(r.calls[1]).not.toContain('--reset');
    await s.end();
  });

  it('a failed pass logs to stderr, sends nothing and doubles the interval; success restores it', async () => {
    vi.useFakeTimers();
    const r = queued([{ stdout: '' }, { exitCode: 1, stderr: '{"error":{"code":"network","message":"down"}}\n' }, { stdout: '{"id":"9"}\n' }]);
    const s = live(r.run);
    await s.send(req(1, 'resources/subscribe', { uri: 'ms365://watch/events' }));
    await vi.advanceTimersByTimeAsync(60_000); await s.flush();
    expect(r.calls).toHaveLength(2);
    expect(s.logs.join('\n')).toMatch(/watch events.*down.*120s/);
    expect(s.out.filter(m => m.method || m.error)).toHaveLength(0);
    await vi.advanceTimersByTimeAsync(119_000); await s.flush();
    expect(r.calls).toHaveLength(2);
    await vi.advanceTimersByTimeAsync(1_000); await s.flush();
    expect(r.calls).toHaveLength(3);
    expect(updated(s.out)).toHaveLength(1);
    await vi.advanceTimersByTimeAsync(60_000); await s.flush();
    expect(r.calls).toHaveLength(4);
    await s.end();
  });

  it('a reset pass that fails is retried with --reset', async () => {
    vi.useFakeTimers();
    const r = queued([{ exitCode: 3, stderr: 'auth\n' }]);
    const s = live(r.run);
    await s.send(req(1, 'resources/subscribe', { uri: 'ms365://watch/contacts' }));
    await vi.advanceTimersByTimeAsync(120_000); await s.flush();
    expect(r.calls[1]).toContain('--reset');
    await s.end();
  });

  it('end of stdin stops every poller', async () => {
    vi.useFakeTimers();
    const r = queued([]);
    const s = live(r.run);
    await s.send(req(1, 'resources/subscribe', { uri: 'ms365://watch/mail' }));
    await s.end();
    await vi.advanceTimersByTimeAsync(600_000);
    expect(r.calls).toHaveLength(1);
  });

  it('the buffer keeps the newest 1000 changes', async () => {
    vi.useFakeTimers();
    const many = Array.from({ length: 1005 }, (_, i) => `{"id":"${i}"}`).join('\n') + '\n';
    const s = live(queued([{ stdout: many }]).run);
    await s.send(req(1, 'resources/subscribe', { uri: 'ms365://watch/mail' }));
    await s.send(req(2, 'resources/read', { uri: 'ms365://watch/mail' }));
    const lines = s.out.find(m => m.id === 2).result.contents[0].text.trim().split('\n');
    expect(lines).toHaveLength(1000);
    expect(lines[0]).toBe('{"id":"5"}');
    await s.end();
  });

  it('rejects unknown resource URIs with -32602', async () => {
    vi.useFakeTimers();
    const s = live(queued([]).run);
    await s.send(req(1, 'resources/subscribe', { uri: 'ms365://watch/nope' }));
    await s.send(req(2, 'resources/read', { uri: 'file:///etc/passwd' }));
    await s.end();
    expect(s.out.find(m => m.id === 1).error.code).toBe(-32602);
    expect(s.out.find(m => m.id === 2).error.code).toBe(-32602);
  });

  it('without the watch tool (--allow / --read-only filters), no resources are advertised or served', async () => {
    const out = await session([req(1, 'initialize', {}), req(2, 'resources/list')], { allow: ['mail_*'] });
    expect(out.find(m => m.id === 1).result.capabilities).toEqual({ tools: { listChanged: false } });
    expect(out.find(m => m.id === 2).error.code).toBe(-32601);
  });
});

describe('mcp serve --watch-interval', () => {
  it('rejects an interval under 15 s with a usage error before serving', async () => {
    const r = await runCommand((p) => registerMcpCommands(p), ['mcp', 'serve', '--watch-interval', '5']);
    expect(r.exitCode).toBe(2);
    expect(r.stderr).toMatch(/--watch-interval must be at least 15/);
  });
});

const root = new URL('..', import.meta.url).pathname;
const tsx = `${root}node_modules/.bin/tsx`;
const registered = (() => {
  try {
    return /^\s+mcp\b/m.test(execFileSync(tsx, ['src/index.ts', '--help'], { cwd: root, encoding: 'utf8' }));
  } catch { return false; }
})();

describe('ms365 mcp serve (spawned)', () => {
  it.skipIf(!registered)('answers initialize and tools/list over stdio, offline', () => {
    const input = [req(1, 'initialize', { protocolVersion: '2025-06-18' }), { jsonrpc: '2.0', method: 'notifications/initialized' }, req(2, 'tools/list')]
      .map(m => JSON.stringify(m)).join('\n') + '\n';
    const r = spawnSync(tsx, ['src/index.ts', 'mcp', 'serve'], {
      cwd: root, input, encoding: 'utf8', timeout: 30_000,
      env: { ...process.env, MS365_CLI_CONFIG_DIR: fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-mcp-')) },
    });
    expect(r.status).toBe(0);
    const out = r.stdout.trim().split('\n').map(l => JSON.parse(l));
    expect(out.map(o => o.id)).toEqual([1, 2]);
    expect(out[0].result.serverInfo.name).toBe('ms365');
    expect(out[1].result.tools.length).toBeGreaterThanOrEqual(50);
  }, 60_000);
});
