import { describe, it, expect } from 'vitest';
import { execFileSync, spawnSync } from 'child_process';
import { Readable } from 'stream';
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
import { buildTools, serve, type Runner, type RunResult } from '../src/mcp-server.js';
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

async function session(lines: Array<string | object>, opts: { runner?: Runner; allow?: string[]; readOnly?: boolean } = {}) {
  const out: any[] = [];
  const tools = buildTools(program(), { allow: opts.allow, readOnly: opts.readOnly });
  const input = Readable.from(lines.map(l => (typeof l === 'string' ? l : JSON.stringify(l)) + '\n'));
  await serve(input, (line) => out.push(JSON.parse(line)), {
    tools, version: '0.1.0', runner: opts.runner ?? fakeRunner().run, readOnly: opts.readOnly,
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
      result: { protocolVersion: '2025-03-26', capabilities: { tools: { listChanged: false } }, serverInfo: { name: 'ms365', version: '0.1.0' } },
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
