/**
 * `ms365 mcp serve` protocol core (spec §6 "Call: one protocol", D16): newline-delimited JSON-RPC 2.0
 * over stdio, MCP 2025-06-18. Each tool is a leaf command; each call spawns the CLI, so a command's
 * process.exit and global state never touch the server.
 */
import { spawn } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import readline from 'readline';
import type { Readable } from 'stream';
import type { Command, Option } from 'commander';
import { commandMeta } from './catalog.js';
import { ceilingSurface } from './ceiling.js';
import { cursorPath } from './commands/watch.js';
import { expiresSoon, runTokenHelper } from './external-token.js';
import { CliError, ExitCode } from './errors.js';

export const PROTOCOL_VERSION = '2025-06-18';
const SUPPORTED_VERSIONS = [PROTOCOL_VERSION, '2025-03-26', '2024-11-05'];
const CALL_TIMEOUT_MS = 120_000;
/** Sign-in and profile removal need a person at a terminal; mcp would serve itself. */
// `auth storage use` migrates credentials (a terminal action that may prompt on macOS, ID-16); `show` stays.
// audit-trail (AO-14): admin oversight of access is not delegated to an MCP client; `import` reads a local file path.
const EXCLUDED = new Set(['auth login', 'auth setup', 'auth logout', 'auth profile delete', 'auth grant-scopes', 'auth storage use',
  'audit-trail mailbox', 'audit-trail changes', 'audit-trail apps', 'audit-trail access', 'audit-trail import']);
/**
 * Options left out of a tool's schema (by attribute name). A call is one bounded pass that must finish
 * inside CALL_TIMEOUT_MS with stdin closed: `watch` loses its --follow-only flags (X-6, R-10; the
 * resources below are the long-running path), `drive share` its stdin password (X-15), and
 * `drive copy` its --wait poll, whose 300 s default outlives the call (poll the returned monitorUrl instead).
 */
const HIDDEN_OPTIONS: Record<string, string[]> = {
  watch: ['follow', 'interval', 'maxFailures'],
  drive_share: ['passwordStdin'],
  drive_copy: ['wait', 'timeout'],
};
/** `watch` resources served as MCP resources (X-8). */
const WATCH_RESOURCES = ['mail', 'events', 'files', 'contacts'] as const;
const WATCH_URI = 'ms365://watch/';
const WATCH_MIME = 'application/x-ndjson';
export const DEFAULT_WATCH_INTERVAL_MS = 60_000;
const MAX_WATCH_BACKOFF_MS = 15 * 60_000;
/** Changes kept per resource between reads; the oldest are dropped first and counted in one leading `gap` line (AH-11(1)). */
const WATCH_BUFFER_LIMIT = 1000;

type JsonSchema = Record<string, unknown>;

/** How one schema property maps back to argv. */
type Param =
  | { kind: 'positional'; name: string; variadic: boolean }
  | { kind: 'option'; name: string; flag: string; negFlag?: string; type: 'boolean' | 'string' | 'array' };

export interface Tool {
  name: string;
  path: string;
  description: string;
  inputSchema: { type: 'object'; properties: Record<string, JsonSchema>; required?: string[]; additionalProperties: false };
  annotations: { readOnlyHint: boolean; destructiveHint: boolean; idempotentHint: boolean };
  params: Param[];
  hasOutput: boolean;
}

export interface RunResult { exitCode: number; stdout: string; stderr: string }
export type Runner = (argv: string[]) => Promise<RunResult>;

class RpcError extends Error {
  constructor(public code: number, message: string) { super(message); }
}

function leafCommands(program: Command): Array<{ cmd: Command; path: string }> {
  const out: Array<{ cmd: Command; path: string }> = [];
  const walk = (c: Command, chain: string[]) => {
    if (c.commands.length === 0) out.push({ cmd: c, path: chain.join(' ') });
    for (const sub of c.commands) walk(sub, [...chain, sub.name()]);
  };
  for (const top of program.commands) walk(top, [top.name()]);
  return out;
}

function optionProperty(o: Option, type: 'boolean' | 'string' | 'array'): JsonSchema {
  const prop: JsonSchema = { description: o.description };
  if (type === 'array') {
    prop.type = 'array';
    prop.items = { type: 'string' };
  } else {
    prop.type = type;
    if (o.argChoices) prop.enum = [...o.argChoices];
  }
  if (o.long === '--output') prop.default = 'json';
  else if (type !== 'array' && o.defaultValue !== undefined && typeof o.defaultValue === type) prop.default = o.defaultValue;
  return prop;
}

function toolFor(cmd: Command, path: string): Tool {
  const meta = commandMeta(path);
  const name = path.replace(/ /g, '_');
  const hidden = HIDDEN_OPTIONS[name] ?? [];
  const properties: Record<string, JsonSchema> = {};
  const required: string[] = [];
  const params: Param[] = [];

  for (const a of cmd.registeredArguments) {
    const name = a.name();
    properties[name] = a.variadic ? { type: 'array', items: { type: 'string' } } : { type: 'string' };
    if (a.description) properties[name].description = a.description;
    if (a.argChoices) properties[name].enum = [...a.argChoices];
    if (a.required) required.push(name);
    params.push({ kind: 'positional', name, variadic: a.variadic });
  }

  const byName = new Map<string, Extract<Param, { kind: 'option' }>>();
  for (const o of cmd.options) {
    if (o.hidden || !o.long || hidden.includes(o.attributeName())) continue;
    const attr = o.attributeName();
    const existing = byName.get(attr);
    if (existing) {
      // --x / --no-x pair: one boolean property.
      if (o.negate) existing.negFlag = o.long;
      else { existing.negFlag = existing.flag; existing.flag = o.long; }
      continue;
    }
    const type = o.isBoolean() || o.negate ? 'boolean' : (o.variadic || Array.isArray(o.defaultValue)) ? 'array' : 'string';
    const param: Extract<Param, { kind: 'option' }> = o.negate
      ? { kind: 'option', name: attr, flag: '', negFlag: o.long, type }
      : { kind: 'option', name: attr, flag: o.long, type };
    byName.set(attr, param);
    params.push(param);
    properties[attr] = optionProperty(o, type);
    if (o.mandatory) required.push(attr);
  }

  const kind = meta?.readOnly ? 'read' : meta?.destructive ? 'destructive' : 'write';
  const hints = [
    kind === 'read' ? 'Read-only.' : kind === 'destructive'
      ? 'Destructive: it may delete data, and a delete needs interactive confirmation, so over MCP it is refused; run it in a terminal.'
      : 'Changes data.',
    meta && !meta.idempotent ? 'Not idempotent: repeating it repeats the action.' : '',
    meta?.scopes.length ? `Graph scopes: ${meta.scopes.join(', ')}.` : '',
  ].filter(Boolean).join(' ');

  return {
    name,
    path,
    description: `${cmd.description()}\n\n${hints}`,
    inputSchema: { type: 'object', properties, ...(required.length ? { required } : {}), additionalProperties: false },
    annotations: { readOnlyHint: kind === 'read', destructiveHint: kind === 'destructive', idempotentHint: meta?.idempotent ?? false },
    params,
    hasOutput: cmd.options.some(o => o.long === '--output'),
  };
}

/** `mail_*,calendar_list`-style patterns: `*` matches any run of characters. */
function matcher(patterns: string[]): (name: string) => boolean {
  const res = patterns.map(p => new RegExp(`^${p.replace(/[.+?^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*')}$`));
  return (name) => res.some(r => r.test(name));
}

/**
 * `tool` under a profile ceiling (AH-1(4)): null when no run of it fits; a selector command keeps only the values
 * that fit in its argument or option enum (`watch` → `mail`, `events` under a mail+calendar ceiling).
 */
function underCeiling(tool: Tool, cmd: Command, ceiling: Set<string>): Tool | null {
  const { allowed, values } = ceilingSurface(tool.path, ceiling);
  if (!allowed) return null;
  const sel = commandMeta(tool.path)?.scopesBy;
  if (!sel || !values || 'base' in sel) return tool;
  const key = 'arg' in sel ? cmd.registeredArguments[sel.arg]?.name() : cmd.options.find(o => o.long === sel.option)?.attributeName();
  const prop = key ? tool.inputSchema.properties[key] : undefined;
  if (!prop) return tool;
  const enumValues = (prop.enum as string[] | undefined)?.filter(v => values.includes(v)) ?? values;
  return { ...tool, inputSchema: { ...tool.inputSchema, properties: { ...tool.inputSchema.properties, [key!]: { ...prop, enum: enumValues } } } };
}

/** The MCP tool list for a commander tree, walked live like `describe`; `ceiling` is the active profile's (AH-1). */
export function buildTools(program: Command, opts: { allow?: string[]; readOnly?: boolean; ceiling?: Set<string> | null } = {}): Tool[] {
  const allowed = opts.allow?.length ? matcher(opts.allow) : () => true;
  return leafCommands(program)
    .filter(({ path }) => path.split(' ')[0] !== 'mcp' && !EXCLUDED.has(path))
    .map(({ cmd, path }) => (opts.ceiling ? underCeiling(toolFor(cmd, path), cmd, opts.ceiling) : toolFor(cmd, path)))
    .filter((t): t is Tool => t !== null && allowed(t.name) && (!opts.readOnly || t.annotations.readOnlyHint));
}

function checkType(name: string, value: unknown, schema: JsonSchema): void {
  const ok = schema.type === 'array'
    ? Array.isArray(value) && value.every(v => typeof v === 'string')
    : typeof value === schema.type;
  if (!ok) throw new RpcError(-32602, `Argument "${name}" must be ${schema.type === 'array' ? 'an array of strings' : `a ${schema.type}`}`);
  if (schema.enum && !(schema.enum as unknown[]).includes(value)) {
    throw new RpcError(-32602, `Argument "${name}" must be one of: ${(schema.enum as string[]).join(', ')}`);
  }
}

/** The CLI argv for one tool call (global flags, command path, positionals, options). */
export function buildArgv(tool: Tool, args: Record<string, unknown>, globals: string[] = []): string[] {
  const { properties, required = [] } = tool.inputSchema;
  for (const key of Object.keys(args)) {
    if (!(key in properties)) throw new RpcError(-32602, `Unknown argument "${key}" for ${tool.name}`);
    if (args[key] !== undefined && args[key] !== null) checkType(key, args[key], properties[key]);
  }
  for (const key of required) {
    if (args[key] === undefined || args[key] === null) throw new RpcError(-32602, `Missing required argument "${key}" for ${tool.name}`);
  }

  // Values ride as --flag=value and positionals after `--`, so no value can be read as a flag.
  const argv = [...globals, ...tool.path.split(' ')];
  for (const p of tool.params) {
    if (p.kind !== 'option') continue;
    const v = args[p.name];
    if (v === undefined || v === null) continue;
    if (p.type === 'boolean') {
      const flag = v ? p.flag : p.negFlag;
      if (flag) argv.push(flag);
    } else if (p.type === 'array') {
      for (const item of v as string[]) argv.push(`${p.flag}=${item}`);
    } else if (p.flag !== '--output') {
      argv.push(`${p.flag}=${v as string}`);
    }
  }
  if (tool.hasOutput) argv.push('-o', (args.output as string | undefined) ?? 'json');
  const positionals = tool.params.flatMap(p => {
    const v = p.kind === 'positional' ? args[p.name] : undefined;
    return v === undefined || v === null ? [] : Array.isArray(v) ? v as string[] : [v as string];
  });
  if (positionals.length) argv.push('--', ...positionals);
  return argv;
}

/**
 * Spawn the CLI (`node [loader flags] <entry>`) with stdin closed, killed after 120 s. The call's argv travels in a
 * 0600 file named by MS365_CLI_MCP_ARGV, never on the child's command line, which other local users can read
 * (AH-11(3): mail bodies, long values). MS365_CLI_MCP=1 tells the child it serves MCP (watch's `mcp-` cursor names,
 * no stdin reads; X-6, X-15). `env` adds variables per call (the token from a parent-held helper, AH-2).
 */
export function spawnRunner(entry: string = process.argv[1], timeoutMs = CALL_TIMEOUT_MS, env?: () => Promise<NodeJS.ProcessEnv>): Runner {
  return async (argv) => {
    let extra: NodeJS.ProcessEnv;
    try {
      extra = env ? await env() : {};
    } catch (err) {
      const e = err instanceof Error ? err : new Error(String(err));
      const fields = { code: Reflect.get(e, 'code') ?? 'error', exitCode: Reflect.get(e, 'exitCode') ?? 1, message: e.message, fix: Reflect.get(e, 'fix') };
      return { exitCode: typeof fields.exitCode === 'number' ? fields.exitCode : 1, stdout: '', stderr: `${JSON.stringify({ error: fields })}\n` };
    }
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-mcp-'));
    const argvFile = path.join(dir, 'argv.json');
    fs.writeFileSync(argvFile, JSON.stringify(argv), { mode: 0o600 });
    return new Promise((resolve) => {
      const child = spawn(process.execPath, [...process.execArgv, entry], {
        stdio: ['ignore', 'pipe', 'pipe'],
        env: { ...process.env, ...extra, MS365_CLI_MCP: '1', MS365_CLI_MCP_ARGV: argvFile },
      });
      let stdout = '';
      let stderr = '';
      child.stdout.setEncoding('utf8').on('data', (d: string) => { stdout += d; });
      child.stderr.setEncoding('utf8').on('data', (d: string) => { stderr += d; });
      const timer = setTimeout(() => {
        stderr += `\nTimed out after ${timeoutMs / 1000}s.`;
        child.kill('SIGKILL');
      }, timeoutMs);
      const done = (result: RunResult) => {
        clearTimeout(timer);
        fs.rmSync(dir, { recursive: true, force: true });
        resolve(result);
      };
      child.on('error', (err) => done({ exitCode: 1, stdout, stderr: `${stderr}${err.message}` }));
      child.on('close', (code) => done({ exitCode: code ?? 1, stdout, stderr }));
    });
  };
}

/** The child's side of spawnRunner: read (and delete) the call's argv file; anything but a JSON string array exits 2. */
export function readCallArgv(file: string): string[] {
  let parsed: unknown;
  try {
    parsed = JSON.parse(fs.readFileSync(file, 'utf8'));
    fs.rmSync(file, { force: true });
  } catch (err) {
    throw new CliError(`Cannot read the MCP call arguments ${file}: ${err instanceof Error ? err.message : String(err)}`, { code: 'usage', exitCode: ExitCode.USAGE });
  }
  if (!Array.isArray(parsed) || !parsed.every(a => typeof a === 'string')) {
    throw new CliError(`The MCP call arguments ${file} are not a list of strings.`, { code: 'usage', exitCode: ExitCode.USAGE });
  }
  return parsed;
}

/**
 * `mcp serve` with MS365_CLI_TOKEN_COMMAND: the server runs the helper (cached until `exp − 5 min`; an opaque token
 * until a call fails) and hands the token to each call in MS365_CLI_TOKEN, so N calls do not run it N times (AH-2).
 * A child that gets a 401 still re-runs the helper itself once.
 */
export function helperTokenEnv(command: string): () => Promise<NodeJS.ProcessEnv> {
  let token: Promise<string> | null = null;
  return async () => {
    if (!token || expiresSoon(await token)) {
      token = runTokenHelper(command);
      token.catch(() => { token = null; });
    }
    return { MS365_CLI_TOKEN: await token };
  };
}

function errorText(stderr: string): string {
  const lines = stderr.split('\n').map(l => l.trim()).filter(Boolean);
  const jsonLine = [...lines].reverse().find(l => l.startsWith('{"error"'));
  return jsonLine ?? lines.slice(-20).join('\n');
}

/** The `[dry-run]` request trace from stderr, as a second text item: under dry-run stdout is only a placeholder. */
function dryRunContent(stderr: string): Array<{ type: 'text'; text: string }> {
  const trace = stderr.split('\n').filter(l => l.startsWith('[dry-run]'));
  return trace.length ? [{ type: 'text', text: trace.join('\n') }] : [];
}

function parseJson(text: string): unknown {
  try { return JSON.parse(text); } catch { return undefined; }
}

export interface ServeOptions {
  tools: Tool[];
  version: string;
  runner: Runner;
  readOnly?: boolean;
  /** Global flags forwarded to every call (e.g. --profile from `ms365 --profile x mcp serve`). */
  globals?: string[];
  /** Seconds between watch-resource passes, in ms (default 60 s). */
  watchIntervalMs?: number;
  /** Server diagnostics (default: stderr). */
  log?: (line: string) => void;
  /** Whether the `mcp-<resource>` watch cursor exists (default: its file in the config dir). */
  hasCursor?: (name: string) => boolean;
}

/** Per-session server state handed to every request. */
interface Session {
  watch?: WatchPoller;
}

function callGlobals(opts: ServeOptions): string[] {
  return ['--error-format', 'json', ...(opts.readOnly ? ['--read-only'] : []), ...(opts.globals ?? [])];
}

type WatchResource = typeof WATCH_RESOURCES[number];
interface Subscription { timer?: ReturnType<typeof setTimeout>; delayMs: number; stopped: boolean }

/**
 * `watch` as subscribable resources (X-8): each subscription runs one `watch <r> --name mcp-<r>` pass
 * now and then every interval, buffers the NDJSON it prints and sends `notifications/resources/updated`.
 * A pass resumes from the `mcp-<r>` cursor, so changes made while the server was down are delivered (AH-11(2));
 * only without one does it use --reset, and its first lines are then a `gap` with reason `no_checkpoint`.
 * A failed pass is logged and retried at twice the delay (at most 15 min); it is never a tool error.
 */
class WatchPoller {
  private subs = new Map<string, Subscription>();
  private buffers = new Map<string, string[]>();
  /** Lines dropped from each buffer since its last read (AH-11(1)). */
  private dropped = new Map<string, number>();
  private baselined = new Set<string>();
  private closed = false;
  private readonly intervalMs: number;
  private readonly log: (line: string) => void;
  private readonly hasCursor: (name: string) => boolean;

  /** The watch resources the `watch` tool's enum allows (narrowed by a profile ceiling). */
  private readonly resources: WatchResource[];

  constructor(private opts: ServeOptions, private notify: (msg: object) => void) {
    this.intervalMs = opts.watchIntervalMs ?? DEFAULT_WATCH_INTERVAL_MS;
    this.log = opts.log ?? ((line) => process.stderr.write(`${line}\n`));
    this.hasCursor = opts.hasCursor ?? ((name) => fs.existsSync(cursorPath(name)));
    const allowed = opts.tools.find(t => t.name === 'watch')?.inputSchema.properties.resource?.enum as string[] | undefined;
    this.resources = WATCH_RESOURCES.filter(r => !allowed || allowed.includes(r));
  }

  list() {
    return this.resources.map(r => ({
      uri: `${WATCH_URI}${r}`, name: `watch ${r}`, mimeType: WATCH_MIME,
      description: `Changes to ${r} (NDJSON lines from \`ms365 watch ${r}\`); subscribe to be notified, read to drain.`,
    }));
  }

  subscribe(uri: unknown): void {
    const r = this.resource(uri);
    if (this.subs.has(r)) return;
    const sub: Subscription = { delayMs: this.intervalMs, stopped: false };
    this.subs.set(r, sub);
    void this.pass(r, sub);
  }

  unsubscribe(uri: unknown): void {
    const r = this.resource(uri);
    const sub = this.subs.get(r);
    if (!sub) return;
    sub.stopped = true;
    clearTimeout(sub.timer);
    this.subs.delete(r);
  }

  read(uri: unknown) {
    const r = this.resource(uri);
    const dropped = this.dropped.get(r);
    const lines = [...(dropped ? [JSON.stringify({ type: 'gap', resource: r, reason: 'buffer_overflow', dropped })] : []), ...this.buffers.get(r) ?? []];
    this.buffers.delete(r);
    this.dropped.delete(r);
    return { contents: [{ uri: `${WATCH_URI}${r}`, mimeType: WATCH_MIME, text: lines.map(l => `${l}\n`).join('') }] };
  }

  close(): void {
    this.closed = true;
    for (const r of [...this.subs.keys()]) this.unsubscribe(`${WATCH_URI}${r}`);
  }

  private resource(uri: unknown): WatchResource {
    const r = typeof uri === 'string' && uri.startsWith(WATCH_URI) ? uri.slice(WATCH_URI.length) : '';
    if (!(this.resources as readonly string[]).includes(r)) throw new RpcError(-32602, `Unknown resource: ${String(uri)}`);
    return r as WatchResource;
  }

  private async pass(r: WatchResource, sub: Subscription): Promise<void> {
    const reset = !this.baselined.has(r) && !this.hasCursor(`mcp-${r}`);
    const argv = [...callGlobals(this.opts), 'watch', `--name=mcp-${r}`, ...(reset ? ['--reset'] : []), '--', r];
    const res = await this.opts.runner(argv).catch((err: Error): RunResult => ({ exitCode: 1, stdout: '', stderr: err.message }));
    if (sub.stopped || this.closed) return;
    if (res.exitCode === 0) {
      this.baselined.add(r);
      sub.delayMs = this.intervalMs;
      const lines = res.stdout.split('\n').filter(l => l.trim());
      if (reset) lines.unshift(JSON.stringify({ type: 'gap', resource: r, reason: 'no_checkpoint' }));
      if (lines.length) {
        const buf = [...(this.buffers.get(r) ?? []), ...lines];
        const over = buf.length - WATCH_BUFFER_LIMIT;
        if (over > 0) this.dropped.set(r, (this.dropped.get(r) ?? 0) + over);
        this.buffers.set(r, over > 0 ? buf.slice(over) : buf);
        this.notify({ method: 'notifications/resources/updated', params: { uri: `${WATCH_URI}${r}` } });
      }
    } else {
      sub.delayMs = Math.min(sub.delayMs * 2, MAX_WATCH_BACKOFF_MS);
      const why = errorText(res.stderr) || `exit ${res.exitCode}`;
      this.log(`ms365 mcp: watch ${r} failed (exit ${res.exitCode}): ${why}; retrying in ${sub.delayMs / 1000}s`);
    }
    sub.timer = setTimeout(() => void this.pass(r, sub), sub.delayMs);
  }
}

async function callTool(params: unknown, opts: ServeOptions) {
  if (!params || typeof params !== 'object') throw new RpcError(-32602, 'params must be an object');
  const { name, arguments: args = {} } = params as { name?: unknown; arguments?: unknown };
  const tool = opts.tools.find(t => t.name === name);
  if (!tool) throw new RpcError(-32602, `Unknown tool: ${String(name)}`);
  if (!args || typeof args !== 'object' || Array.isArray(args)) throw new RpcError(-32602, 'arguments must be an object');

  const globals = callGlobals(opts);
  const argv = buildArgv(tool, args as Record<string, unknown>, globals);
  const r = await opts.runner(argv);
  if (r.exitCode !== 0) {
    let text = errorText(r.stderr) || `Exited with code ${r.exitCode}.`;
    if (r.exitCode === 7) {
      text = `${text}\nThis needs interactive confirmation, which MCP cannot give. Run it in a terminal: ms365 ${argv.slice(globals.length).join(' ')}`;
    }
    return { content: [{ type: 'text', text }, ...dryRunContent(r.stderr)], isError: true };
  }
  const structured = parseJson(r.stdout);
  return {
    content: [{ type: 'text', text: r.stdout }, ...dryRunContent(r.stderr)],
    ...(structured && typeof structured === 'object' && !Array.isArray(structured) ? { structuredContent: structured } : {}),
    isError: false,
  };
}

async function handle(msg: Record<string, unknown>, opts: ServeOptions, session: Session): Promise<unknown> {
  const params = msg.params as Record<string, unknown> | undefined;
  const watch = session.watch;
  if (typeof msg.method === 'string' && msg.method.startsWith('resources/') && !watch) {
    throw new RpcError(-32601, `Method not found: ${msg.method}`);
  }
  switch (msg.method) {
    case 'initialize': {
      const asked = params?.protocolVersion;
      return {
        protocolVersion: typeof asked === 'string' && SUPPORTED_VERSIONS.includes(asked) ? asked : PROTOCOL_VERSION,
        capabilities: { tools: { listChanged: false }, ...(watch ? { resources: { subscribe: true, listChanged: false } } : {}) },
        serverInfo: { name: 'ms365', version: opts.version },
      };
    }
    case 'ping':
      return {};
    case 'tools/list':
      return { tools: opts.tools.map(({ name, description, inputSchema, annotations }) => ({ name, description, inputSchema, annotations })) };
    case 'tools/call':
      return callTool(msg.params, opts);
    case 'resources/list':
      return { resources: watch!.list() };
    case 'resources/templates/list':
      return { resourceTemplates: [] };
    case 'resources/read':
      return watch!.read(params?.uri);
    case 'resources/subscribe':
      watch!.subscribe(params?.uri);
      return {};
    case 'resources/unsubscribe':
      watch!.unsubscribe(params?.uri);
      return {};
    default:
      throw new RpcError(-32601, `Method not found: ${String(msg.method)}`);
  }
}

/** Serve MCP on `input` until it ends, writing one JSON-RPC message per `write` call. */
export async function serve(input: Readable, write: (line: string) => void, opts: ServeOptions): Promise<void> {
  const send = (msg: object) => write(JSON.stringify({ jsonrpc: '2.0', ...msg }));
  const pending = new Set<Promise<void>>();
  // Resources ride on the watch tool: none when --allow or --read-only filtered it out.
  const session: Session = {};
  if (opts.tools.some(t => t.name === 'watch')) session.watch = new WatchPoller(opts, send);
  for await (const line of readline.createInterface({ input, crlfDelay: Infinity })) {
    if (!line.trim()) continue;
    let msg: unknown;
    try {
      msg = JSON.parse(line);
    } catch {
      send({ id: null, error: { code: -32700, message: 'Parse error' } });
      continue;
    }
    if (!msg || typeof msg !== 'object' || Array.isArray(msg)) {
      send({ id: null, error: { code: -32600, message: 'Invalid Request' } });
      continue;
    }
    const m = msg as Record<string, unknown>;
    const isRequest = 'id' in m && m.id !== null && typeof m.method === 'string';
    if (!isRequest) continue; // notifications and client responses get no answer
    const id = m.id;
    const p = handle(m, opts, session)
      .then((result) => send({ id, result }))
      .catch((err) => send({ id, error: err instanceof RpcError ? { code: err.code, message: err.message } : { code: -32603, message: String((err as Error)?.message ?? err) } }))
      .finally(() => pending.delete(p));
    pending.add(p);
  }
  session.watch?.close();
  await Promise.all(pending);
}
