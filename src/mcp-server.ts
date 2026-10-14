/**
 * `ms365 mcp serve` protocol core (spec §6 "Call: one protocol", D16): newline-delimited JSON-RPC 2.0
 * over stdio, MCP 2025-06-18. Each tool is a leaf command; each call spawns the CLI, so a command's
 * process.exit and global state never touch the server.
 */
import { spawn } from 'child_process';
import readline from 'readline';
import type { Readable } from 'stream';
import type { Command, Option } from 'commander';
import { commandMeta } from './catalog.js';

export const PROTOCOL_VERSION = '2025-06-18';
const SUPPORTED_VERSIONS = [PROTOCOL_VERSION, '2025-03-26', '2024-11-05'];
const CALL_TIMEOUT_MS = 120_000;
/** Sign-in and profile removal need a person at a terminal; mcp would serve itself. */
const EXCLUDED = new Set(['auth login', 'auth setup', 'auth logout', 'auth profile delete']);

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
    if (o.hidden || !o.long) continue;
    const name = o.attributeName();
    const existing = byName.get(name);
    if (existing) {
      // --x / --no-x pair: one boolean property.
      if (o.negate) existing.negFlag = o.long;
      else { existing.negFlag = existing.flag; existing.flag = o.long; }
      continue;
    }
    const type = o.isBoolean() || o.negate ? 'boolean' : (o.variadic || Array.isArray(o.defaultValue)) ? 'array' : 'string';
    const param: Extract<Param, { kind: 'option' }> = o.negate
      ? { kind: 'option', name, flag: '', negFlag: o.long, type }
      : { kind: 'option', name, flag: o.long, type };
    byName.set(name, param);
    params.push(param);
    properties[name] = optionProperty(o, type);
    if (o.mandatory) required.push(name);
  }

  const kind = meta?.readOnly ? 'read' : meta?.destructive ? 'destructive' : 'write';
  const hints = [
    kind === 'read' ? 'Read-only.' : kind === 'destructive'
      ? 'Destructive: deletes data and needs interactive confirmation, so over MCP it is refused; run it in a terminal.'
      : 'Changes data.',
    meta && !meta.idempotent ? 'Not idempotent: repeating it repeats the action.' : '',
    meta?.scopes.length ? `Graph scopes: ${meta.scopes.join(', ')}.` : '',
  ].filter(Boolean).join(' ');

  return {
    name: path.replace(/ /g, '_'),
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

/** The MCP tool list for a commander tree, walked live like `describe`. */
export function buildTools(program: Command, opts: { allow?: string[]; readOnly?: boolean } = {}): Tool[] {
  const allowed = opts.allow?.length ? matcher(opts.allow) : () => true;
  return leafCommands(program)
    .filter(({ path }) => path.split(' ')[0] !== 'mcp' && !EXCLUDED.has(path))
    .map(({ cmd, path }) => toolFor(cmd, path))
    .filter(t => allowed(t.name) && (!opts.readOnly || t.annotations.readOnlyHint));
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

/** Spawn the CLI (`node [loader flags] <entry> ...argv`) with stdin closed, killed after 120 s. */
export function spawnRunner(entry: string = process.argv[1], timeoutMs = CALL_TIMEOUT_MS): Runner {
  return (argv) => new Promise((resolve) => {
    const child = spawn(process.execPath, [...process.execArgv, entry, ...argv], { stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '';
    let stderr = '';
    child.stdout.setEncoding('utf8').on('data', (d: string) => { stdout += d; });
    child.stderr.setEncoding('utf8').on('data', (d: string) => { stderr += d; });
    const timer = setTimeout(() => {
      stderr += `\nTimed out after ${timeoutMs / 1000}s.`;
      child.kill('SIGKILL');
    }, timeoutMs);
    child.on('error', (err) => { clearTimeout(timer); resolve({ exitCode: 1, stdout, stderr: `${stderr}${err.message}` }); });
    child.on('close', (code) => { clearTimeout(timer); resolve({ exitCode: code ?? 1, stdout, stderr }); });
  });
}

function errorText(stderr: string): string {
  const lines = stderr.split('\n').map(l => l.trim()).filter(Boolean);
  const jsonLine = [...lines].reverse().find(l => l.startsWith('{"error"'));
  return jsonLine ?? lines.slice(-20).join('\n');
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
}

async function callTool(params: unknown, opts: ServeOptions) {
  if (!params || typeof params !== 'object') throw new RpcError(-32602, 'params must be an object');
  const { name, arguments: args = {} } = params as { name?: unknown; arguments?: unknown };
  const tool = opts.tools.find(t => t.name === name);
  if (!tool) throw new RpcError(-32602, `Unknown tool: ${String(name)}`);
  if (!args || typeof args !== 'object' || Array.isArray(args)) throw new RpcError(-32602, 'arguments must be an object');

  const globals = ['--error-format', 'json', ...(opts.readOnly ? ['--read-only'] : []), ...(opts.globals ?? [])];
  const argv = buildArgv(tool, args as Record<string, unknown>, globals);
  const r = await opts.runner(argv);
  if (r.exitCode !== 0) {
    let text = errorText(r.stderr) || `Exited with code ${r.exitCode}.`;
    if (r.exitCode === 7) {
      text = `${text}\nThis needs interactive confirmation, which MCP cannot give. Run it in a terminal: ms365 ${argv.slice(globals.length).join(' ')}`;
    }
    return { content: [{ type: 'text', text }], isError: true };
  }
  const structured = parseJson(r.stdout);
  return {
    content: [{ type: 'text', text: r.stdout }],
    ...(structured && typeof structured === 'object' && !Array.isArray(structured) ? { structuredContent: structured } : {}),
    isError: false,
  };
}

async function handle(msg: Record<string, unknown>, opts: ServeOptions): Promise<unknown> {
  const params = msg.params as Record<string, unknown> | undefined;
  switch (msg.method) {
    case 'initialize': {
      const asked = params?.protocolVersion;
      return {
        protocolVersion: typeof asked === 'string' && SUPPORTED_VERSIONS.includes(asked) ? asked : PROTOCOL_VERSION,
        capabilities: { tools: { listChanged: false } },
        serverInfo: { name: 'ms365', version: opts.version },
      };
    }
    case 'ping':
      return {};
    case 'tools/list':
      return { tools: opts.tools.map(({ name, description, inputSchema, annotations }) => ({ name, description, inputSchema, annotations })) };
    case 'tools/call':
      return callTool(msg.params, opts);
    default:
      throw new RpcError(-32601, `Method not found: ${String(msg.method)}`);
  }
}

/** Serve MCP on `input` until it ends, writing one JSON-RPC message per `write` call. */
export async function serve(input: Readable, write: (line: string) => void, opts: ServeOptions): Promise<void> {
  const send = (msg: object) => write(JSON.stringify({ jsonrpc: '2.0', ...msg }));
  const pending = new Set<Promise<void>>();
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
    const p = handle(m, opts)
      .then((result) => send({ id, result }))
      .catch((err) => send({ id, error: err instanceof RpcError ? { code: err.code, message: err.message } : { code: -32603, message: String((err as Error)?.message ?? err) } }))
      .finally(() => pending.delete(p));
    pending.add(p);
  }
  await Promise.all(pending);
}
