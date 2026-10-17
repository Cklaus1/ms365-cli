import { vi } from 'vitest';
import { Command } from 'commander';
import { GraphClient } from '../../src/graph-client.js';
import { setExitHandler, resetExitHandler } from '../../src/errors.js';
import { setDryRun } from '../../src/utils.js';

export interface RecordedCall {
  method: string;
  url: string;
  /** Path + query relative to the Graph base, e.g. "/me/messages?$top=10" (decoded). */
  path: string;
  headers: Record<string, string>;
  body?: unknown;
  rawBody?: unknown;
}

/** `signal` is the request's abort signal, so a handler can hang until the client's timeout fires. */
export type RouteHandler = (call: RecordedCall, signal?: AbortSignal) => Response | unknown;
export type Route = [method: string, pattern: RegExp | string, handler: RouteHandler];

export interface RunResult {
  calls: RecordedCall[];
  stdout: string;
  stderr: string;
  exitCode: number;
}

export interface RunOptions {
  routes?: Route[];
  dryRun?: boolean;
  readOnly?: boolean;
  apiVersion?: string;
  /** Simulate an interactive terminal for confirm(); default false (non-TTY). */
  tty?: boolean;
  /** Answer given to confirm() prompts when tty is true. */
  answer?: string;
  env?: Record<string, string>;
  /** The client ID the CLI resolved (GraphClient.clientId()); default none. */
  clientId?: string;
  /** `--refs` (AH-3) on the client. */
  refs?: boolean;
  /** The access token the stub auth returns; default the opaque 'test-token'. */
  token?: string;
}

type Register = (program: Command, getClient: () => GraphClient) => void;

const json = (data: unknown, status = 200) =>
  new Response(JSON.stringify(data), { status, headers: { 'content-type': 'application/json' } });

/**
 * Run one CLI invocation of a command module in-process against a real GraphClient
 * whose fetch is routed to `routes` (first match wins; default: 200 {}).
 */
export async function runCommand(register: Register, argv: string[], opts: RunOptions = {}): Promise<RunResult> {
  const calls: RecordedCall[] = [];
  let stdout = '';
  let stderr = '';
  let exitCode = -1;

  const routes = opts.routes ?? [];
  const origFetch = globalThis.fetch;
  globalThis.fetch = (async (input: string | URL, init: RequestInit = {}) => {
    const url = String(input);
    const method = (init.method ?? 'GET').toUpperCase();
    const headers = { ...(init.headers as Record<string, string> | undefined) };
    let body: unknown = init.body;
    if (typeof body === 'string') { try { body = JSON.parse(body); } catch { /* keep string */ } }
    const base = /^https:\/\/graph\.microsoft\.com\/(v1\.0|beta)/;
    const call: RecordedCall = { method, url, path: decodeURIComponent(url.replace(base, '')), headers, body, rawBody: init.body };
    calls.push(call);
    for (const [m, pattern, handler] of routes) {
      const hit = (m === '*' || m === method) && (typeof pattern === 'string' ? call.path.startsWith(pattern) : pattern.test(call.path));
      if (hit) {
        const r = await handler(call, init.signal ?? undefined);
        return r instanceof Response ? r : json(r ?? {});
      }
    }
    return json({});
  }) as typeof fetch;

  const envBackup: Record<string, string | undefined> = {};
  for (const [k, v] of Object.entries(opts.env ?? {})) { envBackup[k] = process.env[k]; process.env[k] = v; }

  const ttyBackup = process.stdin.isTTY;
  Object.defineProperty(process.stdin, 'isTTY', { value: !!opts.tty, configurable: true });

  const spies = [
    vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => { stdout += a.map(String).join(' ') + '\n'; }),
    vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => { stderr += a.map(String).join(' ') + '\n'; }),
    vi.spyOn(console, 'warn').mockImplementation((...a: unknown[]) => { stderr += a.map(String).join(' ') + '\n'; }),
    vi.spyOn(process.stdout, 'write').mockImplementation(((c: string | Uint8Array, ...rest: unknown[]) => {
      stdout += typeof c === 'string' ? c : Buffer.from(c).toString('latin1');
      const cb = rest.find(r => typeof r === 'function') as (() => void) | undefined;
      cb?.();
      return true;
    }) as never),
    vi.spyOn(process.stderr, 'write').mockImplementation(((c: string | Uint8Array, ...rest: unknown[]) => {
      stderr += String(c);
      const cb = rest.find(r => typeof r === 'function') as (() => void) | undefined;
      cb?.();
      return true;
    }) as never),
  ];

  let readlineMock: { mockRestore(): void } | undefined;
  if (opts.tty) {
    const readline = await import('readline');
    readlineMock = vi.spyOn(readline.default, 'createInterface').mockReturnValue({
      question: (_q: string, cb: (a: string) => void) => cb(opts.answer ?? 'n'),
      close: () => {},
    } as never);
  }

  setExitHandler((c) => { exitCode = c; });
  setDryRun(!!opts.dryRun);
  const auth = { getToken: async () => opts.token ?? 'test-token' };
  const client = new GraphClient(() => auth as never, !!opts.dryRun, opts.apiVersion ?? 'v1.0', !!opts.readOnly, { clientId: opts.clientId, refs: opts.refs });
  const program = new Command().name('ms365').exitOverride();
  program.configureOutput({ writeOut: (s) => { stdout += s; }, writeErr: (s) => { stderr += s; } });
  register(program, () => client);

  try {
    await program.parseAsync(['node', 'ms365', ...argv]);
  } catch (err) {
    const e = err as { code?: string; exitCode?: number };
    if (e.code?.startsWith('commander.')) exitCode = e.exitCode === 0 ? 0 : 2;
    else throw err;
  } finally {
    resetExitHandler();
    setDryRun(false);
    spies.forEach(s => s.mockRestore());
    readlineMock?.mockRestore();
    Object.defineProperty(process.stdin, 'isTTY', { value: ttyBackup, configurable: true });
    globalThis.fetch = origFetch;
    for (const [k, v] of Object.entries(envBackup)) { if (v === undefined) delete process.env[k]; else process.env[k] = v; }
  }
  return { calls, stdout, stderr, exitCode };
}

export { json as jsonResponse };
