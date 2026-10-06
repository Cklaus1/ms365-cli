import { spawn, type ChildProcess } from 'node:child_process';
import fs from 'node:fs';
import http from 'node:http';
import net from 'node:net';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { expect } from 'vitest';
import { buildScopes } from '../../src/auth.js';
import { aadRoutes } from '../auth-fake-aad.js';
import { CLI_ENTRY, FAKE_FETCH, runCli, tempDir, type FakeRoute, type RunResult } from './run-cli.js';

export const APP_CAP = 'example.com/cap/ms365';
export const DEFAULT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';

export type FakeRequest = RunResult['requests'][number];
export interface ProcEvent { event: 'start' | 'exit'; pid: number; t: number }

export interface Server {
  /** Unix socket (default transport). */
  socket?: string;
  /** Base URL in --port mode. */
  url?: string;
  home: string;
  auditLog: string;
  proc: ChildProcess;
  stderr(): string;
  /** Graph requests made by the server's children. */
  requests(): FakeRequest[];
  /** Start/exit of every process that loaded the fake fetch (the server and its children). */
  procs(): ProcEvent[];
  /** Resolves once a child process (not the server) has started. */
  waitForChild(): Promise<void>;
  /** Signal the server and wait for it to exit. */
  stop(signal?: NodeJS.Signals): Promise<{ code: number | null; signal: NodeJS.Signals | null }>;
}

export interface ServerOptions {
  home?: string;
  cwd?: string;
  routes?: FakeRoute[];
  env?: Record<string, string>;
  /** Extra `serve` options (after the transport and audit options). */
  args?: string[];
  /** Global flags placed before `serve`. */
  globals?: string[];
  /** Listen on loopback TCP (`--port 0`) instead of a Unix socket. */
  tcp?: boolean;
  /** Socket path; default `<tmp>/s.sock`. `null` → no --socket (the default path under HOME). */
  socket?: string | null;
}

const readLines = <T>(file: string): T[] =>
  fs.existsSync(file) ? fs.readFileSync(file, 'utf8').split('\n').filter(Boolean).map((l) => JSON.parse(l) as T) : [];

/** Start `ms365 serve` from the built CLI with the fake Graph preloaded (inherited by every child). */
export async function startServer(opts: ServerOptions = {}): Promise<Server> {
  const dir = tempDir('ms365-srv-');
  const home = opts.home ?? tempDir('ms365-home-');
  const routesPath = path.join(dir, 'routes.json');
  const logPath = path.join(dir, 'requests.jsonl');
  const procsPath = path.join(dir, 'procs.jsonl');
  const auditLog = path.join(dir, 'audit.jsonl');
  fs.writeFileSync(routesPath, JSON.stringify(opts.routes ?? []));
  const fake = pathToFileURL(FAKE_FETCH);
  fake.search = new URLSearchParams({ routes: routesPath, log: logPath, procs: procsPath }).toString();

  const socket = opts.tcp || opts.socket === null ? undefined : opts.socket ?? path.join(dir, 's.sock');
  const transport = opts.tcp ? ['--port', '0'] : socket ? ['--socket', socket] : [];
  const proc = spawn(
    process.execPath,
    ['--import', fake.href, CLI_ENTRY, ...(opts.globals ?? []), 'serve', '--app-cap', APP_CAP, '--audit-log', auditLog, ...transport, ...(opts.args ?? [])],
    { cwd: opts.cwd ?? tempDir(), env: { PATH: process.env.PATH ?? '', HOME: home, ...opts.env }, stdio: ['ignore', 'pipe', 'pipe'] },
  );
  let stderr = '';
  let stdout = '';
  proc.stdout!.setEncoding('utf8').on('data', (d) => { stdout += d; });
  const closed = Promise.withResolvers<{ code: number | null; signal: NodeJS.Signals | null }>();
  proc.on('close', (code, signal) => closed.resolve({ code, signal }));
  const exited = closed.promise;

  const ready = Promise.withResolvers<string>();
  proc.stderr!.setEncoding('utf8').on('data', (d) => {
    stderr += d;
    const m = /listening on (\S+)/.exec(stderr);
    if (m) ready.resolve(m[1]);
  });
  void exited.then(({ code }) => ready.reject(new Error(`serve exited (${code}) before listening:\n${stderr}`)));
  const address = await ready.promise;
  const procs = () => readLines<ProcEvent>(procsPath);

  return {
    socket: address.startsWith('unix:') ? address.slice('unix:'.length) : undefined,
    url: address.startsWith('http') ? address : undefined,
    home,
    auditLog,
    proc,
    stderr: () => stderr,
    requests: () => readLines<FakeRequest>(logPath),
    procs,
    async waitForChild() {
      const started = () => procs().some(e => e.event === 'start' && e.pid !== proc.pid);
      if (started()) return;
      const { promise, resolve } = Promise.withResolvers<void>();
      const watcher = fs.watch(dir, () => {
        if (started()) resolve();
      });
      await promise.finally(() => watcher.close());
    },
    async stop(signal = 'SIGTERM') {
      if (proc.exitCode === null && proc.signalCode === null) proc.kill(signal);
      const result = await exited;
      expect(stdout).toBe(''); // the server never writes to stdout
      return result;
    },
  };
}

/** Value of the `Tailscale-App-Capabilities` header for one grant. */
export function capsHeader(allow: string[], agent = 'chief'): string {
  return JSON.stringify({ [APP_CAP]: [{ agent, allow }] });
}

/**
 * Stand-in for `tailscale serve`: forwards HTTP from a loopback port to the server's Unix socket, replacing any
 * client-sent `Tailscale-App-Capabilities` with the grant `caps` and setting `X-Forwarded-For`.
 */
export async function startServeProxy(server: Server, caps: string): Promise<{ url: string; close(): Promise<void> }> {
  const proxy = http.createServer((req, res) => {
    const headers = { ...req.headers, 'tailscale-app-capabilities': caps, 'x-forwarded-for': '100.64.0.9' };
    const upstream = http.request({ socketPath: server.socket, method: req.method, path: req.url, headers }, (up) => {
      res.writeHead(up.statusCode ?? 502, up.headers);
      up.pipe(res);
    });
    upstream.on('error', () => {
      res.writeHead(502);
      res.end();
    });
    req.pipe(upstream);
  });
  const listening = Promise.withResolvers<void>();
  proxy.listen(0, '127.0.0.1', listening.resolve);
  await listening.promise;
  const { port } = proxy.address() as { port: number };
  return {
    url: `http://127.0.0.1:${port}`,
    async close() {
      const closed = Promise.withResolvers<void>();
      proxy.close(() => closed.resolve());
      proxy.closeAllConnections();
      await closed.promise;
    },
  };
}

export interface CallOptions {
  method?: string;
  path?: string;
  /** `Tailscale-App-Capabilities` value; omitted → no header. */
  caps?: string;
  body?: unknown;
  /** Raw request body (e.g. multipart); set `content-type` in `headers`. */
  raw?: Buffer;
  headers?: Record<string, string>;
}

/** Response body of the serve API (all fields optional: tests assert what they need). */
export interface ApiBody {
  ok?: boolean;
  request_id?: string;
  data?: unknown;
  files?: Array<{ name: string; contentType: string; size: number; base64: string }>;
  stderr?: string;
  error?: { code: string; message: string };
  agent?: string;
  capabilities?: string[];
  commands?: string[];
}

export interface CallResult {
  status: number;
  body: ApiBody;
}

/** One HTTP request to the server (socket or TCP), as `tailscale serve` would forward it. */
export function call(server: Server, opts: CallOptions = {}): Promise<CallResult> {
  const payload = opts.raw ?? (opts.body === undefined ? undefined : JSON.stringify(opts.body));
  const headers: Record<string, string> = { ...opts.headers };
  if (opts.raw === undefined && payload !== undefined) headers['content-type'] = 'application/json';
  if (opts.caps !== undefined) headers['tailscale-app-capabilities'] = opts.caps;
  const target = server.socket ? { socketPath: server.socket } : { host: '127.0.0.1', port: Number(new URL(server.url!).port) };
  const { promise, resolve, reject } = Promise.withResolvers<CallResult>();
  const req = http.request({ ...target, method: opts.method ?? 'POST', path: opts.path ?? '/v1/execute', headers }, (res) => {
    let text = '';
    res.setEncoding('utf8').on('data', (d) => { text += d; });
    res.on('end', () => resolve({ status: res.statusCode ?? 0, body: text ? JSON.parse(text) as ApiBody : {} }));
  });
  req.on('error', reject);
  req.end(payload);
  return promise;
}

/**
 * Send `head` (request line and headers, `\r\n`-separated, no final blank line) verbatim over the server's
 * socket, bypassing Node's client-side checks. Resolves with the status code, or 0 if the connection closed
 * without a response.
 */
export function rawCall(server: Server, head: string): Promise<number> {
  const conn = net.connect(server.socket!);
  const { promise, resolve, reject } = Promise.withResolvers<number>();
  let text = '';
  conn.setEncoding('utf8').on('data', (d) => { text += d; });
  conn.on('error', reject);
  conn.on('close', () => resolve(Number(/^HTTP\/1\.1 (\d{3})/.exec(text)?.[1] ?? 0)));
  conn.end(`${head}\r\nConnection: close\r\n\r\n`);
  return promise;
}

/**
 * Log the fake user in with the built CLI; returns the HOME holding the token cache. The token covers every
 * scope the CLI asks for (org mode included), so children use the cached access token without a refresh.
 */
export async function loginHome(opts: { clientId?: string; expiresIn?: number; env?: Record<string, string> } = {}): Promise<string> {
  const home = tempDir('ms365-home-');
  const clientId = opts.clientId ?? DEFAULT_CLIENT_ID;
  const scope = buildScopes(true).join(' ');
  const r = await runCli(['auth', 'login'], {
    home,
    env: opts.env,
    routes: aadRoutes({ clientId, accessToken: 'at-1', refreshToken: 'rt-1', expiresIn: opts.expiresIn ?? 3600, scope }),
  });
  expect(r.code, r.stderr).toBe(0);
  return home;
}

/** Highest number of processes alive at once, from start/exit events (excluding `excludePid`). */
export function maxAlive(events: ProcEvent[], excludePid?: number): number {
  const sorted = events
    .filter((e) => e.pid !== excludePid)
    .sort((a, b) => a.t - b.t || (a.event === 'exit' ? -1 : 1));
  let alive = 0;
  let max = 0;
  for (const e of sorted) {
    alive += e.event === 'start' ? 1 : -1;
    max = Math.max(max, alive);
  }
  return max;
}
