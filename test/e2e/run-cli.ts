import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { E2E_OUT_DIR } from './global-setup.js';

export const CLI_ENTRY = path.join(E2E_OUT_DIR, 'index.js');
const FAKE_FETCH = path.resolve(__dirname, 'fake-fetch.mjs');

export interface FakeRoute {
  match: string;
  method?: string;
  status?: number;
  headers?: Record<string, string>;
  body?: unknown;
  bodyBase64?: string;
  delayMs?: number;
}

export interface RunOptions {
  /** Extra env on top of the scrubbed base env (PATH, HOME, TMPDIR only). */
  env?: Record<string, string>;
  /** Canned Graph responses; when set, global fetch is faked via --import. */
  routes?: FakeRoute[];
  /** Working directory; defaults to a fresh temp dir (never the repo, whose .env would leak in). */
  cwd?: string;
  home?: string;
  stdin?: string;
  timeoutMs?: number;
}

export interface RunResult {
  code: number | null;
  signal: NodeJS.Signals | null;
  stdout: string;
  stderr: string;
  /** Requests seen by the fake fetch (only when routes were given). */
  requests: Array<{ method: string; url: string; headers: Record<string, string>; body: string | null }>;
  home: string;
  cwd: string;
}

export function tempDir(prefix = 'ms365-e2e-'): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix));
}

export async function runCli(args: string[], opts: RunOptions = {}): Promise<RunResult> {
  const cwd = opts.cwd ?? tempDir();
  const home = opts.home ?? tempDir('ms365-home-');
  const env: Record<string, string> = {
    PATH: process.env.PATH ?? '',
    HOME: home,
    TMPDIR: os.tmpdir(),
    ...opts.env,
  };
  const nodeArgs: string[] = [];
  let logPath: string | null = null;
  if (opts.routes) {
    const dir = tempDir('ms365-fake-');
    const routesPath = path.join(dir, 'routes.json');
    logPath = path.join(dir, 'requests.jsonl');
    fs.writeFileSync(routesPath, JSON.stringify(opts.routes));
    env.MS365_FAKE_GRAPH = routesPath;
    env.MS365_FAKE_GRAPH_LOG = logPath;
    nodeArgs.push('--import', FAKE_FETCH);
  }
  const child = spawn(process.execPath, [...nodeArgs, CLI_ENTRY, ...args], { cwd, env, stdio: 'pipe' });
  let stdout = '';
  let stderr = '';
  child.stdout.setEncoding('utf8').on('data', (d) => { stdout += d; });
  child.stderr.setEncoding('utf8').on('data', (d) => { stderr += d; });
  child.stdin.end(opts.stdin ?? '');
  const timeout = setTimeout(() => child.kill('SIGKILL'), opts.timeoutMs ?? 20_000);
  const [code, signal] = await new Promise<[number | null, NodeJS.Signals | null]>((resolve) => {
    child.on('close', (c, s) => resolve([c, s]));
  });
  clearTimeout(timeout);
  const requests = logPath && fs.existsSync(logPath)
    ? fs.readFileSync(logPath, 'utf8').trim().split('\n').filter(Boolean).map((l) => JSON.parse(l))
    : [];
  return { code, signal, stdout, stderr, requests, home, cwd };
}
