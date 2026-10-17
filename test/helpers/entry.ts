import { spawnSync } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';

/**
 * Run the real CLI entry (src/index.ts through tsx) with the offline Graph double preloaded (E5,
 * test/fixtures/fetch-stub.mjs). Spawned cases are capped at 12 for the agent-host (AH) subset — prefer
 * in-process tests; use this only for what needs preAction, profiles or the AuthManager.
 */
const root = path.resolve(__dirname, '..', '..');
const tsxLoader = path.join(root, 'node_modules', 'tsx', 'dist', 'esm', 'index.mjs');
const stub = path.join(root, 'test', 'fixtures', 'fetch-stub.mjs');

export interface StubRoute { method?: string; match: string; status?: number; body?: unknown; headers?: Record<string, string>; hang?: boolean }

export interface EntryResult {
  status: number | null;
  stdout: string;
  stderr: string;
  /** `METHOD URL` of every fetch the run made. */
  fetches: string[];
  /** The last non-empty stderr line parsed as JSON (the error-line contract), or undefined. */
  error?: { code: string; exitCode: number; message: string; fix?: string; [k: string]: unknown };
}

/** A fresh HOME and config dir for one spawned run. */
export function entryHome(): { home: string; configDir: string } {
  const home = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-entry-'));
  const configDir = path.join(home, 'cfg');
  fs.mkdirSync(configDir);
  return { home, configDir };
}

export function runEntry(args: string[], opts: { env?: Record<string, string>; routes?: StubRoute[]; home?: string; configDir?: string; input?: string; preload?: string[] } = {}): EntryResult {
  const { home, configDir } = opts.home ? { home: opts.home, configDir: opts.configDir ?? path.join(opts.home, 'cfg') } : entryHome();
  const r = spawnSync(process.execPath, ['--import', tsxLoader, '--import', stub, ...(opts.preload ?? []).flatMap(p => ['--import', p]), path.join(root, 'src/index.ts'), ...args], {
    cwd: home, encoding: 'utf8', timeout: 30_000, input: opts.input ?? '',
    env: {
      ...process.env, HOME: home, MS365_CLI_CONFIG_DIR: configDir, MS365_CLI_TOKEN: '', MS365_CLI_PROFILE: '',
      MS365_CLI_CLIENT_ID: '', MS365_CLI_API_VERSION: '', MS365_CLI_ERROR_FORMAT: 'json',
      MS365_TEST_FETCH_ROUTES: JSON.stringify(opts.routes ?? []), ...opts.env,
    },
  });
  const lines = r.stderr.split('\n').filter(l => l.trim());
  const fetches = lines.filter(l => l.startsWith('[fetch] ')).map(l => l.slice(8));
  let error: EntryResult['error'];
  try { error = JSON.parse(lines[lines.length - 1] ?? '').error; } catch { /* not JSON */ }
  return { status: r.status, stdout: r.stdout, stderr: r.stderr, fetches, error };
}

/** An unsigned JWT with these claims (tests only; the CLI never verifies signatures). */
export function fakeJwt(claims: Record<string, unknown>): string {
  return ['{"alg":"none"}', JSON.stringify(claims)].map(p => Buffer.from(p).toString('base64url')).join('.') + '.sig';
}
