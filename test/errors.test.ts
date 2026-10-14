import { describe, it, expect, afterEach, vi } from 'vitest';
import { spawnSync } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import {
  ExitCode, CliError, ConfirmationRefusedError, toCliError, reportError,
  setErrorFormat, handleErrors, setExitHandler, resetExitHandler,
} from '../src/errors.js';
import { graphError } from '../src/graph-client.js';
import { confirm } from '../src/utils.js';

afterEach(() => {
  resetExitHandler();
  setErrorFormat('text');
  process.exitCode = undefined;
  vi.restoreAllMocks();
});

describe('toCliError', () => {
  it('classifies auth failures as exit 3 with a fix', () => {
    for (const msg of ['Not logged in. Run: ms365 auth login', 'No client ID configured', 'invalid_grant: AADSTS700082 expired']) {
      const e = toCliError(new Error(msg));
      expect(e.exitCode).toBe(ExitCode.AUTH);
      expect(e.fix).toBe('ms365 auth login');
    }
  });
  it('defaults to generic exit 1', () => {
    expect(toCliError(new Error('boom')).exitCode).toBe(ExitCode.GENERIC);
    expect(toCliError('str').message).toBe('str');
  });
  it('passes CliError through', () => {
    const e = new CliError('x', { exitCode: ExitCode.NOT_FOUND });
    expect(toCliError(e)).toBe(e);
  });
});

describe('graphError', () => {
  it.each([
    [401, ExitCode.AUTH, 'auth_failed'],
    [403, ExitCode.PERMISSION, 'permission_denied'],
    [404, ExitCode.NOT_FOUND, 'not_found'],
    [429, ExitCode.THROTTLED, 'throttled'],
    [500, ExitCode.GENERIC, 'graph_error'],
  ])('maps %i to exit %i', (status, exit, code) => {
    const e = graphError(status, 'msg', { requestId: 'rid', graphCode: 'gc' });
    expect(e.exitCode).toBe(exit);
    expect(e.code).toBe(code);
    expect(e.status).toBe(status);
    expect(e.requestId).toBe('rid');
  });
});

describe('reportError', () => {
  it('writes one JSON line in json mode', () => {
    const writes: string[] = [];
    vi.spyOn(process.stderr, 'write').mockImplementation((c: string | Uint8Array) => { writes.push(String(c)); return true; });
    setErrorFormat('json');
    const code = reportError(new CliError('nope', { code: 'not_found', exitCode: ExitCode.NOT_FOUND, status: 404, requestId: 'r1' }));
    expect(code).toBe(5);
    const parsed = JSON.parse(writes.join('').trim());
    expect(parsed.error).toMatchObject({ code: 'not_found', exitCode: 5, status: 404, requestId: 'r1', message: 'nope' });
  });
  it('writes a text line with fix hint', () => {
    const writes: string[] = [];
    vi.spyOn(process.stderr, 'write').mockImplementation((c: string | Uint8Array) => { writes.push(String(c)); return true; });
    reportError(new CliError('bad', { fix: 'do x' }));
    expect(writes.join('')).toBe('Error: bad (fix: do x)\n');
  });
});

describe('handleErrors', () => {
  async function run(fn: () => Promise<void>): Promise<number> {
    let code = -1;
    setExitHandler((c) => { code = c; });
    vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
    await handleErrors(fn)();
    return code;
  }
  it('exits 0 on success', async () => { expect(await run(async () => {})).toBe(0); });
  it('exits with the CliError code', async () => {
    expect(await run(async () => { throw new CliError('x', { exitCode: ExitCode.THROTTLED }); })).toBe(6);
  });
  it('exits 1 on a plain Error', async () => { expect(await run(async () => { throw new Error('x'); })).toBe(1); });
  it('honors an exitCode set by the command', async () => {
    expect(await run(async () => { process.exitCode = 4; })).toBe(4);
  });
  it('exits 7 when confirmation is refused', async () => {
    expect(await run(async () => { throw new ConfirmationRefusedError(); })).toBe(7);
  });
});

describe('confirm', () => {
  it('throws ConfirmationRefusedError on non-TTY stdin', async () => {
    const orig = process.stdin.isTTY;
    Object.defineProperty(process.stdin, 'isTTY', { value: false, configurable: true });
    try {
      await expect(confirm('Delete?')).rejects.toBeInstanceOf(ConfirmationRefusedError);
    } finally {
      Object.defineProperty(process.stdin, 'isTTY', { value: orig, configurable: true });
    }
  });
});

// Critical-path gate: the real CLI process exits with the contracted codes.
describe('exit-code matrix (spawned CLI)', () => {
  const root = path.resolve(__dirname, '..');
  const tsx = path.join(root, 'node_modules', '.bin', 'tsx');
  const home = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-exit-'));
  const env = {
    ...process.env, HOME: home, MS365_CLI_TOKEN: '',
    MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000',
  };
  const cli = (...args: string[]) => spawnSync(tsx, [path.join(root, 'src/index.ts'), ...args],
    { env, cwd: home, encoding: 'utf8', timeout: 30_000, input: '' });

  it.each([
    ['success (dry-run)', ['--dry-run', 'mail', 'send', '--to', 'a@b.com', '--subject', 's', '--body', 'b'], 0],
    ['generic (bad email)', ['--dry-run', 'mail', 'send', '--to', 'not-an-email', '--subject', 's', '--body', 'b'], 1],
    ['usage (unknown option)', ['mail', 'send', '--bogus'], 2],
    ['auth (not logged in)', ['--read-only', 'user', 'me'], 3],
    ['confirmation refused (non-TTY)', ['todo', 'delete', 'L', 'T'], 7],
  ])('%s', (_name, args, want) => {
    const r = cli(...(args as string[]));
    expect(r.status, `stderr: ${r.stderr}`).toBe(want);
  }, 30_000);

  it('json error format on auth failure', () => {
    const r = cli('--error-format', 'json', '--read-only', 'user', 'me');
    expect(r.status).toBe(3);
    const last = r.stderr.trim().split('\n').pop()!;
    expect(JSON.parse(last).error).toMatchObject({ exitCode: 3, code: 'auth_required' });
  }, 30_000);
});
