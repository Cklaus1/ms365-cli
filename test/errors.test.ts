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
import { confirm, setDryRun, setQuiet, notePreview, fromGet } from '../src/utils.js';

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
  it('includes maybeProcessed in the JSON payload only when set (R-7)', () => {
    const writes: string[] = [];
    vi.spyOn(process.stderr, 'write').mockImplementation((c: string | Uint8Array) => { writes.push(String(c)); return true; });
    setErrorFormat('json');
    reportError(new CliError('timed out', { code: 'timeout', maybeProcessed: true }));
    reportError(new CliError('nope', { code: 'not_found' }));
    const [tagged, plain] = writes.join('').trim().split('\n').map(l => JSON.parse(l).error);
    expect(tagged).toMatchObject({ code: 'timeout', maybeProcessed: true });
    expect('maybeProcessed' in plain).toBe(false);
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

  // DR-1 / F2: one trailer on the success path of a dry run; N counts this command's previews only.
  describe('dry-run trailer', () => {
    afterEach(() => { setDryRun(false); setQuiet(false); });
    const stderrOf = async (fn: () => Promise<void>) => {
      const lines: string[] = [];
      setExitHandler(() => {});
      vi.spyOn(console, 'error').mockImplementation((s: string) => { lines.push(s); });
      await handleErrors(fn)();
      return lines;
    };
    it('prints "[dry-run] N request(s) previewed; nothing was sent." once', async () => {
      setDryRun(true);
      notePreview(); // left over from an earlier command: reset at the start of each action
      expect(await stderrOf(async () => { notePreview(); notePreview(); })).toEqual(['[dry-run] 2 request(s) previewed; nothing was sent.']);
    });
    it('notes placeholders when a body used <from GET> in this command (DR-11)', async () => {
      setDryRun(true);
      expect(await stderrOf(async () => { notePreview(); expect(fromGet()).toBe('<from GET>'); notePreview(); }))
        .toEqual(['[dry-run] 2 request(s) previewed; nothing was sent; bodies use placeholders for values read from Graph.']);
      expect(await stderrOf(async () => { notePreview(); })).toEqual(['[dry-run] 1 request(s) previewed; nothing was sent.']);
    });
    it('is absent without --dry-run, under --quiet, and on failure', async () => {
      expect(await stderrOf(async () => { notePreview(); })).toEqual([]);
      setDryRun(true);
      vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
      expect(await stderrOf(async () => { throw new CliError('bad', { exitCode: ExitCode.USAGE }); })).toEqual([]);
      setQuiet(true);
      expect(await stderrOf(async () => { notePreview(); })).toEqual([]);
    });
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
  const cliWithToken = (...args: string[]) => spawnSync(tsx, [path.join(root, 'src/index.ts'), ...args],
    { env: { ...env, MS365_CLI_TOKEN: 'opaque-test-token' }, cwd: home, encoding: 'utf8', timeout: 30_000, input: '' });

  it.each([
    ['success (dry-run)', ['--dry-run', 'mail', 'send', '--to', 'a@b.com', '--subject', 's', '--body', 'b'], 0],
    ['usage (bad email, V-11)', ['--dry-run', 'mail', 'send', '--to', 'not-an-email', '--subject', 's', '--body', 'b'], 2],
    ['usage (unknown option)', ['mail', 'send', '--bogus'], 2],
    ['auth (not logged in)', ['--read-only', 'user', 'me'], 3],
    ['confirmation refused (non-TTY)', ['todo', 'delete', 'L', 'T'], 7],
  ])('%s', (_name, args, want) => {
    // The confirmation case needs a token (an opaque one skips the scope pre-flight)
    const r = want === 7 ? cliWithToken(...(args as string[])) : cli(...(args as string[]));
    expect(r.status, `stderr: ${r.stderr}`).toBe(want);
  }, 30_000);

  it('json error format on auth failure', () => {
    const r = cli('--error-format', 'json', '--read-only', 'user', 'me');
    expect(r.status).toBe(3);
    const last = r.stderr.trim().split('\n').pop()!;
    expect(JSON.parse(last).error).toMatchObject({ exitCode: 3, code: 'auth_required' });
  }, 30_000);
});
