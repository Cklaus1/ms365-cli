// Dry-run fidelity (DR-T1–DR-T3): success() silence, the one trailer, confirm() under dry-run, offline commands
// that must not act (DR-22), and the critical-path gate: the real CLI with no client ID.
import { describe, it, expect, vi, afterEach } from 'vitest';
import { spawnSync } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { Command } from 'commander';
import { success, confirm, setDryRun, isDryRun, setQuiet } from '../src/utils.js';
import { runCommand } from './helpers/cli.js';
import { registerTodoCommands } from '../src/commands/todo.js';
import { registerAuthCommands } from '../src/commands/auth.js';
import { saveProfile, profileExists } from '../src/profiles.js';
import { configDir } from '../src/auth.js';

const TRAILER = /^\[dry-run\] (\d+) request\(s\) previewed; nothing was sent\.$/;
const trailers = (s: string) => s.split('\n').filter(l => TRAILER.test(l));

afterEach(() => { setDryRun(false); setQuiet(false); vi.restoreAllMocks(); });

describe('success() and confirm() under dry-run (DR-1, DR-19)', () => {
  it('success() prints through info() normally and stays silent under dry-run', () => {
    const err = vi.spyOn(console, 'error').mockImplementation(() => {});
    success('Email sent.');
    expect(err).toHaveBeenCalledWith('Email sent.');
    err.mockClear();
    setDryRun(true);
    expect(isDryRun()).toBe(true);
    success('Email sent.');
    expect(err).not.toHaveBeenCalled();
  });

  it('confirm() resolves without prompting (even without a TTY) and prints a would-ask line', async () => {
    const orig = process.stdin.isTTY;
    Object.defineProperty(process.stdin, 'isTTY', { value: false, configurable: true });
    const err = vi.spyOn(console, 'error').mockImplementation(() => {});
    try {
      setDryRun(true);
      setQuiet(true); // the would-ask line is part of the trace, not a confirmation
      await expect(confirm('Delete task T?')).resolves.toBe(true);
      expect(err).toHaveBeenCalledWith('[dry-run] would ask: Delete task T?');
    } finally {
      Object.defineProperty(process.stdin, 'isTTY', { value: orig, configurable: true });
    }
  });
});

describe('a dry-run delete through runCommand (AC2)', () => {
  it('previews without a TTY, prints one trailer whose N matches the request lines', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'delete', 'AAMkLIST1', 'T1'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toContain('[dry-run] would ask: ');
    expect(r.stderr).toContain('[dry-run] DELETE https://graph.microsoft.com/v1.0/me/todo/lists/AAMkLIST1/tasks/T1');
    const t = trailers(r.stderr);
    expect(t).toHaveLength(1);
    const requestLines = r.stderr.split('\n').filter(l => /^\[dry-run\] (GET|POST|PUT|PATCH|DELETE) /.test(l));
    expect(Number(TRAILER.exec(t[0])![1])).toBe(requestLines.length);
    expect(r.stderr.trimEnd().split('\n').at(-1)).toMatch(TRAILER);
  });

  it('--quiet drops the trailer but keeps the trace', async () => {
    setQuiet(true);
    const r = await runCommand(registerTodoCommands, ['todo', 'delete', 'AAMkLIST1', 'T1'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain('[dry-run] DELETE');
    expect(trailers(r.stderr)).toHaveLength(0);
  });

  it('no trailer without --dry-run', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'lists', '-o', 'json'], { routes: [['GET', '/me/todo/lists', () => ({ value: [] })]] });
    expect(r.exitCode).toBe(0);
    expect(trailers(r.stderr)).toHaveLength(0);
  });
});

// DR-22 / AC11: with confirm() skipped under dry-run, local-only auth commands must not act.
describe('offline auth commands under --dry-run (DR-22)', () => {
  const noAuth = () => { throw new Error('AuthManager must not be built under dry-run'); };
  const register = (program: Command) => registerAuthCommands(program, noAuth as never);

  it('auth profile delete leaves the profile file unchanged', async () => {
    saveProfile('keepme', { tenantId: 't' });
    const file = path.join(configDir(), 'profiles', 'keepme.json');
    const before = fs.readFileSync(file, 'utf8');
    const r = await runCommand(register, ['auth', 'profile', 'delete', 'keepme'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain('[dry-run] would delete profile keepme');
    expect(fs.readFileSync(file, 'utf8')).toBe(before);
    expect(trailers(r.stderr)).toHaveLength(1);
  });

  it('auth profile delete of a missing profile still exits 5', async () => {
    const r = await runCommand(register, ['auth', 'profile', 'delete', 'nope'], { dryRun: true });
    expect(r.exitCode).toBe(5);
  });

  it('auth profile set writes nothing', async () => {
    const r = await runCommand(register, ['auth', 'profile', 'set', 'fresh', '--tenant-id', 't'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain('[dry-run] would save profile fresh');
    expect(profileExists('fresh')).toBe(false);
  });

  it('auth logout, select and remove leave the token cache and selected account untouched', async () => {
    fs.mkdirSync(configDir(), { recursive: true });
    const cache = path.join(configDir(), 'token-cache.json');
    const sel = path.join(configDir(), 'selected-account.json');
    fs.writeFileSync(cache, '{"cache":1}');
    fs.writeFileSync(sel, '{"accountId":"A"}');
    for (const [argv, line] of [
      [['auth', 'logout'], 'would log out'],
      [['auth', 'select', 'B'], 'would select account B'],
      [['auth', 'remove', 'A'], 'would remove account A'],
      [['auth', 'login'], 'would sign in'],
    ] as const) {
      const r = await runCommand(register, [...argv], { dryRun: true });
      expect(r.exitCode, argv.join(' ')).toBe(0);
      expect(r.stderr).toContain(`[dry-run] ${line}`);
    }
    expect(fs.readFileSync(cache, 'utf8')).toBe('{"cache":1}');
    expect(fs.readFileSync(sel, 'utf8')).toBe('{"accountId":"A"}');
  });
});

// Critical-path gate (DR-T1 + DR-T2): the real CLI, no client ID, empty config dir, a matrix of command families.
describe('dry-run gate: spawned CLI with no client ID', () => {
  const root = path.resolve(__dirname, '..');
  const tsx = path.join(root, 'node_modules', '.bin', 'tsx');
  const home = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-dry-'));
  const cfg = path.join(home, 'cfg');
  fs.mkdirSync(cfg);
  const small = path.join(home, 'small.txt');
  fs.writeFileSync(small, 'hello');
  const PASSWORD = 'Hunter2-s3cret-pw';
  const cli = (args: string[], input = '') => spawnSync(tsx, [path.join(root, 'src/index.ts'), '--dry-run', ...args], {
    cwd: home, encoding: 'utf8', timeout: 30_000, input,
    env: { ...process.env, HOME: home, MS365_CLI_CONFIG_DIR: cfg, MS365_CLI_CLIENT_ID: '', MS365_CLI_TENANT_ID: '',
      MS365_CLI_PROFILE: '', MS365_CLI_TOKEN: '', MS365_CLI_API_VERSION: '' },
  });

  const matrix: Array<{ name: string; args: string[]; json: boolean; input?: string; method: string }> = [
    { name: 'mail send', args: ['mail', 'send', '--to', 'a@b.com', '--subject', 's', '--body', 'b'], json: false, method: 'POST' },
    { name: 'cal create', args: ['calendar', 'create', '--subject', 's', '--start', '2026-10-20T09:00', '--end', '2026-10-20T10:00', '--timezone', 'UTC', '-o', 'json'], json: true, method: 'POST' },
    { name: 'drive upload', args: ['drive', 'upload', 'D1', 'P1', small, '-o', 'json'], json: true, method: 'PUT' },
    { name: 'todo create', args: ['todo', 'create', 'AAMkLIST1', '--title', 't', '-o', 'json'], json: true, method: 'POST' },
    { name: 'drive share --password-stdin', args: ['drive', 'share', 'D1', 'I1', '--scope', 'anonymous', '--password-stdin', '-o', 'json'], json: true, input: `${PASSWORD}\n`, method: 'POST' },
    { name: 'planner delete-task', args: ['planner', 'delete-task', 'T1'], json: false, method: 'DELETE' },
  ];

  for (const c of matrix) {
    it(c.name, () => {
      const r = cli(c.args, c.input);
      expect(r.status, r.stderr).toBe(0);
      expect(r.stderr).toMatch(new RegExp(`^\\[dry-run\\] ${c.method} https://graph\\.microsoft\\.com/`, 'm'));
      expect(r.stdout).not.toContain('[dry-run]');
      if (c.json) expect(JSON.parse(r.stdout)).toMatchObject({ dryRun: true });
      expect(r.stdout + r.stderr).not.toContain(PASSWORD);
      expect(trailers(r.stderr)).toHaveLength(1);
      expect(fs.readdirSync(cfg)).toEqual([]);
    }, 30_000);
  }

  it('the shared link password is shown redacted', () => {
    const r = cli(matrix[4].args, `${PASSWORD}\n`);
    expect(r.stderr).toContain('"password":"<redacted>"');
    expect(JSON.parse(r.stdout).body.password).toBe('<redacted>');
  }, 30_000);
});
