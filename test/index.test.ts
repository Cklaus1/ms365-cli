import { describe, it, expect } from 'vitest';
import { spawnSync } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';

// Global-flag parsing and offline commands, exercised through the real CLI entry point.
const root = path.resolve(__dirname, '..');
const tsx = path.join(root, 'node_modules', '.bin', 'tsx');
const home = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-idx-'));
function cli(args: string[], extraEnv: Record<string, string> = {}) {
  return spawnSync(tsx, [path.join(root, 'src/index.ts'), ...args], {
    cwd: home, encoding: 'utf8', timeout: 30_000, input: '',
    env: { ...process.env, HOME: home, MS365_CLI_TOKEN: '',
      MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_API_VERSION: '', ...extraEnv },
  });
}

describe('global flags (C-15)', () => {
  it('accepts --api-version=beta (= form)', () => {
    const r = cli(['--dry-run', '--api-version=beta', 'user', 'me']);
    expect(r.status, r.stderr).toBe(0);
    expect(r.stdout).toContain('https://graph.microsoft.com/beta/me');
  }, 30_000);

  it('rejects an invalid --api-version with exit 2', () => {
    const r = cli(['--dry-run', '--api-version', '../x', 'user', 'me']);
    expect(r.status).toBe(2);
  }, 30_000);

  it('rejects an invalid MS365_CLI_API_VERSION with exit 2', () => {
    const r = cli(['--dry-run', 'user', 'me'], { MS365_CLI_API_VERSION: '../x' });
    expect(r.status).toBe(2);
  }, 30_000);

  it('does not treat a flag-looking option VALUE as a global flag', () => {
    // Old argv.includes() parsing turned this into a dry-run (exit 0). Now it's a real
    // send attempt with no login in this HOME → auth error (exit 3).
    const r = cli(['mail', 'send', '--to', 'a@b.com', '--subject', 's', '--body=--dry-run']);
    expect(r.stdout).not.toContain('[dry-run]');
    expect(r.status).toBe(3);
  }, 30_000);

  it('global flags work after the subcommand too', () => {
    const r = cli(['user', 'me', '--dry-run']);
    expect(r.status, r.stderr).toBe(0);
    expect(r.stdout).toContain('[dry-run] GET');
  }, 30_000);
});

describe('offline commands (C-16, M-25)', () => {
  const noClientId = { MS365_CLI_CLIENT_ID: '' };
  it('schema --list works without a client ID', () => {
    const r = cli(['schema', '--list'], noClientId);
    expect(r.status, r.stderr).toBe(0);
    expect(r.stdout).toContain('Available schemas');
  }, 30_000);

  it('schema with no argument lists schemas', () => {
    const r = cli(['schema'], noClientId);
    expect(r.status, r.stderr).toBe(0);
  }, 30_000);

  it('schema <endpoint> works without a client ID; unknown endpoint exits 2', () => {
    expect(cli(['schema', 'mail.list'], noClientId).status).toBe(0);
    const bad = cli(['schema', 'mail.nope'], noClientId);
    expect(bad.status).toBe(2);
    expect(bad.stderr).toContain('schema --list');
  }, 30_000);
});
