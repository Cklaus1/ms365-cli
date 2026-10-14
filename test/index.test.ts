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

  it('--dry-run needs no client ID and writes nothing to the config dir (DR-16, F23)', () => {
    const cfg = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-idx-cfg-'));
    const r = cli(['--dry-run', 'mail', 'send', '--to', 'a@b.com', '--subject', 's', '--body', 'b'],
      { MS365_CLI_CLIENT_ID: '', MS365_CLI_CONFIG_DIR: cfg });
    expect(r.status, r.stderr).toBe(0);
    expect(r.stderr).toContain('[dry-run] POST https://graph.microsoft.com/v1.0/me/sendMail');
    expect(fs.readdirSync(cfg)).toEqual([]);
  }, 30_000);

  it('global flags work after the subcommand too', () => {
    const r = cli(['user', 'me', '--dry-run']);
    expect(r.status, r.stderr).toBe(0);
    expect(r.stderr).toContain('[dry-run] GET');
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

describe('Wave A wiring (T30, T37, T39 pre-flight)', () => {
  const jwt = (claims: Record<string, unknown>) =>
    ['{"alg":"none"}', JSON.stringify(claims)].map(p => Buffer.from(p).toString('base64url')).join('.') + '.sig';

  it('fails fast with missing_scope (exit 4) before calling Graph', () => {
    const token = jwt({ scp: 'User.Read', aud: 'https://graph.microsoft.com', exp: Math.floor(Date.now() / 1000) + 3600 });
    const r = cli(['--error-format', 'json', 'mail', 'list'], { MS365_CLI_TOKEN: token });
    expect(r.status, r.stderr).toBe(4);
    expect(JSON.parse(r.stderr.trim().split('\n').pop()!).error.code).toBe('missing_scope');
  }, 30_000);

  it('describe --json and audit path run with no client ID', () => {
    const noClientId = { MS365_CLI_CLIENT_ID: '' };
    const d = cli(['describe', '--json'], noClientId);
    expect(d.status, d.stderr).toBe(0);
    const paths = JSON.parse(d.stdout).commands.map((c: { path: string }) => c.path);
    expect(paths).toEqual(expect.arrayContaining(['meetings transcript', 'watch', 'audit tail', 'calendar find-time']));
    const a = cli(['audit', 'path'], noClientId);
    expect(a.status, a.stderr).toBe(0);
    expect(a.stdout.trim()).toMatch(/audit\.log$/);
  }, 30_000);
});

describe('lazy MSAL import (X-1)', () => {
  // A resolve hook (module.register) that writes a marker to stderr whenever anything resolves
  // @azure/msal-node. Registered after tsx, so it runs first in the hook chain.
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-probe-'));
  const hooks = 'export async function resolve(s, c, next) {\n'
    + "  if (s === '@azure/msal-node' || s.startsWith('@azure/msal-node/')) "
    + "(await import('node:fs')).writeSync(2, '[probe] msal-node loaded\\n');\n"
    + '  return next(s, c);\n}\n';
  fs.writeFileSync(path.join(dir, 'hooks.mjs'), hooks);
  const probe = path.join(dir, 'probe.mjs');
  fs.writeFileSync(probe, "import { register } from 'node:module';\n"
    + `register(${JSON.stringify(new URL('file://' + path.join(dir, 'hooks.mjs')).href)});\n`);
  const tsxLoader = path.join(root, 'node_modules', 'tsx', 'dist', 'esm', 'index.mjs');
  function probed(args: string[]) {
    return spawnSync(process.execPath, ['--import', tsxLoader, '--import', probe, path.join(root, 'src/index.ts'), ...args], {
      cwd: home, encoding: 'utf8', timeout: 30_000, input: '',
      env: { ...process.env, HOME: home, MS365_CLI_CONFIG_DIR: path.join(home, 'cfg'), MS365_CLI_TOKEN: '',
        MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_API_VERSION: '' },
    });
  }

  it.each([['--help'], ['--version'], ['describe', '--json'], ['--dry-run', 'mail', 'list'], ['completion', 'bash'],
    ['schema', '--list']])('%s does not load @azure/msal-node', (...args) => {
    const r = probed(args);
    expect(r.status, r.stderr).toBe(0);
    expect(r.stderr).not.toContain('[probe] msal-node loaded');
  }, 30_000);

  it('auth status loads @azure/msal-node (the probe works)', () => {
    const r = probed(['auth', 'status']);
    expect(r.stderr).toContain('[probe] msal-node loaded');
    expect(r.status).toBe(3);
  }, 30_000);
});
