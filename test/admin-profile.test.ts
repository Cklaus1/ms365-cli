import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { spawnSync } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { Command } from 'commander';
import { AuthManager } from '../src/auth.js';
import { registerAuthCommands } from '../src/commands/auth.js';
import { readProfile, saveProfile, profileType } from '../src/profiles.js';
import { assertConsented } from '../src/preflight.js';
import { azDeps } from '../src/grant-scopes.js';
import { CliError, ExitCode } from '../src/errors.js';
import { runCommand } from './helpers/cli.js';

// ID-7 / ID-8: profile types and admin profiles whose token is borrowed from the Azure CLI.

function jwt(payload: Record<string, unknown>): string {
  const b64 = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${b64({ alg: 'RS256', typ: 'JWT' })}.${b64(payload)}.sig`;
}
const TENANT = 'contoso.onmicrosoft.com';
const AZ_SCP = 'Directory.AccessAsUser.All Group.ReadWrite.All User.ReadWrite.All AuditLog.Read.All';
const azToken = (expSec = Math.floor(Date.now() / 1000) + 3600) =>
  jwt({ scp: AZ_SCP, aud: 'https://graph.microsoft.com', tid: 'tid-1', upn: 'admin@contoso.com', exp: expSec });

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-admin-'));
  vi.stubEnv('MS365_CLI_CONFIG_DIR', dir);
  vi.stubEnv('MS365_CLI_TOKEN', '');
  vi.stubEnv('MS365_CLI_CLIENT_ID', '');
});
afterEach(() => { vi.unstubAllEnvs(); vi.restoreAllMocks(); });

const writeRaw = (name: string, data: unknown) => {
  fs.mkdirSync(path.join(dir, 'profiles'), { recursive: true });
  fs.writeFileSync(path.join(dir, 'profiles', `${name}.json`), JSON.stringify(data));
};

describe('ID-7: Profile.type', () => {
  it('defaults to user; type round-trips; unknown types are invalid_profile', () => {
    saveProfile('u', { tenantId: 't' });
    expect(profileType(readProfile('u'))).toBe('user');
    saveProfile('a', { type: 'admin', tenantId: TENANT });
    expect(readProfile('a')).toEqual({ type: 'admin', tenantId: TENANT });
    writeRaw('bad', { type: 'robot' });
    expect(() => readProfile('bad')).toThrow(expect.objectContaining({ code: 'invalid_profile' }));
  });

  it('admin requires a real tenant and rejects account and scopes', () => {
    for (const [name, data] of [
      ['nt', { type: 'admin' }],
      ['common', { type: 'admin', tenantId: 'common' }],
      ['orgs', { type: 'admin', tenantId: 'organizations' }],
      ['acct', { type: 'admin', tenantId: TENANT, account: 'a@x.com' }],
      ['scp', { type: 'admin', tenantId: TENANT, scopes: ['User.Read'] }],
    ] as const) {
      writeRaw(name, data);
      expect(() => readProfile(name), name).toThrow(expect.objectContaining({ code: 'invalid_profile' }));
    }
    // clientId/orgMode are allowed (ignored with a warning).
    writeRaw('ok', { type: 'admin', tenantId: TENANT, clientId: 'c', orgMode: true });
    expect(readProfile('ok').type).toBe('admin');
  });

  const reg = (program: Command) => registerAuthCommands(program, () => { throw new Error('no AuthManager needed'); });

  it('`auth profile set --type admin --tenant <t>` saves; without a tenant or with --scopes exits 2', async () => {
    let r = await runCommand(reg, ['auth', 'profile', 'set', 'adm', '--type', 'admin', '--tenant', TENANT]);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(readProfile('adm')).toEqual({ type: 'admin', tenantId: TENANT });

    r = await runCommand(reg, ['auth', 'profile', 'set', 'adm2', '--type', 'admin']);
    expect(r.exitCode).toBe(ExitCode.USAGE);
    expect(r.stderr).toContain('needs a tenant');
    expect(fs.existsSync(path.join(dir, 'profiles', 'adm2.json'))).toBe(false);

    r = await runCommand(reg, ['auth', 'profile', 'set', 'adm', '--scopes', 'User.Read']);
    expect(r.exitCode).toBe(ExitCode.USAGE);
    expect(readProfile('adm')).toEqual({ type: 'admin', tenantId: TENANT });

    r = await runCommand(reg, ['auth', 'profile', 'set', 'x', '--type', 'robot']);
    expect(r.exitCode).toBe(ExitCode.USAGE);
  });

  it('`auth profile show` warns that clientId/orgMode are ignored by an admin profile', async () => {
    saveProfile('adm', { type: 'admin', tenantId: TENANT, clientId: 'c', orgMode: true });
    const r = await runCommand(reg, ['auth', 'profile', 'show', 'adm', '-o', 'json']);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.stderr).toMatch(/"clientId" is ignored/);
    expect(r.stderr).toMatch(/"orgMode" is ignored/);
    expect(JSON.parse(r.stdout).type).toBe('admin');
  });
});

describe('ID-8: az token source', () => {
  const az = (impl: (args: string[]) => ReturnType<typeof azDeps.run>) =>
    vi.spyOn(azDeps, 'run').mockImplementation(impl);
  const okRun = (token = azToken()) => () => ({ status: 0, stdout: JSON.stringify({ accessToken: token, tenant: 'tid-1' }), stderr: '', missing: false });
  const adminManager = () => new AuthManager([], { source: 'az', tenantId: TENANT, profileName: 'adm' });

  it('needs no client ID, borrows from az --tenant, keeps the token in memory and writes nothing', async () => {
    const run = az(okRun());
    const m = adminManager();
    await m.initialize();
    const t1 = await m.getToken();
    const t2 = await m.getToken();
    expect(t1).toBe(t2);
    expect(run).toHaveBeenCalledTimes(1);
    expect(run.mock.calls[0][0]).toEqual(expect.arrayContaining(['account', 'get-access-token', '--tenant', TENANT]));
    expect(await m.grantedScopes()).toEqual(AZ_SCP.split(' '));
    expect(fs.readdirSync(dir)).toEqual([]);
  });

  it('re-borrows from az once the token is within 5 minutes of expiry', async () => {
    const run = az(okRun(azToken(Math.floor(Date.now() / 1000) + 120)));
    const m = adminManager();
    await m.getToken();
    await m.getToken();
    expect(run).toHaveBeenCalledTimes(2);
  });

  it('az missing → exit 3 az_required with the az login fix; az not logged in → same', async () => {
    az(() => ({ status: null, stdout: '', stderr: '', missing: true }));
    await expect(adminManager().getToken()).rejects.toMatchObject({
      code: 'az_required', exitCode: ExitCode.AUTH, fix: expect.stringContaining(`az login --tenant ${TENANT}`),
    });
    vi.restoreAllMocks();
    az(() => ({ status: 1, stdout: '', stderr: "ERROR: Please run 'az login' to setup account.", missing: false }));
    await expect(adminManager().getToken()).rejects.toMatchObject({
      code: 'az_required', exitCode: ExitCode.AUTH, fix: `az login --tenant ${TENANT}`,
    });
  });

  it('login/logout/listAccounts/select/remove on an az manager are usage errors', async () => {
    const m = adminManager();
    for (const call of [() => m.login(), () => m.logout(), () => m.listAccounts(), () => m.selectAccount('x'), () => m.removeAccount('x')]) {
      await expect(call()).rejects.toMatchObject({ code: 'usage', exitCode: ExitCode.USAGE });
    }
  });

  const regAdmin = (m: AuthManager) => (program: Command) => registerAuthCommands(program, () => m, () => 'adm');

  it.each([
    [['auth', 'login']], [['auth', 'logout']], [['auth', 'accounts']], [['auth', 'select', 'x']], [['auth', 'remove', 'x']],
  ])('`ms365 %j` under an admin profile exits 2 (also under --dry-run)', async (argv) => {
    const run = az(okRun());
    for (const dryRun of [false, true]) {
      const r = await runCommand(regAdmin(adminManager()), argv, { dryRun });
      expect(r.exitCode, r.stderr).toBe(ExitCode.USAGE);
      expect(r.stderr).toContain('az login');
    }
    expect(run).not.toHaveBeenCalled();
  });

  it('`auth status` and `auth scopes` report the az identity and its scp', async () => {
    az(okRun());
    let r = await runCommand(regAdmin(adminManager()), ['auth', 'status']);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.stdout).toContain('admin@contoso.com');
    expect(r.stdout).toContain('Azure CLI');
    expect(r.stdout).toContain('Group.ReadWrite.All');
    r = await runCommand(regAdmin(adminManager()), ['auth', 'scopes', '-o', 'json']);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(JSON.parse(r.stdout)).toMatchObject({ upn: 'admin@contoso.com', scopes: AZ_SCP.split(' '), source: 'azure-cli', profile: 'adm' });
  });
});

describe('ID-8: profile-aware missing_scope', () => {
  it('names the admin profile and points at a user profile', () => {
    let err: CliError | undefined;
    try { assertConsented('mail list', [{ scopes: ['Mail.Read'] }], AZ_SCP.split(' '), 'adm'); } catch (e) { err = e as CliError; }
    expect(err).toMatchObject({ code: 'missing_scope', exitCode: ExitCode.PERMISSION });
    expect(err!.message).toBe('"mail list" needs Mail.Read; the Azure CLI token used by admin profile "adm" does not include it — use a user profile (ms365 --org-mode auth login) for this command.');
    expect(() => assertConsented('groups list', [{ scopes: ['Group.Read.All'] }], AZ_SCP.split(' '), 'adm')).not.toThrow();
  });
});

describe('ID-7/ID-8 end to end through index.ts (stub az on PATH)', () => {
  const root = path.resolve(__dirname, '..');
  const tsx = path.join(root, 'node_modules', '.bin', 'tsx');
  function cli(args: string[], azScript: string) {
    const bin = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-azstub-'));
    fs.writeFileSync(path.join(bin, 'az'), `#!/bin/sh\necho "$@" >> "${path.join(bin, 'args')}"\n${azScript}\n`, { mode: 0o755 });
    const r = spawnSync(tsx, [path.join(root, 'src/index.ts'), ...args], {
      cwd: dir, encoding: 'utf8', timeout: 30_000, input: '',
      env: { ...process.env, PATH: `${bin}${path.delimiter}${process.env.PATH}`, HOME: dir, MS365_CLI_CONFIG_DIR: dir,
        MS365_CLI_TOKEN: '', MS365_CLI_CLIENT_ID: '', MS365_CLI_ACCOUNT: '', MS365_CLI_PROFILE: '', MS365_CLI_API_VERSION: '' },
    });
    const azArgs = fs.existsSync(path.join(bin, 'args')) ? fs.readFileSync(path.join(bin, 'args'), 'utf8') : '';
    return { ...r, azArgs };
  }
  const azOk = `echo '${JSON.stringify({ accessToken: azToken(), tenant: 'tid-1' })}'`;

  it('a command outside the az scp exits 4 with the profile-aware message; no token file is written', () => {
    saveProfile('adm', { type: 'admin', tenantId: TENANT });
    const r = cli(['--profile', 'adm', '--error-format', 'json', 'mail', 'list'], azOk);
    expect(r.status, r.stderr).toBe(ExitCode.PERMISSION);
    expect(r.stderr).toContain('"code":"missing_scope"');
    expect(r.stderr).toContain('the Azure CLI token used by admin profile \\"adm\\" does not include it');
    expect(r.azArgs).toContain(`--tenant ${TENANT}`);
    expect(fs.existsSync(path.join(dir, 'token-cache.json'))).toBe(false);
  }, 30_000);

  it('az not signed in → exit 3 az_required', () => {
    saveProfile('adm', { type: 'admin', tenantId: TENANT });
    const r = cli(['--profile', 'adm', '--error-format', 'json', 'auth', 'scopes'], `echo "ERROR: Please run 'az login' to setup account." >&2; exit 1`);
    expect(r.status, r.stderr).toBe(ExitCode.AUTH);
    expect(r.stderr).toContain('"code":"az_required"');
    expect(r.stderr).toContain(`az login --tenant ${TENANT}`);
  }, 30_000);

  it('`auth login` under an admin profile exits 2 without calling az', () => {
    saveProfile('adm', { type: 'admin', tenantId: TENANT });
    const r = cli(['--profile', 'adm', 'auth', 'login'], azOk);
    expect(r.status, r.stderr).toBe(ExitCode.USAGE);
    expect(r.stderr).toContain('az login');
    expect(r.azArgs).toBe('');
  }, 30_000);
});
