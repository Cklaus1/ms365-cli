import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { spawnSync } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { Command } from 'commander';
import { AuthManager, decodeAccessToken, tokenInfo } from '../src/auth.js';
import { registerAuthCommands } from '../src/commands/auth.js';
import { openBrowserCommand } from '../src/commands/setup.js';
import { readProfile, saveProfile, listProfiles, deleteProfile, validateProfileName } from '../src/profiles.js';
import { CliError, ExitCode } from '../src/errors.js';
import { runCommand } from './helpers/cli.js';

const execMock = vi.hoisted(() => vi.fn());
vi.mock('child_process', async (importOriginal) => {
  const m = await importOriginal<typeof import('child_process')>();
  return { ...m, default: { ...m.default, exec: execMock }, exec: execMock };
});

const CID = '11111111-1111-1111-1111-111111111111';

function jwt(payload: Record<string, unknown>): string {
  const b64 = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${b64({ alg: 'RS256', typ: 'JWT' })}.${b64(payload)}.sig`;
}
const EXP = 2_000_000_000; // 2033-05-18
const TOKEN = jwt({ scp: 'Mail.Read User.Read  Files.ReadWrite', aud: 'https://graph.microsoft.com', tid: 'tenant-1', upn: 'ada@contoso.com', exp: EXP });

function cacheJson(...ns: number[]): string {
  const Account: Record<string, unknown> = {};
  const RefreshToken: Record<string, unknown> = {};
  const IdToken: Record<string, unknown> = {};
  for (const n of ns) {
    const hid = `uid${n}.utid`;
    Account[`${hid}-login.microsoftonline.com-utid`] = {
      home_account_id: hid, environment: 'login.microsoftonline.com', realm: 'utid',
      local_account_id: `uid${n}`, username: `user${n}@contoso.com`, authority_type: 'MSSTS',
    };
    RefreshToken[`${hid}-login.microsoftonline.com-refreshtoken-${CID}--`] = {
      home_account_id: hid, environment: 'login.microsoftonline.com', credential_type: 'RefreshToken', client_id: CID, secret: `rt-${n}`,
    };
    // MSAL takes the account's username from its ID token.
    IdToken[`${hid}-login.microsoftonline.com-idtoken-${CID}-utid-`] = {
      home_account_id: hid, environment: 'login.microsoftonline.com', credential_type: 'IdToken', client_id: CID, realm: 'utid',
      secret: jwt({ preferred_username: `user${n}@contoso.com`, oid: `uid${n}`, tid: 'utid' }),
    };
  }
  return JSON.stringify({ Account, IdToken, AccessToken: {}, RefreshToken, AppMetadata: {} });
}

type Msal = { acquireTokenSilent: Function; acquireTokenByDeviceCode: Function; acquireTokenInteractive: Function };
// The PublicClientApplication is built on first use (X-1, lazy MSAL import).
const msal = (m: AuthManager) => (m as unknown as { app(): Promise<Msal> }).app();
const register = (m: AuthManager) => (program: Command) => registerAuthCommands(program, () => m);

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-authfeat-'));
  vi.stubEnv('MS365_CLI_CONFIG_DIR', dir);
  vi.stubEnv('MS365_CLI_TOKEN', '');
  vi.stubEnv('MS365_CLI_CLIENT_ID', CID);
  execMock.mockReset();
});
afterEach(() => { vi.unstubAllEnvs(); vi.restoreAllMocks(); });

describe('T39: access-token introspection (§6)', () => {
  it('decodes the JWT payload without verifying the signature', () => {
    expect(decodeAccessToken(TOKEN)).toMatchObject({ tid: 'tenant-1', upn: 'ada@contoso.com' });
  });

  it('tokenInfo extracts scopes, audience, tenant, user and expiry', () => {
    expect(tokenInfo(TOKEN, new Date('2030-01-01T00:00:00Z'))).toEqual({
      scopes: ['Mail.Read', 'User.Read', 'Files.ReadWrite'],
      audience: 'https://graph.microsoft.com', tenantId: 'tenant-1', user: 'ada@contoso.com',
      expiresOn: new Date(EXP * 1000).toISOString(), expired: false,
    });
  });

  it('falls back to preferred_username and reports expiry', () => {
    const info = tokenInfo(jwt({ scp: 'User.Read', preferred_username: 'bob@x.com', exp: 1 }));
    expect(info.user).toBe('bob@x.com');
    expect(info.expired).toBe(true);
  });

  it('an opaque (non-JWT) token is a clear error', () => {
    expect(() => decodeAccessToken('EwBwA8l6BAAU-opaque')).toThrow(/not a JWT/);
    expect(() => decodeAccessToken('a.!!!.c')).toThrow(/not a JWT/);
  });

  it('`auth scopes -o json` prints the claims of MS365_CLI_TOKEN', async () => {
    vi.stubEnv('MS365_CLI_TOKEN', TOKEN);
    const r = await runCommand(register(new AuthManager([])), ['auth', 'scopes', '-o', 'json']);
    expect(r.exitCode, r.stderr).toBe(0);
    const out = JSON.parse(r.stdout);
    expect(out.scopes).toEqual(['Mail.Read', 'User.Read', 'Files.ReadWrite']);
    expect(out).toMatchObject({ audience: 'https://graph.microsoft.com', tenantId: 'tenant-1', user: 'ada@contoso.com' });
  });

  it('`auth scopes` with no login exits 3', async () => {
    const r = await runCommand(register(new AuthManager([])), ['auth', 'scopes']);
    expect(r.exitCode).toBe(ExitCode.AUTH);
  });

  it('grantedScopes() returns the scp claim of the current token', async () => {
    vi.stubEnv('MS365_CLI_TOKEN', TOKEN);
    expect(await new AuthManager([]).grantedScopes()).toEqual(['Mail.Read', 'User.Read', 'Files.ReadWrite']);
  });
});

describe('T40: --account (PM-8, D21)', () => {
  beforeEach(() => {
    fs.writeFileSync(path.join(dir, 'token-cache.json'), cacheJson(1, 2));
    fs.writeFileSync(path.join(dir, 'selected-account.json'), JSON.stringify({ accountId: 'uid1.utid' }));
  });

  async function tokenFor(account: string) {
    const m = new AuthManager(['User.Read'], { account });
    await m.initialize();
    const silent = vi.spyOn(await msal(m), 'acquireTokenSilent').mockResolvedValue({ accessToken: 'at', expiresOn: new Date(Date.now() + 3600e3) } as never);
    await m.getToken();
    return (silent.mock.calls[0][0] as { account: { homeAccountId: string } }).account.homeAccountId;
  }

  it('matches the username case-insensitively', async () => {
    expect(await tokenFor('USER2@Contoso.com')).toBe('uid2.utid');
  });

  it('matches the homeAccountId', async () => {
    expect(await tokenFor('uid2.utid')).toBe('uid2.utid');
  });

  it('does not change the saved selection', async () => {
    await tokenFor('user2@contoso.com');
    expect(JSON.parse(fs.readFileSync(path.join(dir, 'selected-account.json'), 'utf8')).accountId).toBe('uid1.utid');
  });

  it('an unknown account exits 3 pointing at `ms365 auth accounts`', async () => {
    const m = new AuthManager(['User.Read'], { account: 'nobody@contoso.com' });
    await m.initialize();
    const err = await m.getToken().catch(e => e) as CliError;
    expect(err).toBeInstanceOf(CliError);
    expect(err.exitCode).toBe(ExitCode.AUTH);
    expect(err.fix).toContain('ms365 auth accounts');
  });
});

describe('VT5: `auth accounts` marks the account in effect (V-14)', () => {
  beforeEach(() => { fs.writeFileSync(path.join(dir, 'token-cache.json'), cacheJson(1, 2)); });
  const active = (stdout: string) => stdout.split('\n').filter(l => l.includes('(active)')).map(l => l.trim());

  it('the --account override wins over the saved selection', async () => {
    fs.writeFileSync(path.join(dir, 'selected-account.json'), JSON.stringify({ accountId: 'uid1.utid' }));
    const r = await runCommand(register(new AuthManager([], { account: 'USER2@contoso.com' })), ['auth', 'accounts']);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(active(r.stdout)).toEqual(['user2@contoso.com (active)']);
  });

  it('the saved selection is marked when there is no override', async () => {
    fs.writeFileSync(path.join(dir, 'selected-account.json'), JSON.stringify({ accountId: 'uid2.utid' }));
    const r = await runCommand(register(new AuthManager([])), ['auth', 'accounts']);
    expect(active(r.stdout)).toEqual(['user2@contoso.com (active)']);
  });

  it('with nothing saved, the first cached account (the one commands use) is marked', async () => {
    const m = new AuthManager([]);
    const r = await runCommand(register(m), ['auth', 'accounts']);
    expect(active(r.stdout)).toHaveLength(1);
    await m.initialize();
    expect(active(r.stdout)[0]).toBe(`${(await m.getCurrentAccount())!.username} (active)`);
  });

  it('an override that matches nothing marks none and warns on stderr', async () => {
    fs.writeFileSync(path.join(dir, 'selected-account.json'), JSON.stringify({ accountId: 'uid1.utid' }));
    const r = await runCommand(register(new AuthManager([], { account: 'nobody@contoso.com' })), ['auth', 'accounts']);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(active(r.stdout)).toEqual([]);
    expect(r.stdout).toContain('user1@contoso.com');
    expect(r.stderr).toContain('nobody@contoso.com');
  });

  it('activeAccountId() follows getCurrentAccount() precedence and returns null instead of throwing', async () => {
    fs.writeFileSync(path.join(dir, 'selected-account.json'), JSON.stringify({ accountId: 'uid1.utid' }));
    const m = new AuthManager([], { account: 'uid2.utid' });
    await m.initialize();
    expect(await m.activeAccountId()).toBe('uid2.utid');
    const none = new AuthManager([], { account: 'nobody@contoso.com' });
    await none.initialize();
    expect(await none.activeAccountId()).toBeNull();
  });
});

describe('[ADDED] `auth status` / `auth scopes` under --dry-run stay offline (DR-22 class)', () => {
  for (const cmd of ['status', 'scopes']) {
    it(`--dry-run auth ${cmd} prints the note, exits 0 and never builds an AuthManager`, async () => {
      const getAuth = vi.fn(() => { throw new Error('AuthManager built'); });
      const r = await runCommand((p: Command) => registerAuthCommands(p, getAuth as never), ['auth', cmd], { dryRun: true });
      expect(r.exitCode, r.stderr).toBe(0);
      expect(getAuth).not.toHaveBeenCalled();
      expect(r.stderr).toContain('[dry-run]');
      expect(r.stdout).toBe('');
    });
  }
});

describe('T40: profiles (PM-8, D21)', () => {
  it('validates names so no path can escape the profiles dir', () => {
    for (const bad of ['', '../x', 'a/b', 'A', 'x.json', '..']) {
      expect(() => validateProfileName(bad), bad).toThrow(CliError);
    }
    expect(() => validateProfileName('work-1_a')).not.toThrow();
  });

  it('save → read → list → delete round-trip under configDir/profiles', () => {
    saveProfile('work', { clientId: 'c', tenantId: 't', scopes: ['User.Read'], orgMode: true, account: 'a@x.com' });
    expect(fs.existsSync(path.join(dir, 'profiles', 'work.json'))).toBe(true);
    expect(readProfile('work')).toEqual({ clientId: 'c', tenantId: 't', scopes: ['User.Read'], orgMode: true, account: 'a@x.com' });
    saveProfile('home', {});
    expect(listProfiles()).toEqual(['home', 'work']);
    expect(deleteProfile('work')).toBe(true);
    expect(deleteProfile('work')).toBe(false);
    expect(listProfiles()).toEqual(['home']);
  });

  it('reading a missing profile exits 5; a malformed one is an error', () => {
    expect(() => readProfile('nope')).toThrow(expect.objectContaining({ exitCode: ExitCode.NOT_FOUND }));
    fs.mkdirSync(path.join(dir, 'profiles'), { recursive: true });
    fs.writeFileSync(path.join(dir, 'profiles', 'bad.json'), '{"scopes":"User.Read"}');
    expect(() => readProfile('bad')).toThrow(CliError);
  });

  const m = () => new AuthManager([]);

  it('`auth profile set` creates and then merges into a profile; `show` prints it', async () => {
    let r = await runCommand(register(m()), ['auth', 'profile', 'set', 'work', '--client-id', 'cid', '--tenant-id', 'tid',
      '--scopes', 'User.Read, Mail.Read', '--org-mode', '--account', 'ada@contoso.com']);
    expect(r.exitCode, r.stderr).toBe(0);
    r = await runCommand(register(m()), ['auth', 'profile', 'set', 'work', '--tenant-id', 'tid2', '--no-org-mode']);
    expect(r.exitCode, r.stderr).toBe(0);
    r = await runCommand(register(m()), ['auth', 'profile', 'show', 'work', '-o', 'json']);
    expect(JSON.parse(r.stdout)).toEqual({ clientId: 'cid', tenantId: 'tid2', scopes: ['User.Read', 'Mail.Read'], orgMode: false, account: 'ada@contoso.com' });
  });

  it('`auth profile list -o json` lists profile names', async () => {
    saveProfile('b', { tenantId: 't' });
    saveProfile('a', {});
    const r = await runCommand(register(m()), ['auth', 'profile', 'list', '-o', 'json']);
    expect(JSON.parse(r.stdout)).toEqual([{ name: 'a' }, { name: 'b', tenantId: 't' }]);
  });

  it('`auth profile set` with a traversal name exits 2 and writes nothing', async () => {
    const r = await runCommand(register(m()), ['auth', 'profile', 'set', '../evil', '--tenant-id', 't']);
    expect(r.exitCode).toBe(ExitCode.USAGE);
    expect(fs.existsSync(path.join(dir, 'evil.json'))).toBe(false);
    expect(fs.existsSync(path.join(dir, 'profiles'))).toBe(false);
  });

  it('`auth profile delete` confirms; refuses without a TTY; 5 when missing', async () => {
    saveProfile('work', {});
    let r = await runCommand(register(m()), ['auth', 'profile', 'delete', 'work']);
    expect(r.exitCode).toBe(ExitCode.CONFIRMATION_REFUSED);
    expect(listProfiles()).toEqual(['work']);
    r = await runCommand(register(m()), ['auth', 'profile', 'delete', 'work'], { tty: true, answer: 'y' });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(listProfiles()).toEqual([]);
    r = await runCommand(register(m()), ['auth', 'profile', 'delete', 'work'], { tty: true, answer: 'y' });
    expect(r.exitCode).toBe(ExitCode.NOT_FOUND);
  });
});

describe('T41: auth login --browser (PM-2)', () => {
  const AUTH_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=x&redirect_uri=http%3A%2F%2Flocalhost%3A5000';

  async function setup() {
    const m = new AuthManager(['User.Read', 'Mail.ReadWrite']);
    vi.spyOn(m, 'verifyLogin').mockResolvedValue({ success: true, user: { displayName: 'Ada', email: 'ada@contoso.com' } });
    const interactive = vi.spyOn(await msal(m), 'acquireTokenInteractive').mockImplementation(async (req: { openBrowser: (u: string) => Promise<void> }) => {
      await req.openBrowser(AUTH_URL);
      return { accessToken: 'at', expiresOn: new Date(Date.now() + 3600e3), account: { homeAccountId: 'uid9.utid', username: 'ada@contoso.com' } };
    });
    const device = vi.spyOn(await msal(m), 'acquireTokenByDeviceCode').mockResolvedValue({ accessToken: 'at', account: { homeAccountId: 'uid8.utid' } } as never);
    return { m, interactive, device };
  }

  it('uses acquireTokenInteractive with the login scopes and opens the system browser at the auth URL', async () => {
    const { m, interactive, device } = await setup();
    const r = await runCommand(register(m), ['auth', 'login', '--browser']);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(device).not.toHaveBeenCalled();
    const req = interactive.mock.calls[0][0] as { scopes: string[]; successTemplate: string; errorTemplate: string };
    expect(req.scopes).toEqual(['User.Read', 'Mail.ReadWrite']);
    expect(req.successTemplate).toMatch(/<html/i);
    expect(req.errorTemplate).toMatch(/<html/i);
    expect(execMock).toHaveBeenCalledWith(openBrowserCommand(AUTH_URL), expect.any(Function));
    expect(r.stderr).toContain(AUTH_URL);
    expect(r.stderr).toContain('Logged in as Ada');
  });

  it('selects and saves the account that signed in', async () => {
    const { m } = await setup();
    await runCommand(register(m), ['auth', 'login', '--browser']);
    expect(m.getSelectedAccountId()).toBe('uid9.utid');
    expect(JSON.parse(fs.readFileSync(path.join(dir, 'selected-account.json'), 'utf8')).accountId).toBe('uid9.utid');
  });

  it('device code stays the default', async () => {
    const { m, interactive, device } = await setup();
    const r = await runCommand(register(m), ['auth', 'login']);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(interactive).not.toHaveBeenCalled();
    expect(device).toHaveBeenCalled();
    expect(m.getSelectedAccountId()).toBe('uid8.utid');
  });
});

describe('T40: global --account / --profile wiring in index.ts', () => {
  const root = path.resolve(__dirname, '..');
  const tsx = path.join(root, 'node_modules', '.bin', 'tsx');
  function cli(args: string[], env: Record<string, string> = {}) {
    return spawnSync(tsx, [path.join(root, 'src/index.ts'), ...args], {
      cwd: dir, encoding: 'utf8', timeout: 30_000, input: '',
      env: { ...process.env, HOME: dir, MS365_CLI_CONFIG_DIR: dir, MS365_CLI_TOKEN: '', MS365_CLI_CLIENT_ID: '',
        MS365_CLI_ACCOUNT: '', MS365_CLI_PROFILE: '', MS365_CLI_API_VERSION: '', ...env },
    });
  }

  it('an unknown --profile exits 5 with a fix', () => {
    const r = cli(['--profile', 'missing', '--dry-run', 'user', 'me']);
    expect(r.status).toBe(ExitCode.NOT_FOUND);
    expect(r.stderr).toContain('ms365 auth profile list');
  }, 30_000);

  it('--profile (and MS365_CLI_PROFILE) supplies the client ID', () => {
    saveProfile('work', { clientId: CID });
    // Live runs (dry-run needs no client ID, DR-16): without the profile there is none; with it, only the login is missing.
    expect(cli(['--error-format', 'json', 'user', 'me']).stderr).toContain('"code":"no_client_id"');
    const r = cli(['--profile', 'work', '--error-format', 'json', 'user', 'me']);
    expect(r.status).toBe(ExitCode.AUTH);
    expect(r.stderr).toContain('"code":"auth_required"');
    expect(cli(['--error-format', 'json', 'user', 'me'], { MS365_CLI_PROFILE: 'work' }).stderr).toContain('"code":"auth_required"');
  }, 60_000);

  it('an unknown --account (or MS365_CLI_ACCOUNT) exits 3 pointing at `ms365 auth accounts`', () => {
    fs.writeFileSync(path.join(dir, 'token-cache.json'), cacheJson(1));
    let r = cli(['--account', 'nobody@contoso.com', 'user', 'me'], { MS365_CLI_CLIENT_ID: CID });
    expect(r.status).toBe(ExitCode.AUTH);
    expect(r.stderr).toContain('ms365 auth accounts');
    r = cli(['user', 'me'], { MS365_CLI_CLIENT_ID: CID, MS365_CLI_ACCOUNT: 'nobody@contoso.com' });
    expect(r.status).toBe(ExitCode.AUTH);
    expect(r.stderr).toContain('ms365 auth accounts');
  }, 60_000);

  it('a profile account is used when no --account is given', () => {
    fs.writeFileSync(path.join(dir, 'token-cache.json'), cacheJson(1));
    saveProfile('p', { clientId: CID, account: 'nobody@contoso.com' });
    const r = cli(['--profile', 'p', 'user', 'me']);
    expect(r.status).toBe(ExitCode.AUTH);
    expect(r.stderr).toContain('nobody@contoso.com');
  }, 30_000);
});
