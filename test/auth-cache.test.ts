import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { spawn } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { AuthManager, FileTokenCachePlugin, writeFileAtomic, GRAPH_DEFAULT_SCOPE } from '../src/auth.js';
import { CliError, ExitCode } from '../src/errors.js';

const CID = '11111111-1111-1111-1111-111111111111';

function account(n: number) {
  const hid = `uid${n}.utid`;
  return {
    account: {
      [`${hid}-login.microsoftonline.com-utid`]: {
        home_account_id: hid, environment: 'login.microsoftonline.com', realm: 'utid',
        local_account_id: `uid${n}`, username: `user${n}@contoso.com`, authority_type: 'MSSTS',
      },
    },
    rt: {
      [`${hid}-login.microsoftonline.com-refreshtoken-${CID}--`]: {
        home_account_id: hid, environment: 'login.microsoftonline.com', credential_type: 'RefreshToken',
        client_id: CID, secret: `rt-secret-${n}`,
      },
    },
  };
}
function cacheJson(...ns: number[]): string {
  const a = ns.map(account);
  return JSON.stringify({
    Account: Object.assign({}, ...a.map(x => x.account)), IdToken: {}, AccessToken: {},
    RefreshToken: Object.assign({}, ...a.map(x => x.rt)), AppMetadata: {},
  });
}

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-auth-'));
  vi.stubEnv('MS365_CLI_CONFIG_DIR', dir);
  vi.stubEnv('MS365_CLI_TOKEN', '');
  vi.stubEnv('MS365_MCP_OAUTH_TOKEN', '');
  vi.stubEnv('MS365_CLI_CLIENT_ID', CID);
});
afterEach(() => { vi.unstubAllEnvs(); vi.restoreAllMocks(); });
const cacheFile = () => path.join(dir, 'token-cache.json');
const msal = (m: AuthManager) => (m as unknown as { msalApp: { acquireTokenSilent: Function; acquireTokenByDeviceCode: Function } }).msalApp;

describe('token cache persistence [critical-path gate]', () => {
  it('round-trips: load → mutate → persist → a fresh manager sees exactly the persisted state', async () => {
    fs.writeFileSync(cacheFile(), cacheJson(1, 2));
    const a = new AuthManager(['User.Read']);
    await a.initialize();
    expect((await a.listAccounts()).map(x => x.homeAccountId).sort()).toEqual(['uid1.utid', 'uid2.utid']);

    expect(await a.removeAccount('uid1.utid')).toBe(true);
    const onDisk = JSON.parse(fs.readFileSync(cacheFile(), 'utf8'));
    expect(Object.keys(onDisk.Account)).toHaveLength(1);
    expect(JSON.stringify(onDisk.RefreshToken)).toContain('rt-secret-2');
    expect(JSON.stringify(onDisk.RefreshToken)).not.toContain('rt-secret-1');
    expect(fs.statSync(cacheFile()).mode & 0o777).toBe(0o600);

    const b = new AuthManager(['User.Read']);
    await b.initialize();
    expect((await b.listAccounts()).map(x => x.homeAccountId)).toEqual(['uid2.utid']);
  });

  it('plugin writes only when MSAL reports a change (C-4: silent refreshes persist)', async () => {
    const plugin = new FileTokenCachePlugin();
    await plugin.afterCacheAccess({ cacheHasChanged: false, tokenCache: { serialize: () => 'unchanged' } } as never);
    expect(fs.existsSync(cacheFile())).toBe(false);
    await plugin.afterCacheAccess({ cacheHasChanged: true, tokenCache: { serialize: () => '{"rotated":true}' } } as never);
    expect(fs.readFileSync(cacheFile(), 'utf8')).toBe('{"rotated":true}');
  });

  it('a corrupt cache file starts empty instead of crashing', async () => {
    fs.writeFileSync(cacheFile(), '{not json');
    const m = new AuthManager(['User.Read']);
    await m.initialize();
    expect(await m.listAccounts()).toEqual([]);
  });

  it('8 concurrent writer processes leave a valid cache and no temp/lock files (C-18)', async () => {
    const target = cacheFile();
    const tsx = path.resolve(__dirname, '..', 'node_modules', '.bin', 'tsx');
    const script = path.resolve(__dirname, 'fixtures', 'cache-writer.ts');
    const writers = Array.from({ length: 8 }, (_, i) => new Promise<number>((resolve, reject) => {
      const p = spawn(tsx, [script, target, String(i), '25'], { stdio: 'inherit' });
      p.on('exit', code => resolve(code ?? 1));
      p.on('error', reject);
    }));
    expect(await Promise.all(writers)).toEqual(Array(8).fill(0));
    const final = JSON.parse(fs.readFileSync(target, 'utf8'));
    expect(final.filler.length).toBe(200_000 + Number(final.writer) * 1000);
    expect(fs.readdirSync(dir).filter(f => f !== 'token-cache.json')).toEqual([]);
  }, 60_000);

  it('writeFileAtomic breaks a stale lock', () => {
    const lock = `${cacheFile()}.lock`;
    fs.writeFileSync(lock, '');
    const old = new Date(Date.now() - 60_000);
    fs.utimesSync(lock, old, old);
    writeFileAtomic(cacheFile(), 'ok');
    expect(fs.readFileSync(cacheFile(), 'utf8')).toBe('ok');
    expect(fs.existsSync(lock)).toBe(false);
  });
});

describe('AuthManager behavior', () => {
  it('selected account persists in a file and survives a restart (C-5)', async () => {
    fs.writeFileSync(cacheFile(), cacheJson(1, 2));
    const a = new AuthManager(['User.Read']);
    await a.initialize();
    expect(await a.selectAccount('uid2.utid')).toBe(true);
    expect(JSON.parse(fs.readFileSync(path.join(dir, 'selected-account.json'), 'utf8')).accountId).toBe('uid2.utid');
    const b = new AuthManager(['User.Read']);
    await b.initialize();
    expect((await b.getCurrentAccount())?.homeAccountId).toBe('uid2.utid');
  });

  it('silent acquisition uses .default and passes forceRefresh through (C-6, D26)', async () => {
    fs.writeFileSync(cacheFile(), cacheJson(1));
    const m = new AuthManager(['User.Read']);
    await m.initialize();
    const spy = vi.spyOn(msal(m), 'acquireTokenSilent').mockResolvedValue({ accessToken: 'at', expiresOn: new Date(Date.now() + 3600e3) } as never);
    expect(await m.getToken()).toBe('at');
    expect(await m.getToken(true)).toBe('at');
    expect(spy).toHaveBeenNthCalledWith(1, expect.objectContaining({ scopes: [GRAPH_DEFAULT_SCOPE], forceRefresh: false }));
    expect(spy).toHaveBeenNthCalledWith(2, expect.objectContaining({ forceRefresh: true }));
  });

  it('refreshes tokens within 5 minutes of expiry (C-12)', async () => {
    fs.writeFileSync(cacheFile(), cacheJson(1));
    const m = new AuthManager(['User.Read']);
    await m.initialize();
    const spy = vi.spyOn(msal(m), 'acquireTokenSilent')
      .mockResolvedValueOnce({ accessToken: 'soon', expiresOn: new Date(Date.now() + 120e3) } as never)
      .mockResolvedValueOnce({ accessToken: 'fresh', expiresOn: new Date(Date.now() + 3600e3) } as never);
    expect(await m.getToken()).toBe('soon');
    expect(await m.getToken()).toBe('fresh');
    expect(spy).toHaveBeenCalledTimes(2);
  });

  it('not logged in → CliError exit 3', async () => {
    const m = new AuthManager(['User.Read']);
    await m.initialize();
    const err = await m.getToken().catch(e => e);
    expect(err).toBeInstanceOf(CliError);
    expect((err as CliError).exitCode).toBe(ExitCode.AUTH);
  });

  it('login selects the account that just signed in and removes device-code.txt (C-13, C-14)', async () => {
    fs.writeFileSync(cacheFile(), cacheJson(1));
    fs.writeFileSync(path.join(dir, 'selected-account.json'), JSON.stringify({ accountId: 'uid1.utid' }));
    const m = new AuthManager(['User.Read']);
    await m.initialize();
    vi.spyOn(msal(m), 'acquireTokenByDeviceCode').mockImplementation(async (req: { deviceCodeCallback: (r: { message: string }) => void }) => {
      vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
      req.deviceCodeCallback({ message: 'go to https://microsoft.com/devicelogin' });
      expect(fs.existsSync(path.join(dir, 'device-code.txt'))).toBe(true);
      return { accessToken: 'new', expiresOn: new Date(Date.now() + 3600e3), account: { homeAccountId: 'uid9.utid' } };
    });
    await m.login();
    expect(m.getSelectedAccountId()).toBe('uid9.utid');
    expect(fs.existsSync(path.join(dir, 'device-code.txt'))).toBe(false);
  });

  it('login with no token response throws (C-17)', async () => {
    const m = new AuthManager(['User.Read']);
    vi.spyOn(msal(m), 'acquireTokenByDeviceCode').mockResolvedValue(null as never);
    await expect(m.login()).rejects.toThrow(/did not return an access token/);
  });

  it('logout removes cache, selection, device code, temp and lock files (C-14)', async () => {
    fs.writeFileSync(cacheFile(), cacheJson(1));
    for (const f of ['selected-account.json', 'device-code.txt', 'token-cache.json.123.ab.tmp', 'token-cache.json.lock', 'unrelated.txt']) {
      fs.writeFileSync(path.join(dir, f), 'x');
    }
    const old = new Date(Date.now() - 60_000); // a stale lock left by a crashed process
    fs.utimesSync(path.join(dir, 'token-cache.json.lock'), old, old);
    const m = new AuthManager(['User.Read']);
    await m.initialize();
    await m.logout();
    expect(fs.readdirSync(dir)).toEqual(['unrelated.txt']);
  });

  it('missing client ID → exit 3 with a setup hint', () => {
    vi.stubEnv('MS365_CLI_CLIENT_ID', '');
    vi.stubEnv('MS365_MCP_CLIENT_ID', '');
    try { new AuthManager([]); expect.fail(); } catch (e) {
      expect((e as CliError).exitCode).toBe(ExitCode.AUTH);
      expect((e as CliError).fix).toBe('ms365 auth setup');
    }
  });
});
