import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { PassThrough } from 'stream';
import { InteractionRequiredAuthError, ServerError } from '@azure/msal-node';
import { AuthManager } from '../src/auth.js';
import { azDeps } from '../src/grant-scopes.js';
import { CliError, ExitCode, toCliError } from '../src/errors.js';
import { serve, type Tool, type Runner } from '../src/mcp-server.js';

// IDT-5: long-running processes (mcp serve, watch --follow) hold one AuthManager, which reads the credential store
// once (ID-12). A silent refresh that fails with invalid_grant / interaction_required re-reads the store once and
// retries (another process may have refreshed or re-logged in) before the normal re-login error. Admin (az) profiles
// re-borrow from az when the memory token nears expiry; an az failure mid-run is exit 3 az_required (ID-8).
// Fake MSAL (spied acquireTokenSilent) and stubbed az only: no keychain, no network.

const CID = '11111111-1111-1111-1111-111111111111';
const HID = 'uid1.utid';

function cacheJson(rtSecret: string): string {
  return JSON.stringify({
    Account: {
      [`${HID}-login.microsoftonline.com-utid`]: {
        home_account_id: HID, environment: 'login.microsoftonline.com', realm: 'utid',
        local_account_id: 'uid1', username: 'user1@contoso.com', authority_type: 'MSSTS',
      },
    },
    IdToken: {}, AccessToken: {},
    RefreshToken: {
      [`${HID}-login.microsoftonline.com-refreshtoken-${CID}--`]: {
        home_account_id: HID, environment: 'login.microsoftonline.com', credential_type: 'RefreshToken',
        client_id: CID, secret: rtSecret,
      },
    },
    AppMetadata: {},
  });
}

function jwt(payload: Record<string, unknown>): string {
  const b64 = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${b64({ alg: 'RS256', typ: 'JWT' })}.${b64(payload)}.sig`;
}

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-longrun-'));
  vi.stubEnv('MS365_CLI_CONFIG_DIR', dir);
  vi.stubEnv('MS365_CLI_TOKEN', '');
  vi.stubEnv('MS365_CLI_CLIENT_ID', CID);
});
afterEach(() => { vi.useRealTimers(); vi.unstubAllEnvs(); vi.restoreAllMocks(); });
const cacheFile = () => path.join(dir, 'token-cache.json');

type FakeApp = { acquireTokenSilent: (...a: unknown[]) => unknown; getTokenCache(): { getAllAccounts(): Promise<unknown[]>; serialize(): string } };
const msal = (m: AuthManager) => (m as unknown as { app(): Promise<FakeApp> }).app();
const expired = () => new InteractionRequiredAuthError('invalid_grant', 'AADSTS70043: The refresh token has expired or is invalid due to sign-in frequency checks.');
const ok = (accessToken: string) => ({ accessToken, expiresOn: new Date(Date.now() + 3600e3) });

/**
 * A fake acquireTokenSilent that behaves like MSAL's: it reads the cache through the plugin (beforeCacheAccess)
 * and succeeds only when the refresh token it sees is `good`; otherwise invalid_grant.
 */
async function fakeSilent(m: AuthManager, good: string) {
  const app = await msal(m);
  return vi.spyOn(app, 'acquireTokenSilent').mockImplementation(async () => {
    const cache = app.getTokenCache();
    await cache.getAllAccounts();
    if (cache.serialize().includes(good)) return ok(`at-${good}`) as never;
    throw expired();
  });
}

describe('IDT-5 / ID-12: re-read the store once on invalid_grant / interaction_required', () => {
  it('another process re-logged in after this one read the store → one re-read, silent retry succeeds', async () => {
    fs.writeFileSync(cacheFile(), cacheJson('rt-old'), { mode: 0o600 });
    const m = new AuthManager(['User.Read']);
    await m.initialize();
    const spy = await fakeSilent(m, 'rt-new');
    // This process has read (and cached) the store: rt-old.
    expect((await m.listAccounts()).map(a => a.homeAccountId)).toEqual([HID]);
    // Another process signs in again and rewrites the store.
    fs.writeFileSync(cacheFile(), cacheJson('rt-new'), { mode: 0o600 });

    expect(await m.getToken()).toBe('at-rt-new');
    expect(spy).toHaveBeenCalledTimes(2);
  });

  it('still invalid after the re-read → retried exactly once, then the normal re-login error (exit 3)', async () => {
    fs.writeFileSync(cacheFile(), cacheJson('rt-old'), { mode: 0o600 });
    const m = new AuthManager(['User.Read']);
    await m.initialize();
    const spy = await fakeSilent(m, 'never');
    const err = await m.getToken().catch(e => e);
    expect(spy).toHaveBeenCalledTimes(2);
    expect(toCliError(err)).toMatchObject({ code: 'auth_required', exitCode: ExitCode.AUTH, fix: 'ms365 auth login' });
  });

  it('a server invalid_grant (ServerError) is retried the same way', async () => {
    fs.writeFileSync(cacheFile(), cacheJson('rt-old'), { mode: 0o600 });
    const m = new AuthManager(['User.Read']);
    await m.initialize();
    const app = await msal(m);
    const spy = vi.spyOn(app, 'acquireTokenSilent')
      .mockRejectedValueOnce(new ServerError('invalid_grant', 'AADSTS50173: The provided grant has expired.'))
      .mockResolvedValueOnce(ok('second') as never);
    expect(await m.getToken()).toBe('second');
    expect(spy).toHaveBeenCalledTimes(2);
  });

  it('other failures (network, throttling) are not retried here', async () => {
    fs.writeFileSync(cacheFile(), cacheJson('rt-old'), { mode: 0o600 });
    const m = new AuthManager(['User.Read']);
    await m.initialize();
    const spy = vi.spyOn(await msal(m), 'acquireTokenSilent').mockRejectedValue(new Error('network down'));
    await expect(m.getToken()).rejects.toThrow('network down');
    expect(spy).toHaveBeenCalledTimes(1);
  });

  it('a retried refresh that succeeds is persisted, and later calls use the in-memory token', async () => {
    fs.writeFileSync(cacheFile(), cacheJson('rt-old'), { mode: 0o600 });
    const m = new AuthManager(['User.Read']);
    await m.initialize();
    const spy = await fakeSilent(m, 'rt-new');
    await m.listAccounts();
    fs.writeFileSync(cacheFile(), cacheJson('rt-new'), { mode: 0o600 });
    await m.getToken();
    await m.getToken();
    expect(spy).toHaveBeenCalledTimes(2);
    expect(fs.readFileSync(cacheFile(), 'utf8')).toContain('rt-new');
  });
});

describe('IDT-5 / ID-8: admin profile over a long run', () => {
  const TENANT = 'contoso.onmicrosoft.com';
  const azOk = (expMs: number, tag: string) => ({
    status: 0, stdout: JSON.stringify({ accessToken: jwt({ scp: 'User.Read.All', exp: Math.floor(expMs / 1000), tag }), tenant: 'tid-1' }), stderr: '', missing: false,
  });
  const azLoggedOut = { status: 1, stdout: '', stderr: "ERROR: Please run 'az login' to setup account.", missing: false };

  it('re-borrows from az before expiry; an az failure mid-run is exit 3 az_required; az recovering heals it', async () => {
    vi.useFakeTimers({ now: new Date('2026-10-15T10:00:00Z'), toFake: ['Date'] });
    const results: Array<() => ReturnType<typeof azDeps.run>> = [];
    const run = vi.spyOn(azDeps, 'run').mockImplementation(() => results.shift()!());
    const m = new AuthManager([], { source: 'az', tenantId: TENANT, profileName: 'adm' });
    await m.initialize();

    results.push(() => azOk(Date.now() + 60 * 60e3, 'one'));
    const t1 = await m.getToken();
    // 50 minutes in: still > 5 min left → memory token, no az call.
    vi.setSystemTime(Date.now() + 50 * 60e3);
    expect(await m.getToken()).toBe(t1);
    expect(run).toHaveBeenCalledTimes(1);

    // 56 minutes in: within 5 min of exp → borrowed again.
    vi.setSystemTime(Date.now() + 6 * 60e3);
    results.push(() => azOk(Date.now() + 60 * 60e3, 'two'));
    const t2 = await m.getToken();
    expect(t2).not.toBe(t1);
    expect(run).toHaveBeenCalledTimes(2);

    // Over an hour later az's session is gone: exit 3 az_required, not a crash.
    vi.setSystemTime(Date.now() + 61 * 60e3);
    results.push(() => azLoggedOut);
    const err = await m.getToken().catch(e => e);
    expect(err).toBeInstanceOf(CliError);
    expect(err).toMatchObject({ code: 'az_required', exitCode: ExitCode.AUTH, fix: `az login --tenant ${TENANT}` });
    // The expired token is never handed out after a failed borrow.
    results.push(() => azLoggedOut);
    await expect(m.getToken()).rejects.toMatchObject({ code: 'az_required' });

    // The admin runs az login; the same manager recovers without a restart.
    results.push(() => azOk(Date.now() + 60 * 60e3, 'three'));
    const t3 = await m.getToken();
    expect(t3).not.toBe(t2);
    expect(run).toHaveBeenCalledTimes(5);
    expect(fs.readdirSync(dir)).toEqual([]);
  });

  it('mcp serve: an az failure in a tool call is an isError result with az_required; the server keeps serving', async () => {
    let azUp = true;
    vi.spyOn(azDeps, 'run').mockImplementation(() => (azUp ? azOk(Date.now() + 2 * 60e3, 'x') : azLoggedOut));
    // A long-lived holder: one manager for the whole session (token within 5 min of exp → az on every call).
    const m = new AuthManager([], { source: 'az', tenantId: TENANT, profileName: 'adm' });
    const runner: Runner = async () => {
      try {
        await m.getToken();
        return { exitCode: 0, stdout: '{"ok":true}\n', stderr: '' };
      } catch (err) {
        const e = toCliError(err);
        return { exitCode: e.exitCode, stdout: '', stderr: JSON.stringify({ error: { code: e.code, exitCode: e.exitCode, message: e.message, fix: e.fix } }) + '\n' };
      }
    };
    const tool: Tool = {
      name: 'user_get', path: 'user get', description: 'Get a user',
      inputSchema: { type: 'object', properties: {}, additionalProperties: false },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true }, params: [], hasOutput: true,
    };
    const out: Array<Record<string, any>> = [];
    const input = new PassThrough();
    let answered: () => void = () => {};
    const next = () => new Promise<void>(r => { answered = r; });
    const done = serve(input, (line) => { out.push(JSON.parse(line)); answered(); }, { tools: [tool], version: '0.1.0', runner });
    const call = async (id: number) => {
      const p = next();
      input.write(JSON.stringify({ jsonrpc: '2.0', id, method: 'tools/call', params: { name: 'user_get', arguments: {} } }) + '\n');
      await p;
    };
    await call(1);
    azUp = false;
    await call(2);
    azUp = true;
    await call(3);
    input.end();
    await done;
    const byId = Object.fromEntries(out.map(r => [r.id, r]));
    expect(byId[1].result.isError).toBe(false);
    expect(byId[2].error).toBeUndefined();
    expect(byId[2].result.isError).toBe(true);
    expect(JSON.parse(byId[2].result.content[0].text).error).toMatchObject({ code: 'az_required', exitCode: 3 });
    expect(byId[3].result.isError).toBe(false);
  });
});
