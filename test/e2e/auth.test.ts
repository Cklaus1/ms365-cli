import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { runCli, tempDir, type FakeRoute, type RunResult } from './run-cli.js';
import { aadError, aadRoutes, DEVICE_CODE_MESSAGE, HOME_ACCOUNT_ID } from '../auth-fake-aad.js';

const DEFAULT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const configDir = (home: string) => path.join(home, '.config', 'ms365-cli');
const cacheFile = (home: string) => path.join(configDir(home), 'token-cache.json');
const mode = (p: string) => fs.statSync(p).mode & 0o777;

/** Log in with the built binary against fake AAD; short-lived access token so the next silent call refreshes. */
async function login(env: Record<string, string> = {}): Promise<string> {
  const home = tempDir('ms365-home-');
  const r = await runCli(['auth', 'login'], {
    home,
    env,
    routes: aadRoutes({ clientId: DEFAULT_CLIENT_ID, accessToken: 'at-1', refreshToken: 'rt-1', expiresIn: 60 }),
  });
  expect(r.stderr).not.toContain('Error');
  expect(r.code).toBe(0);
  return home;
}

const refreshRoutes = (accessToken: string, refreshToken: string, me?: FakeRoute) =>
  aadRoutes({ clientId: DEFAULT_CLIENT_ID, accessToken, refreshToken, expiresIn: 60, me });

describe('AUTH-1/2 file token cache (built binary, headless env)', () => {
  it('auth status with no client ID configured reports "Not logged in" (built-in app ID, no crash)', async () => {
    const r = await runCli(['auth', 'status']);
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('Not logged in');
    expect(r.stderr).not.toContain('client ID');
  });

  it('auth login writes the cache at 0600 in a 0700 dir and removes device-code.txt', async () => {
    const home = tempDir('ms365-home-');
    const r = await runCli(['auth', 'login'], {
      home,
      routes: aadRoutes({ clientId: DEFAULT_CLIENT_ID, accessToken: 'at-1', refreshToken: 'rt-1' }),
    });
    expect(r.code).toBe(0);
    expect(r.stderr).toContain(DEVICE_CODE_MESSAGE);
    expect(r.stdout).toContain('Logged in as Ada Lovelace');
    expect(mode(configDir(home))).toBe(0o700);
    expect(mode(cacheFile(home))).toBe(0o600);
    expect(fs.readFileSync(cacheFile(home), 'utf8')).toContain('rt-1');
    const selected = path.join(configDir(home), 'selected-account.json');
    expect(mode(selected)).toBe(0o600);
    expect(JSON.parse(fs.readFileSync(selected, 'utf8')).accountId).toBe(HOME_ACCOUNT_ID);
    expect(fs.existsSync(path.join(configDir(home), 'device-code.txt'))).toBe(false);
    expect(fs.readdirSync(configDir(home)).sort()).toEqual(['selected-account.json', 'token-cache.json']);
  });

  it('auth login never writes device-code.txt, not even transiently', async () => {
    const home = tempDir('ms365-home-');
    fs.mkdirSync(configDir(home), { recursive: true, mode: 0o700 });
    const touched: string[] = [];
    const watcher = fs.watch(configDir(home), (_event, name) => { if (name) touched.push(name); });
    const r = await runCli(['auth', 'login'], {
      home,
      routes: aadRoutes({ clientId: DEFAULT_CLIENT_ID, accessToken: 'at-1', refreshToken: 'rt-1' }),
    }).finally(() => watcher.close());
    expect(r.code).toBe(0);
    expect(r.stderr).toContain(DEVICE_CODE_MESSAGE);
    expect(touched).toContain('token-cache.json');
    expect(touched).not.toContain('device-code.txt');
  });

  it('uses the built-in client ID by default and MS365_CLI_CLIENT_ID when set', async () => {
    const routes = aadRoutes({ clientId: DEFAULT_CLIENT_ID, accessToken: 'at-1', refreshToken: 'rt-1' });
    const byDefault = await runCli(['auth', 'login'], { routes });
    const custom = await runCli(['auth', 'login'], { routes, env: { MS365_CLI_CLIENT_ID: 'custom-client-id' } });
    const clientIdOf = (r: RunResult) =>
      new URLSearchParams(r.requests.find((q) => q.url.includes('/devicecode'))?.body ?? '').get('client_id');
    expect(clientIdOf(byDefault)).toBe(DEFAULT_CLIENT_ID);
    expect(clientIdOf(custom)).toBe('custom-client-id');
  });

  it('forces an existing config dir to 0700 and its files to 0600 on start', async () => {
    const home = await login();
    fs.chmodSync(configDir(home), 0o755);
    fs.chmodSync(cacheFile(home), 0o644);
    fs.chmodSync(path.join(configDir(home), 'selected-account.json'), 0o644);
    await runCli(['auth', 'accounts'], { home });
    expect(mode(configDir(home))).toBe(0o700);
    expect(mode(cacheFile(home))).toBe(0o600);
    expect(mode(path.join(configDir(home), 'selected-account.json'))).toBe(0o600);
  });

  it('persists tokens from a silent refresh and uses the new access token', async () => {
    const home = await login();
    const r = await runCli(['auth', 'status'], { home, routes: refreshRoutes('at-2', 'rt-2') });
    expect(r.code).toBe(0);
    const me = r.requests.find((q) => q.url.endsWith('/v1.0/me'));
    expect(me?.headers.authorization).toBe('Bearer at-2');
    const cache = fs.readFileSync(cacheFile(home), 'utf8');
    expect(cache).toContain('rt-2');
    expect(cache).not.toContain('rt-1');
  });

  it('10 concurrent refreshing processes leave one parseable cache and no temp files', async () => {
    const home = await login();
    const runs = await Promise.all(
      Array.from({ length: 10 }, (_, i) => runCli(['auth', 'status'], { home, routes: refreshRoutes(`at-c${i}`, `rt-c${i}`) })),
    );
    for (const r of runs) expect(r.code, r.stderr).toBe(0);
    const cache = JSON.parse(fs.readFileSync(cacheFile(home), 'utf8'));
    expect(Object.keys(cache.Account)).toHaveLength(1);
    expect(Object.values<{ secret: string }>(cache.RefreshToken).map((t) => t.secret)[0]).toMatch(/^rt-c\d$/);
    expect(fs.readdirSync(configDir(home)).sort()).toEqual(['selected-account.json', 'token-cache.json']);
  });

  it('auth logout removes the cache and selected account', async () => {
    const home = await login();
    const r = await runCli(['auth', 'logout'], { home });
    expect(r.code).toBe(0);
    expect(fs.readdirSync(configDir(home))).toEqual([]);
    expect((await runCli(['auth', 'status'], { home })).stderr).toContain('Not logged in');
  });
});

describe('CORE-5 auth status: Graph failure vs not logged in', () => {
  it('reports a Graph error distinctly from "Not logged in" (exit 1)', async () => {
    const home = await login();
    const r = await runCli(['auth', 'status'], {
      home,
      routes: refreshRoutes('at-2', 'rt-2', { match: '/v1\\.0/me$', status: 503, body: { error: { code: 'ServiceUnavailable', message: 'down' } } }),
    });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('Could not verify');
    expect(r.stderr).toContain('503');
    expect(r.stderr).not.toContain('Not logged in');
  });

  it('asks for a new login when the refresh needs interaction (permissions changed)', async () => {
    const home = await login();
    const r = await runCli(['auth', 'status'], {
      home,
      routes: [{ match: '/oauth2/v2\\.0/token', method: 'POST', status: 400, body: aadError(65001, 'invalid_grant', 'consent_required') }],
    });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('Permissions changed');
    expect(r.stderr).toContain('ms365 auth login');
  });
});

describe('AUTH-4 built-in app errors', () => {
  const blocked: FakeRoute[] = [
    aadRoutes({ clientId: DEFAULT_CLIENT_ID, accessToken: 'x', refreshToken: 'x' })[0],
    { match: '/oauth2/v2\\.0/token', method: 'POST', status: 400, body: aadError(53003, 'invalid_grant') },
  ];

  it('points to the custom-app setup when the tenant blocks the built-in app', async () => {
    const r = await runCli(['auth', 'login'], { routes: blocked });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('AADSTS53003');
    expect(r.stderr).toContain('ms365 auth setup');
  });

  it('does not suggest the custom app when one is already configured', async () => {
    const r = await runCli(['auth', 'login'], { routes: blocked, env: { MS365_CLI_CLIENT_ID: 'custom-client-id' } });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('AADSTS53003');
    expect(r.stderr).not.toContain('ms365 auth setup');
  });
});
