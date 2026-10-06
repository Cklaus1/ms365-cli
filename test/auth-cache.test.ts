import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { AuthManager, NotLoggedInError, writeSecureFile } from '../src/auth.js';
import { DEVICE_CODE_RESPONSE, ME, tokenResponse } from './auth-fake-aad.js';

const CLIENT_ID = '33333333-3333-3333-3333-333333333333';
const SAVED_ENV = ['HOME', 'MS365_CLI_CLIENT_ID', 'MS365_CLI_TOKEN', 'MS365_MCP_OAUTH_TOKEN'] as const;

/** Fake identity platform + Graph over global fetch; each token request takes the next queued response. */
function stubAad(tokenQueue: Array<Record<string, unknown>>) {
  const tokenRequests: URLSearchParams[] = [];
  vi.stubGlobal('fetch', async (input: string | URL | Request, init: RequestInit = {}) => {
    const url = String(input instanceof Request ? input.url : input);
    const json = (body: unknown) => new Response(JSON.stringify(body), { headers: { 'content-type': 'application/json' } });
    if (url.includes('/devicecode')) return json(DEVICE_CODE_RESPONSE);
    if (url.includes('/oauth2/v2.0/token')) {
      tokenRequests.push(new URLSearchParams(String(init.body)));
      const next = tokenQueue.shift();
      if (!next) throw new Error('unexpected token request');
      return json(next);
    }
    if (url.endsWith('/v1.0/me')) return json(ME);
    return new Response('{}', { status: 404 });
  });
  return tokenRequests;
}

describe('AuthManager with the file token cache', () => {
  let saved: Record<string, string | undefined>;
  let home: string;

  beforeEach(() => {
    saved = Object.fromEntries(SAVED_ENV.map((k) => [k, process.env[k]]));
    for (const k of SAVED_ENV) delete process.env[k];
    home = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-home-'));
    process.env.HOME = home;
    vi.spyOn(process.stderr, 'write').mockReturnValue(true); // device-code message
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    for (const k of SAVED_ENV) {
      if (saved[k] === undefined) delete process.env[k];
      else process.env[k] = saved[k];
    }
  });

  it('getToken(true) skips the cached access token and refreshes (GRAPH-7 auth side)', async () => {
    const tokenRequests = stubAad([
      tokenResponse({ clientId: CLIENT_ID, accessToken: 'at-1', refreshToken: 'rt-1', scope: 'User.Read' }),
      tokenResponse({ clientId: CLIENT_ID, accessToken: 'at-2', refreshToken: 'rt-2', scope: 'User.Read' }),
    ]);
    const auth = new AuthManager(['User.Read'], { clientId: CLIENT_ID });
    await auth.initialize();
    await auth.login();

    // A fresh process (new AuthManager) reads the cached, still-valid token without a network call.
    const next = new AuthManager(['User.Read'], { clientId: CLIENT_ID });
    await next.initialize();
    expect(await next.getToken()).toBe('at-1');
    expect(tokenRequests).toHaveLength(1);

    expect(await next.getToken(true)).toBe('at-2');
    expect(tokenRequests[1].get('grant_type')).toBe('refresh_token');
    expect(tokenRequests[1].get('refresh_token')).toBe('rt-1');
    expect(fs.readFileSync(path.join(home, '.config/ms365-cli/token-cache.json'), 'utf8')).toContain('rt-2');
  });

  it('treats a corrupt cache file as logged out, and the next login replaces it', async () => {
    const dir = path.join(home, '.config', 'ms365-cli');
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, 'token-cache.json'), '{"Account": {', { mode: 0o600 });
    stubAad([tokenResponse({ clientId: CLIENT_ID, accessToken: 'at-1', refreshToken: 'rt-1' })]);

    const auth = new AuthManager(['User.Read'], { clientId: CLIENT_ID });
    await auth.initialize();
    await expect(auth.whoAmI()).rejects.toBeInstanceOf(NotLoggedInError);

    await auth.login();
    expect(JSON.parse(fs.readFileSync(path.join(dir, 'token-cache.json'), 'utf8')).RefreshToken).toBeTruthy();
    expect((await auth.whoAmI()).displayName).toBe(ME.displayName);
  });
});

describe('writeSecureFile', () => {
  it('concurrent writers leave exactly one complete version and no temp files', async () => {
    const dir = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-wsf-')), 'cfg');
    const file = path.join(dir, 'token-cache.json');
    const payloads = Array.from({ length: 10 }, (_, i) => JSON.stringify({ writer: i, pad: 'x'.repeat(200_000) }));
    await Promise.all(payloads.map((p) => writeSecureFile(file, p)));
    expect(payloads).toContain(fs.readFileSync(file, 'utf8'));
    expect(fs.readdirSync(dir)).toEqual(['token-cache.json']);
    expect(fs.statSync(file).mode & 0o777).toBe(0o600);
    expect(fs.statSync(dir).mode & 0o777).toBe(0o700);
  });
});
