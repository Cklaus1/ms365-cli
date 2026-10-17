import { describe, it, expect, afterEach, beforeEach, vi } from 'vitest';
import fs from 'fs';
import path from 'path';
import { AuthManager } from '../src/auth.js';
import { childEnv } from '../src/child-env.js';
import { checkExternalToken, EXPIRY_MARGIN_MS, externalTokenMode } from '../src/external-token.js';
import { CliError, ExitCode, setErrorFormat, toCliError } from '../src/errors.js';
import { GraphClient } from '../src/graph-client.js';
import { registerAuthCommands } from '../src/commands/auth.js';
import { saveProfile } from '../src/profiles.js';
import { runCommand } from './helpers/cli.js';
import { entryHome, fakeJwt, runEntry } from './helpers/entry.js';

const TENANT = '11111111-2222-3333-4444-555555555555';
const OTHER_TENANT = '99999999-2222-3333-4444-555555555555';
const OID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
const now = Date.UTC(2026, 9, 17, 12);
const claims = { aud: 'https://graph.microsoft.com', tid: TENANT, oid: OID, upn: 'ceo@contoso.com', scp: 'Mail.Read User.Read', exp: now / 1000 + 3600 };
const token = (over: Record<string, unknown> = {}) => fakeJwt({ ...claims, ...over });
const host = { host: true, tenantId: TENANT, profileName: 'chief' };

function refusal(fn: () => unknown): CliError {
  try { fn(); } catch (e) { return e as CliError; }
  throw new Error('expected a refusal');
}

afterEach(() => { vi.unstubAllEnvs(); vi.unstubAllGlobals(); });

describe('checkExternalToken (AH-2, AC 6)', () => {
  it('accepts a fresh Graph token for the pinned tenant and account', () => {
    expect(() => checkExternalToken(token(), { ...host, account: 'CEO@contoso.com' }, now)).not.toThrow();
    expect(() => checkExternalToken(token({ aud: '00000003-0000-0000-c000-000000000000' }), { ...host, account: `${OID}.${TENANT}` }, now)).not.toThrow();
  });

  it.each([
    ['expired', token({ exp: now / 1000 - 1 }), 'token_expired'],
    ['inside the 5-minute margin', token({ exp: (now + EXPIRY_MARGIN_MS - 1000) / 1000 }), 'token_expired'],
    ['for another audience', token({ aud: 'https://management.azure.com' }), 'token_audience'],
    ['with no audience', token({ aud: undefined }), 'token_audience'],
    ['for another tenant', token({ tid: OTHER_TENANT }), 'token_identity_mismatch'],
    ['opaque, under a host profile', 'EwBwA8l6BAAU-opaque', 'token_audience'],
  ])('a token %s exits 3 %s', (_what, t, code) => {
    const e = refusal(() => checkExternalToken(t, host, now));
    expect(e.code).toBe(code);
    expect(e.exitCode).toBe(ExitCode.AUTH);
    expect(e.fix).not.toContain('auth login');
  });

  it('another account than the pinned one is token_identity_mismatch (UPN or object ID)', () => {
    expect(refusal(() => checkExternalToken(token(), { ...host, account: 'cfo@contoso.com' }, now)).code).toBe('token_identity_mismatch');
    expect(refusal(() => checkExternalToken(token({ oid: 'x' }), { ...host, account: OID }, now)).code).toBe('token_identity_mismatch');
  });

  it('without a host profile an opaque token is allowed unchecked, and a domain tenant is not compared', () => {
    expect(() => checkExternalToken('EwBwA8l6BAAU-opaque', { host: false }, now)).not.toThrow();
    expect(() => checkExternalToken(token({ tid: OTHER_TENANT }), { host: false, tenantId: 'contoso.com' }, now)).not.toThrow();
  });
});

describe('external AuthManager (AH-2, AC 7)', () => {
  it('needs no client ID, is the default when MS365_CLI_TOKEN is set, and a host profile without a token is token_missing', async () => {
    vi.stubEnv('MS365_CLI_CLIENT_ID', '');
    vi.stubEnv('MS365_CLI_TOKEN', token({ exp: Date.now() / 1000 + 3600 }));
    const m = new AuthManager([]);
    expect(m.tokenSource()).toBe('external');
    await m.initialize();
    expect(await m.getToken()).toContain('.');
    vi.stubEnv('MS365_CLI_TOKEN', '');
    await expect(new AuthManager([], { host: true, profileName: 'chief' }).getToken()).rejects.toMatchObject({ code: 'token_missing', exitCode: 3 });
    expect(externalTokenMode({ type: 'host' })).toBe(true);
    expect(externalTokenMode({})).toBe(false);
  });

  it('a 401 is token_rejected after exactly one request, and the fix never says auth login', async () => {
    vi.stubEnv('MS365_CLI_CLIENT_ID', '');
    vi.stubEnv('MS365_CLI_TOKEN', token({ exp: Date.now() / 1000 + 3600 }));
    const fetches: string[] = [];
    vi.stubGlobal('fetch', async (url: string) => { fetches.push(String(url)); return new Response('{"error":{"code":"InvalidAuthenticationToken","message":"expired"}}', { status: 401 }); });
    const client = new GraphClient(new AuthManager([], { source: 'external' }));
    const e = await client.request('/me').catch((err: unknown) => err as CliError);
    expect(e).toMatchObject({ code: 'token_rejected', exitCode: ExitCode.AUTH, status: 401 });
    expect((e as CliError).fix).not.toContain('auth login');
    expect(fetches).toHaveLength(1);
    expect(toCliError(new Error('Authentication failed')).fix).not.toContain('auth login');
  });
});

describe('auth subcommands in external mode (AC 8)', () => {
  beforeEach(() => setErrorFormat('json'));
  afterEach(() => setErrorFormat('text'));
  const noAuth = () => { throw new Error('no AuthManager in external mode'); };
  const subcommands = [['login'], ['logout'], ['select', 'x'], ['remove', 'x'], ['refresh'], ['grant-scopes'], ['accounts'], ['storage', 'show'], ['storage', 'use', 'file']];

  it.each(subcommands)('auth %s exits 2 external_token_mode with MS365_CLI_TOKEN set (also under --dry-run)', async (...sub) => {
    for (const dryRun of [false, true]) {
      const r = await runCommand((p) => registerAuthCommands(p, noAuth as never), ['auth', ...sub], { env: { MS365_CLI_TOKEN: token() }, dryRun });
      expect(r.exitCode, r.stderr).toBe(2);
      expect(r.stderr).toContain('external_token_mode');
    }
  });

  it('a host profile alone (no token yet) also refuses them', async () => {
    saveProfile('chief', { type: 'host', tenantId: TENANT, scopes: ['Mail.Read'] });
    const r = await runCommand((p) => registerAuthCommands(p, noAuth as never), ['auth', 'login'], { env: { MS365_CLI_PROFILE: 'chief' } });
    expect(r.exitCode).toBe(2);
    expect(r.stderr).toContain('external_token_mode');
  });

  it('no command takes a --token option', () => {
    const src = path.resolve(__dirname, '..', 'src');
    const files = fs.readdirSync(src, { recursive: true }).map(String).filter(f => f.endsWith('.ts'));
    for (const f of files) expect(fs.readFileSync(path.join(src, f), 'utf8'), f).not.toMatch(/['"`]--token\b/);
  });
});

describe('childEnv (AC 9)', () => {
  it('drops every MS365_CLI_TOKEN* variable and keeps the rest', () => {
    expect(childEnv({ MS365_CLI_TOKEN: 't', MS365_CLI_TOKEN_COMMAND: '/bin/x', MS365_CLI_PROFILE: 'p', PATH: '/bin' })).toEqual({ MS365_CLI_PROFILE: 'p', PATH: '/bin' });
  });

  it('guard: every spawn/exec in src/ except the MCP runner (which runs ms365 itself) takes its env from childEnv()', () => {
    const src = path.resolve(__dirname, '..', 'src');
    const files = fs.readdirSync(src, { recursive: true }).map(String).filter(f => f.endsWith('.ts'));
    const sites: string[] = [];
    for (const f of files) {
      const text = fs.readFileSync(path.join(src, f), 'utf8');
      for (const m of text.matchAll(/(?<![.\w])(spawn|spawnSync|exec|execSync|execFile|execFileSync)\(/g)) {
        // The call's text up to its closing parenthesis.
        let depth = 0;
        let end = m.index! + m[0].length - 1;
        for (; end < text.length; end++) {
          if (text[end] === '(') depth++;
          else if (text[end] === ')' && --depth === 0) break;
        }
        const call = text.slice(m.index!, end + 1);
        if (f === 'mcp-server.ts' && call.startsWith('spawn(process.execPath')) continue;
        sites.push(`${f}: ${call.includes('childEnv(') ? 'ok' : call.slice(0, 80)}`);
      }
    }
    expect(sites.length).toBeGreaterThanOrEqual(4);
    expect(sites.filter(s => !s.endsWith(': ok'))).toEqual([]);
  });
});

describe('through the real entry point (AC 5, 6, 31; spawned)', () => {
  const probe = path.resolve(__dirname, 'fixtures', 'msal-probe.mjs');
  const live = () => fakeJwt({ ...claims, exp: Math.floor(Date.now() / 1000) + 3600 });
  const inbox = [{ match: '/messages', body: { value: [{ id: 'm1', subject: 's' }] } }];

  it('AC 5: MS365_CLI_TOKEN, no client ID, empty config dir: mail list succeeds, writes no credential and never loads MSAL', () => {
    const { home, configDir } = entryHome();
    const r = runEntry(['mail', 'list', '-o', 'json'], { home, configDir, env: { MS365_CLI_TOKEN: live() }, routes: inbox, preload: [probe] });
    expect(r.status, r.stderr).toBe(0);
    expect(JSON.parse(r.stdout).value[0].id).toBe('m1');
    // The page, then mail list's own GET /me (to mark mail from the user's own address): both with the supplied token.
    expect(r.fetches[0]).toContain('/me/messages');
    expect(r.stderr).not.toContain('[probe] msal-node loaded');
    expect(fs.readdirSync(configDir).filter(f => f !== 'audit.log')).toEqual([]);
  }, 30_000);

  it('AC 6: a host profile with no token is token_missing; an expired token is token_expired; zero fetches', () => {
    const { home, configDir } = entryHome();
    fs.mkdirSync(path.join(configDir, 'profiles'));
    fs.writeFileSync(path.join(configDir, 'profiles', 'chief.json'), JSON.stringify({ type: 'host', tenantId: TENANT, scopes: ['User.Read', 'Mail.Read'] }));
    const missing = runEntry(['--profile', 'chief', 'mail', 'list'], { home, configDir, routes: inbox });
    expect(missing.status, missing.stderr).toBe(3);
    expect(missing.error?.code).toBe('token_missing');
    const expired = runEntry(['--profile', 'chief', 'mail', 'list'], { home, configDir, routes: inbox, env: { MS365_CLI_TOKEN: fakeJwt({ ...claims, exp: 1 }) } });
    expect(expired.status, expired.stderr).toBe(3);
    expect(expired.error?.code).toBe('token_expired');
    expect([...missing.fetches, ...expired.fetches]).toEqual([]);
  }, 30_000);

  it('AC 31: MS365_CLI_TOKEN from a .env in the working directory does not switch on external mode', () => {
    const { home, configDir } = entryHome();
    fs.writeFileSync(path.join(home, '.env'), `MS365_CLI_TOKEN=${live()}\nMS365_CLI_TOKEN_COMMAND=/bin/true\n`);
    const r = runEntry(['mail', 'list'], { home, configDir, routes: inbox });
    expect(r.status, r.stderr).toBe(3);
    expect(r.error?.code).toBe('no_client_id');
    expect(r.fetches).toEqual([]);
  }, 30_000);
});
