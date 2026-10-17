import { describe, it, expect, afterEach, vi } from 'vitest';
import fs from 'fs';
import path from 'path';
import type { Command } from 'commander';
import { AuthManager, configDir } from '../src/auth.js';
import { registerAuthCommands } from '../src/commands/auth.js';
import { healthOf, registerHealthCommand } from '../src/commands/health.js';
import { registerUserCommands } from '../src/commands/user.js';
import { CliError, ExitCode, setErrorFormat } from '../src/errors.js';
import { GraphClient, requestBudget } from '../src/graph-client.js';
import { saveProfile } from '../src/profiles.js';
import { jsonResponse, runCommand, type Route } from './helpers/cli.js';
import { entryHome, fakeJwt, runEntry } from './helpers/entry.js';

const TENANT = '11111111-2222-3333-4444-555555555555';
const OID = '22222222-3333-4444-5555-666666666666';
const claims = { aud: 'https://graph.microsoft.com', tid: TENANT, oid: OID, upn: 'ceo@contoso.com', scp: 'Mail.Read User.Read' };
const live = () => fakeJwt({ ...claims, exp: Math.floor(Date.now() / 1000) + 3600 });

const health = (program: Command, getClient: () => GraphClient) => registerHealthCommand(program, getClient);
const user = (program: Command, getClient: () => GraphClient) => registerUserCommands(program, getClient);
const lastLine = (s: string) => s.trim().split('\n').filter(l => l.trim()).pop() ?? '';
const errorOf = (stderr: string) => JSON.parse(lastLine(stderr)).error;
const err = (status: number, extra: Record<string, string> = {}) => () =>
  new Response(JSON.stringify({ error: { code: 'x', message: 'boom' } }), { status, headers: { 'content-type': 'application/json', ...extra } });
/** A route that never answers; it fails when the client aborts the request. */
const hang = (_c: unknown, signal?: AbortSignal) => {
  const { promise, reject } = Promise.withResolvers<never>();
  signal?.addEventListener('abort', () => reject(signal.reason ?? new Error('aborted')));
  return promise;
};
const me: Route = ['GET', '/me?', () => ({ id: OID })];

async function healthRun(argv: string[], routes: Route[], env: Record<string, string> = {}) {
  const r = await runCommand(health, ['health', ...argv, '-o', 'json'], { routes, env });
  expect(r.exitCode === -1 || r.exitCode === 0, r.stderr).toBe(true);
  return { ...r, records: JSON.parse(r.stdout) as Array<Record<string, unknown>> };
}

afterEach(() => {
  setErrorFormat('text');
  vi.restoreAllMocks();
  vi.unstubAllEnvs();
});

describe('AH-8(3) mapping table (AC 16): every row', () => {
  it.each([
    ['exit 3', new CliError('x', { code: 'token_expired', exitCode: ExitCode.AUTH }), 'AUTH_REQUIRED', 'token_expired'],
    ['scope_ceiling', new CliError('x', { code: 'scope_ceiling', exitCode: ExitCode.PERMISSION }), 'UNSUPPORTED', 'scope_ceiling'],
    ['missing_scope', new CliError('x', { code: 'missing_scope', exitCode: ExitCode.PERMISSION }), 'AUTH_REQUIRED', 'missing_scope'],
    ['403', new CliError('x', { code: 'permission_denied', exitCode: ExitCode.PERMISSION, status: 403 }), 'AUTH_REQUIRED', 'access_denied'],
    ['429', new CliError('x', { code: 'throttled', exitCode: ExitCode.THROTTLED, status: 429, retryAfterMs: 5000 }), 'RATE_LIMITED', 'throttled'],
    ['complete:false', { value: [], complete: false, incompleteReason: 'page_limit' }, 'PARTIAL', 'page_limit'],
    ['gap', { type: 'gap', reason: 'sync_state_expired', from: 'a', to: 'b' }, 'PARTIAL', 'sync_state_expired'],
    ['timeout', new CliError('x', { code: 'timeout' }), 'STALE', 'timeout'],
    ['graph_error 503', new CliError('x', { code: 'graph_error', status: 503 }), 'STALE', 'graph_error'],
    ['network_error', new CliError('x', { code: 'network_error' }), 'STALE', 'network_error'],
    ['bad_response', new CliError('x', { code: 'bad_response' }), 'STALE', 'bad_response'],
    ['generic error', new Error('boom'), 'STALE', 'error'],
    ['delta_expired', new CliError('x', { code: 'delta_expired' }), 'PARTIAL', 'delta_expired'],
    ['exit 2', new CliError('x', { code: 'usage', exitCode: ExitCode.USAGE }), 'UNSUPPORTED', 'usage'],
    ['cursor_mismatch', new CliError('x', { code: 'cursor_mismatch', exitCode: ExitCode.USAGE }), 'UNSUPPORTED', 'usage'],
    ['checkpoint_mismatch', new CliError('x', { code: 'checkpoint_mismatch', exitCode: ExitCode.USAGE }), 'UNSUPPORTED', 'usage'],
    ['a complete read', { value: [], complete: true }, 'HEALTHY', null],
  ])('%s', (_row, signal, status, reason) => {
    expect(healthOf(signal)).toEqual({ status, reason });
  });
});

describe('ms365 health (AH-8(2), AC 16)', () => {
  it('probes GET /me then one $top=1 read per resource, and reports HEALTHY with the connector-health fields', async () => {
    const r = await healthRun(['mail', 'events'], [me, ['GET', '/me/messages', () => ({ value: [] })], ['GET', '/me/events', () => ({ value: [] })]]);
    expect(r.calls.map(c => c.path)).toEqual(['/me?$select=id', '/me/messages?$top=1&$select=id', '/me/events?$top=1&$select=id']);
    expect(r.records.map(x => Object.keys(x).sort())).toEqual(Array(2).fill(['checkedAt', 'checkpoint', 'coverageEnd', 'coverageStart', 'lastSuccessAt', 'reason', 'resource', 'status']));
    expect(r.records).toMatchObject([{ resource: 'mail', status: 'HEALTHY', reason: null }, { resource: 'events', status: 'HEALTHY', reason: null }]);
    expect(Number.isNaN(Date.parse(String(r.records[0].checkedAt)))).toBe(false);
  });

  // 401 is sent once more with a fresh token (the client's re-auth, not a retry); every other answer is final.
  it.each([
    ['401 (exit 3)', 'AUTH_REQUIRED', 'auth_failed', 2, err(401)],
    ['403', 'AUTH_REQUIRED', 'access_denied', 1, err(403)],
    ['429', 'RATE_LIMITED', 'throttled', 1, err(429, { 'Retry-After': '3' })],
    ['503', 'STALE', 'graph_error', 1, err(503)],
    ['a network error', 'STALE', 'network_error', 1, () => { throw new TypeError('fetch failed'); }],
  ])('a probe answered with %s is %s %s, with no retry', async (_label, status, reason, attempts, answer) => {
    const r = await healthRun(['mail'], [me, ['GET', '/me/messages', answer]]);
    expect(r.records).toMatchObject([{ resource: 'mail', status, reason }]);
    expect(r.calls.filter(c => c.path.startsWith('/me/messages'))).toHaveLength(attempts);
  });

  it('a failing GET /me is every resource\'s state; the resources are not probed', async () => {
    const r = await healthRun([], [['GET', '/me?', err(503)]]);
    expect(r.records).toMatchObject([{ resource: 'mail', status: 'STALE' }, { resource: 'events', status: 'STALE' }]);
    expect(r.calls).toHaveLength(1);
  });

  it('probes run with 0 retries and a 10 s timeout, lower if MS365_CLI_REQUEST_TIMEOUT_MS is lower', async () => {
    const spy = vi.spyOn(GraphClient.prototype, 'request');
    await healthRun(['mail'], [me]);
    expect(spy.mock.calls.map(c => c[1])).toEqual([expect.objectContaining({ retries: 0, timeoutMs: 10_000 }), expect.objectContaining({ retries: 0, timeoutMs: 10_000 })]);
    spy.mockClear();
    await healthRun(['mail'], [me], { MS365_CLI_REQUEST_TIMEOUT_MS: '2000' });
    expect(spy.mock.calls.map(c => c[1])).toEqual([expect.objectContaining({ timeoutMs: 2000 }), expect.objectContaining({ timeoutMs: 2000 })]);
  });

  it('a probe that never answers is STALE timeout after one attempt, within its timeout', async () => {
    const started = Date.now();
    const r = await healthRun(['mail'], [me, ['GET', '/me/messages', hang]], { MS365_CLI_REQUEST_TIMEOUT_MS: '1000' });
    expect(Date.now() - started).toBeLessThan(3000);
    expect(r.records).toMatchObject([{ resource: 'mail', status: 'STALE', reason: 'timeout' }]);
    expect(r.calls.filter(c => c.path.startsWith('/me/messages'))).toHaveLength(1);
  });

  it('coverage comes from the saved watch cursor, or from --checkpoint for its resource', async () => {
    const dir = path.join(configDir(), 'watch');
    fs.mkdirSync(dir, { recursive: true });
    const cursor = { resource: 'mail', endpoint: '/me/mailFolders/inbox/messages/delta', deltaLink: 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=t', syncedAt: '2026-10-17T10:00:00.000Z', coverageStart: '2026-10-01T00:00:00.000Z' };
    fs.writeFileSync(path.join(dir, 'mail.json'), JSON.stringify(cursor));
    try {
      let r = await healthRun(['mail', 'events'], [me]);
      expect(r.records).toMatchObject([
        { resource: 'mail', lastSuccessAt: cursor.syncedAt, coverageStart: cursor.coverageStart, coverageEnd: cursor.syncedAt, checkpoint: null },
        { resource: 'events', lastSuccessAt: null, coverageStart: null, coverageEnd: null, checkpoint: null },
      ]);
      const cp = { v: 1, resource: 'events', endpoint: '/me/calendarView/delta', deltaLink: 'https://graph.microsoft.com/v1.0/me/calendarView/delta?$deltatoken=e', syncedAt: '2026-10-17T11:00:00.000Z', coverageStart: '2026-10-02T00:00:00.000Z' };
      const blob = Buffer.from(JSON.stringify(cp)).toString('base64url');
      r = await healthRun(['mail', 'events', '--checkpoint', blob], [me]);
      expect(r.records[1]).toMatchObject({ resource: 'events', lastSuccessAt: cp.syncedAt, coverageStart: cp.coverageStart, coverageEnd: cp.syncedAt, checkpoint: blob });
      expect(r.records[0]).toMatchObject({ checkpoint: null, coverageStart: cursor.coverageStart });
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('a --checkpoint for a resource not checked, or an unknown resource, exits 2 before any request', async () => {
    setErrorFormat('json');
    const blob = Buffer.from(JSON.stringify({ v: 1, resource: 'events', endpoint: 'e', deltaLink: 'd', syncedAt: 's', coverageStart: 'c' })).toString('base64url');
    let r = await runCommand(health, ['health', 'mail', '--checkpoint', blob]);
    expect(r.exitCode).toBe(2);
    expect(errorOf(r.stderr).code).toBe('checkpoint_mismatch');
    r = await runCommand(health, ['health', 'files']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toEqual([]);
  });

  it('a resource outside the profile\'s ceiling is UNSUPPORTED scope_ceiling and is not probed (E12)', async () => {
    saveProfile('chief', { type: 'host', tenantId: TENANT, scopes: ['User.Read', 'Mail.Read'] });
    const r = await healthRun(['mail', 'events'], [me], { MS365_CLI_PROFILE: 'chief' });
    expect(r.records).toMatchObject([{ resource: 'mail', status: 'HEALTHY' }, { resource: 'events', status: 'UNSUPPORTED', reason: 'scope_ceiling' }]);
    expect(r.calls.some(c => c.path.startsWith('/me/events'))).toBe(false);
  });
});

describe('AH-8(4),(6) request budget and the error line (AC 29)', () => {
  it('MS365_CLI_MAX_RETRIES=0: a 503 makes exactly one fetch and exits 1 graph_error with status 503', async () => {
    setErrorFormat('json');
    const r = await runCommand(user, ['user', 'me'], { routes: [['GET', '/me', err(503)]], env: { MS365_CLI_MAX_RETRIES: '0' } });
    expect(r.exitCode).toBe(1);
    expect(r.calls).toHaveLength(1);
    expect(errorOf(r.stderr)).toMatchObject({ code: 'graph_error', status: 503 });
  });

  it('MS365_CLI_REQUEST_TIMEOUT_MS=1000 against a mock that never answers exits 1 timeout after one attempt within ~1 s', async () => {
    setErrorFormat('json');
    const started = Date.now();
    const r = await runCommand(user, ['user', 'me'], { routes: [['GET', '/me', hang]], env: { MS365_CLI_REQUEST_TIMEOUT_MS: '1000', MS365_CLI_MAX_RETRIES: '0' } });
    expect(Date.now() - started).toBeLessThan(2500);
    expect(r.exitCode).toBe(1);
    expect(r.calls).toHaveLength(1);
    expect(errorOf(r.stderr).code).toBe('timeout');
  });

  it('with retries, the error JSON is the last non-empty stderr line and the retry notices come first', async () => {
    setErrorFormat('json');
    const r = await runCommand(user, ['user', 'me'], { routes: [['GET', '/me', err(503)]], env: { MS365_CLI_MAX_RETRIES: '1' } });
    expect(r.calls).toHaveLength(2);
    const lines = r.stderr.split('\n').filter(l => l.trim());
    expect(lines.slice(0, -1).some(l => /Retrying in/.test(l))).toBe(true);
    expect(errorOf(r.stderr)).toMatchObject({ code: 'graph_error', status: 503, exitCode: 1 });
  });

  it('a 429 error carries retryAfterMs from Retry-After and exits 6', async () => {
    setErrorFormat('json');
    const r = await runCommand(user, ['user', 'me'], { routes: [['GET', '/me', err(429, { 'Retry-After': '7' })]], env: { MS365_CLI_MAX_RETRIES: '0' } });
    expect(r.exitCode).toBe(ExitCode.THROTTLED);
    expect(errorOf(r.stderr)).toMatchObject({ code: 'throttled', status: 429, retryAfterMs: 7000 });
  });

  it('a non-idempotent request is still not retried on 5xx', async () => {
    const r = await runCommand(user, ['user', 'me'], { routes: [['GET', '/me', () => jsonResponse({ id: OID })]] });
    expect(r.exitCode, r.stderr).not.toBe(1);
    const client = new GraphClient(() => ({ getToken: async () => 't' }) as never);
    const origFetch = globalThis.fetch;
    let fetches = 0;
    globalThis.fetch = (async () => { fetches++; return err(503)(); }) as typeof fetch;
    try {
      await expect(client.request('/me/sendMail', { method: 'POST', body: '{}' })).rejects.toMatchObject({ status: 503 });
    } finally {
      globalThis.fetch = origFetch;
    }
    expect(fetches).toBe(1);
  });

  it.each([
    [{}, { retries: 2, timeoutMs: 30_000 }],
    [{ MS365_CLI_MAX_RETRIES: '0', MS365_CLI_REQUEST_TIMEOUT_MS: '1000' }, { retries: 0, timeoutMs: 1000 }],
    [{ MS365_CLI_MAX_RETRIES: '2', MS365_CLI_REQUEST_TIMEOUT_MS: '30000' }, { retries: 2, timeoutMs: 30_000 }],
  ])('requestBudget(%o)', (env, budget) => {
    expect(requestBudget(env)).toEqual(budget);
  });

  it.each([
    { MS365_CLI_MAX_RETRIES: '3' }, { MS365_CLI_MAX_RETRIES: '-1' }, { MS365_CLI_MAX_RETRIES: '1.5' }, { MS365_CLI_MAX_RETRIES: 'two' },
    { MS365_CLI_REQUEST_TIMEOUT_MS: '999' }, { MS365_CLI_REQUEST_TIMEOUT_MS: '30001' },
  ])('out of range %o is a usage error (exit 2)', (env) => {
    expect(() => requestBudget(env)).toThrow(expect.objectContaining({ code: 'usage', exitCode: ExitCode.USAGE }));
  });
});

describe('auth status -o json (AH-8(1), AC 15)', () => {
  const register = (m: AuthManager) => (program: Command) => registerAuthCommands(program, () => m);

  it('msal (mocked): source, account, expiry, scopes; no ceiling without a profile', async () => {
    vi.stubEnv('MS365_CLI_TOKEN', '');
    vi.stubEnv('MS365_CLI_CLIENT_ID', '11111111-1111-1111-1111-111111111111');
    const m = new AuthManager([]);
    const exp = Math.floor(Date.now() / 1000) + 3600;
    vi.spyOn(m, 'getToken').mockResolvedValue(fakeJwt({ ...claims, exp }));
    const r = await runCommand(register(m), ['auth', 'status', '-o', 'json']);
    expect(r.exitCode === -1 || r.exitCode === 0, r.stderr).toBe(true);
    expect(JSON.parse(r.stdout)).toEqual({
      source: 'msal', profile: null, type: null, tenantId: TENANT, accountId: OID, upn: 'ceo@contoso.com',
      expiresAt: new Date(exp * 1000).toISOString(), scopes: ['Mail.Read', 'User.Read'], ceiling: null, ok: true,
    });
  });

  it('msal with no usable token still exits 3', async () => {
    vi.stubEnv('MS365_CLI_TOKEN', '');
    vi.stubEnv('MS365_CLI_CLIENT_ID', '11111111-1111-1111-1111-111111111111');
    const r = await runCommand(register(new AuthManager([])), ['auth', 'status', '-o', 'json']);
    expect(r.exitCode).toBe(ExitCode.AUTH);
    expect(r.stdout).toBe('');
  });
});

describe('through the real entry point (AC 15, 29, 32; spawned)', () => {
  const probe = path.resolve(__dirname, 'fixtures', 'msal-probe.mjs');

  it('AC 15: a host profile\'s auth status -o json is external with its ceiling, and MSAL is never loaded', () => {
    const { home, configDir: cfg } = entryHome();
    fs.mkdirSync(path.join(cfg, 'profiles'));
    fs.writeFileSync(path.join(cfg, 'profiles', 'chief.json'), JSON.stringify({ type: 'host', tenantId: TENANT, scopes: ['User.Read', 'Mail.Read'] }));
    const r = runEntry(['--profile', 'chief', 'auth', 'status', '-o', 'json'], { home, configDir: cfg, env: { MS365_CLI_TOKEN: live() }, preload: [probe] });
    expect(r.status, r.stderr).toBe(0);
    expect(JSON.parse(r.stdout)).toMatchObject({ source: 'external', profile: 'chief', type: 'host', tenantId: TENANT, accountId: OID, upn: 'ceo@contoso.com', scopes: ['Mail.Read', 'User.Read'], ceiling: ['User.Read', 'Mail.Read'], ok: true });
    expect(r.stderr).not.toContain('[probe] msal-node loaded');
    expect(r.fetches).toEqual([]);
  }, 30_000);

  it('AC 32: describe --json detects the agent-host features offline; a bad budget exits 2 before anything runs', () => {
    const { home, configDir: cfg } = entryHome();
    const r = runEntry(['describe', '--json'], { home, configDir: cfg });
    expect(r.status, r.stderr).toBe(0);
    const d = JSON.parse(r.stdout);
    const flags = (opts: Array<{ flags: string }>) => opts.map(o => o.flags);
    expect(flags(d.globalOptions)).toContain('--refs');
    const cmd = (p: string) => d.commands.find((c: { path: string }) => c.path === p);
    expect(flags(cmd('watch').options)).toEqual(expect.arrayContaining([expect.stringContaining('--checkpoint'), expect.stringContaining('--select')]));
    expect(cmd('health')).toBeDefined();
    expect(flags(cmd('auth status').options)).toContain('-o, --output <format>');
    expect(r.fetches).toEqual([]);
    const help = runEntry(['watch', '--help'], { home, configDir: cfg });
    expect(help.stdout).toContain('--checkpoint');
    const bad = runEntry(['--error-format', 'json', 'describe', '--json'], { home, configDir: cfg, env: { MS365_CLI_MAX_RETRIES: '5' } });
    expect(bad.status).toBe(2);
    expect(bad.error?.code).toBe('usage');
  }, 30_000);
});
