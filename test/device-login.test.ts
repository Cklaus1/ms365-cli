import { describe, it, expect, afterEach, beforeEach, vi } from 'vitest';
import fs from 'fs';
import path from 'path';
import { registerAuthCommands } from '../src/commands/auth.js';
import { configDir, type AuthManager, type LoginOptions } from '../src/auth.js';
import { deviceLoginDeps, holdDeviceLogin } from '../src/device-login.js';
import { setErrorFormat } from '../src/errors.js';
import { runCommand } from './helpers/cli.js';

/** AH-8(5), OQ-12 (a), AC 17: `auth login --start` / `--wait` with the detached holder run in-process. */
const TOKEN = 'eyJ.secret-access-token.sig';
const CODE = { userCode: 'ABCD-EFGH', verificationUri: 'https://microsoft.com/devicelogin', expiresIn: 900, message: 'go there' };

/** A fake AuthManager whose device-code call gives CODE and then waits for `finish` (the mocked token endpoint). */
function fakeAuth() {
  const finish = Promise.withResolvers<string>();
  const auth = {
    initialize: async () => {},
    getClientId: () => '00000000-0000-0000-0000-0000000000aa',
    login: async (o: LoginOptions) => { o.onDeviceCode?.(CODE); return finish.promise; },
  };
  return { auth: auth as unknown as AuthManager, finish };
}

const handleFiles = () => fs.readdirSync(configDir()).filter(f => f.startsWith('device-login-'));

describe('auth login --start / --wait (AH-8(5), AC 17)', () => {
  let held: Promise<void>[];
  let spawned: string[][];
  beforeEach(() => {
    fs.mkdirSync(configDir(), { recursive: true });
    for (const f of handleFiles()) fs.rmSync(path.join(configDir(), f));
    held = [];
    spawned = [];
    setErrorFormat('json');
  });
  afterEach(() => { vi.restoreAllMocks(); setErrorFormat('text'); });

  /** --start with the holder run in-process; returns its JSON and the fake's controls. */
  async function start() {
    const f = fakeAuth();
    vi.spyOn(deviceLoginDeps, 'spawnHolder').mockImplementation((argv) => {
      spawned.push(argv);
      held.push(holdDeviceLogin(f.auth, argv[argv.length - 1]));
    });
    const t0 = Date.now();
    const r = await runCommand((p) => registerAuthCommands(p, () => f.auth), ['auth', 'login', '--start', '-o', 'json']);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(Date.now() - t0).toBeLessThan(2000);
    return { ...f, out: JSON.parse(r.stdout) };
  }
  const wait = (handle: string, ...extra: string[]) => runCommand((p) => registerAuthCommands(p, () => { throw new Error('no AuthManager for --wait'); }),
    ['auth', 'login', '--wait', handle, '-o', 'json', ...extra]);

  it('--start returns the code, URI, expiry, handle and app at once; --wait is pending, then done after the token endpoint answers', async () => {
    const s = await start();
    expect(s.out).toMatchObject({ userCode: 'ABCD-EFGH', verificationUri: CODE.verificationUri, handle: expect.stringMatching(/^[0-9a-f]{32}$/), app: expect.stringContaining('ms365') });
    expect(Date.parse(s.out.expiresAt) - Date.now()).toBeGreaterThan(800_000);
    expect(spawned).toEqual([['auth', 'login', '--hold', s.out.handle]]);
    const [file] = handleFiles();
    expect(fs.statSync(path.join(configDir(), file)).mode & 0o777).toBe(0o600);

    const pending = await wait(s.out.handle, '--timeout', '0');
    expect(JSON.parse(pending.stdout)).toMatchObject({ state: 'pending', handle: s.out.handle });

    s.finish.resolve(TOKEN);
    await Promise.all(held);
    expect(fs.readFileSync(path.join(configDir(), file), 'utf8')).not.toContain(TOKEN);
    const done = await wait(s.out.handle);
    expect(done.exitCode).toBe(0);
    expect(JSON.parse(done.stdout)).toEqual({ state: 'done', handle: s.out.handle });
    expect(handleFiles()).toEqual([]);
  });

  it('--wait reports expired (exit 3 login_expired) once the code has expired', async () => {
    const s = await start();
    const later = Date.parse(s.out.expiresAt) + 1000;
    vi.spyOn(deviceLoginDeps, 'now').mockReturnValue(later);
    const r = await wait(s.out.handle);
    expect(JSON.parse(r.stdout)).toMatchObject({ state: 'expired' });
    expect(r.exitCode).toBe(3);
    expect(r.stderr).toContain('login_expired');
  });

  it('a declined sign-in is denied; the holder records it without a token', async () => {
    const s = await start();
    s.finish.reject(Object.assign(new Error('AADSTS70000: the user declined'), { errorCode: 'authorization_declined' }));
    await Promise.all(held);
    const r = await wait(s.out.handle);
    expect(JSON.parse(r.stdout)).toMatchObject({ state: 'denied' });
    expect(r.exitCode).toBe(3);
  });

  it('an unknown or malformed handle exits 2 unknown_handle; --timeout needs --wait', async () => {
    for (const h of ['0'.repeat(32), '../../etc/passwd']) {
      const r = await wait(h);
      expect(r.exitCode).toBe(2);
      expect(r.stderr).toContain('unknown_handle');
    }
    const r = await runCommand((p) => registerAuthCommands(p, fakeAuth as never), ['auth', 'login', '--timeout', '5']);
    expect(r.exitCode).toBe(2);
  });

  it('is unavailable in external-token mode (exit 2 external_token_mode)', async () => {
    for (const args of [['--start'], ['--wait', 'a'.repeat(32)]]) {
      const r = await runCommand((p) => registerAuthCommands(p, fakeAuth as never), ['auth', 'login', ...args], { env: { MS365_CLI_TOKEN: 'x' } });
      expect(r.exitCode).toBe(2);
      expect(r.stderr).toContain('external_token_mode');
    }
  });
});
