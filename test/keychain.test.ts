import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import {
  configuredStorage, createKeychainStore, encodePowerShell, keychainUnavailable, platformBackend, readStorageConfig,
  resolveStorage, runTool, serviceName, storageConfigPath, KEYCHAIN_BACKENDS, type KeychainBackend,
} from '../src/keychain.js';
import { CorruptItemError, ensurePrivateMode } from '../src/credential-store.js';
import { configDir, writeFileAtomic } from '../src/auth.js';
import { CliError, ExitCode } from '../src/errors.js';
import { keychainStubs } from './helpers/keychain-stubs.js';

const posix = process.platform !== 'win32';
let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-kcfg-'));
  vi.stubEnv('MS365_CLI_CONFIG_DIR', dir);
});
afterEach(() => { vi.unstubAllEnvs(); vi.restoreAllMocks(); });

const store = (backend: KeychainBackend, extra: Record<string, string> = {}, timeoutMs?: number) => {
  const stubs = keychainStubs(backend, extra);
  const s = createKeychainStore({ backend, configDir: dir, env: stubs.env, writeFileAtomic, timeoutMs });
  return { stubs, s };
};
const SECRET = '{"RefreshToken":{"k":{"home_account_id":"u.t","secret":"rt-SUPER-SECRET-0123456789"}}}';
const KEY = 'token-cache.u.t';
const writeBlob = (data: string) => {
  fs.mkdirSync(path.join(dir, 'credentials'), { recursive: true });
  fs.writeFileSync(path.join(dir, 'credentials', 'u.t.dpapi'), data, { mode: 0o600 });
};

describe('ID-14: the test environment never reaches a real keychain', () => {
  it('setup-env forces the file store', () => {
    expect(process.env.MS365_CLI_STORAGE).toBe('file');
    expect(resolveStorage({ configDir: configDir(), phase: 'write', writeFileAtomic }).kind).toBe('file');
  });
});

describe.skipIf(!posix).each(KEYCHAIN_BACKENDS)('keychain backend %s (stub tools on PATH)', backend => {
  it('set → get → list → delete round-trips; a missing item is undefined, deleting it again is fine', () => {
    const { s } = store(backend);
    expect(s.available()).toBe(true);
    expect(s.get(KEY)).toBeUndefined();
    s.set(KEY, SECRET);
    s.set('token-cache.v.t', 'second');
    expect(s.get(KEY)).toBe(SECRET);
    expect(s.list()).toEqual(['token-cache.u.t', 'token-cache.v.t']);
    s.set(KEY, SECRET.replace('0123', '9999')); // replace in place
    expect(s.get(KEY)).toBe(SECRET.replace('0123', '9999'));
    s.delete(KEY);
    s.delete(KEY);
    expect(s.get(KEY)).toBeUndefined();
    expect(s.list()).toEqual(['token-cache.v.t']);
  });

  it('the secret never appears in any argv (ID-3)', () => {
    const { s, stubs } = store(backend);
    s.set(KEY, SECRET);
    s.get(KEY);
    s.list();
    s.delete(KEY);
    const argv = stubs.argv();
    if (backend !== 'dpapi') expect(argv).not.toBe(''); // the tool really ran
    expect(argv).not.toContain('rt-SUPER-SECRET');
    expect(argv).not.toContain(Buffer.from(SECRET).toString('base64').slice(0, 24));
  });

  it('a hanging tool is SIGKILLed at the timeout and surfaces as exit 3 keychain_unavailable', () => {
    const { s } = store(backend, { STUB_MODE: 'hang' }, 300);
    if (backend === 'dpapi') writeBlob('U1RVQkRQQVBJOmVB');
    const t = Date.now();
    const err = (() => { try { s.get(KEY); } catch (e) { return e; } })() as CliError;
    expect(Date.now() - t).toBeLessThan(2_500);
    expect(err).toBeInstanceOf(CliError);
    expect(err.code).toBe('keychain_unavailable');
    expect(err.exitCode).toBe(ExitCode.AUTH);
    expect(err.message).toMatch(/killed/);
    expect(err.fix).toMatch(/ms365 auth storage use file/);
  });

  it('a failing tool is exit 3, not a crash', () => {
    const { s } = store(backend, { STUB_MODE: 'fail' });
    expect(() => s.set(KEY, SECRET)).toThrow(expect.objectContaining({ code: 'keychain_unavailable', exitCode: ExitCode.AUTH }));
  });

  it('a missing tool makes the backend unavailable (no spawn), and use is exit 3', () => {
    const stubs = keychainStubs(backend);
    const empty = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-nopath-'));
    const s = createKeychainStore({ backend, configDir: dir, env: { ...stubs.env, PATH: empty }, writeFileAtomic });
    expect(s.available()).toBe(false);
    expect(s.unavailableReason()).toMatch(/not found on PATH/);
    expect(() => s.set(KEY, 'x')).toThrow(expect.objectContaining({ code: 'keychain_unavailable' }));
  });
});

describe.skipIf(!posix)('backend specifics', () => {
  it('macOS writes through `security -i` with the base64 value on stdin and items are keyed by service + homeAccountId', () => {
    const { s, stubs } = store('macos');
    s.set(KEY, SECRET);
    expect(stubs.argv()).toMatch(/^security -i$/m);
    const items = stubs.items(serviceName(dir));
    expect(Object.keys(items)).toEqual(['u.t']);
    expect(Buffer.from(items['u.t'], 'base64').toString('utf8')).toBe(SECRET);
  });

  it('macOS: a non-base64 item is reported as corrupt (ID-18)', () => {
    const { s, stubs } = store('macos');
    fs.mkdirSync(path.join(stubs.db, 'macos', serviceName(dir)), { recursive: true });
    fs.writeFileSync(path.join(stubs.db, 'macos', serviceName(dir), 'u.t'), 'not base64 !!');
    expect(() => s.get(KEY)).toThrow(CorruptItemError);
  });

  it('libsecret needs a D-Bus session: without DBUS_SESSION_BUS_ADDRESS it is unavailable and never spawns', () => {
    const stubs = keychainStubs('libsecret');
    const s = createKeychainStore({ backend: 'libsecret', configDir: dir, env: { ...stubs.env, DBUS_SESSION_BUS_ADDRESS: '' }, writeFileAtomic });
    expect(s.available()).toBe(false);
    expect(s.unavailableReason()).toMatch(/D-Bus/);
    expect(() => s.get(KEY)).toThrow(expect.objectContaining({ code: 'keychain_unavailable' }));
    expect(stubs.argv()).toBe('');
  });

  it('libsecret stores the secret via stdin (store … service <svc> account <acct>)', () => {
    const { s, stubs } = store('libsecret');
    s.set(KEY, SECRET);
    expect(stubs.argv()).toContain(`secret-tool store --label=ms365-cli token cache (${serviceName(dir)}) service ${serviceName(dir)} account u.t`);
    expect(stubs.items(serviceName(dir))).toEqual({ 'u.t': SECRET });
  });

  it('DPAPI keeps the blob in credentials/<acct>.dpapi (dir 0700, file 0600); the script goes in -EncodedCommand', () => {
    const { s, stubs } = store('dpapi');
    s.set(KEY, SECRET);
    const blob = path.join(dir, 'credentials', 'u.t.dpapi');
    expect(fs.readFileSync(blob, 'utf8')).not.toContain('SUPER-SECRET');
    expect(fs.statSync(path.join(dir, 'credentials')).mode & 0o777).toBe(0o700);
    expect(fs.statSync(blob).mode & 0o777).toBe(0o600);
    expect(stubs.argv()).toMatch(/^powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand [A-Za-z0-9+/=]+$/m);
    // pwsh is preferred when present: the stub dir has both names, so the stubbed pwsh ran.
    expect(fs.existsSync(path.join(stubs.bin, 'pwsh'))).toBe(true);
    expect(s.get(KEY)).toBe(SECRET);
    expect(s.list()).toEqual([KEY]);
  });

  it('DPAPI: a blob that does not decrypt is corrupt (ID-18)', () => {
    const { s } = store('dpapi');
    writeBlob(Buffer.from('foreign').toString('base64'));
    expect(() => s.get(KEY)).toThrow(CorruptItemError);
  });

  it('encodePowerShell is base64 UTF-16LE', () => {
    expect(Buffer.from(encodePowerShell('exit 0'), 'base64').toString('utf16le')).toBe('exit 0');
  });
});

describe('runner and naming', () => {
  it.skipIf(!posix)('runTool kills a hung child with SIGKILL at the timeout', () => {
    const t = Date.now();
    const r = runTool('/bin/sh', ['-c', 'exec sleep 30'], { timeoutMs: 200 });
    expect(r.failure).toBe('timeout');
    expect(Date.now() - t).toBeLessThan(2_000);
  });

  it('runTool reports a missing tool as missing', () => {
    expect(runTool(path.join(dir, 'no-such-tool'), [], { timeoutMs: 1_000 }).failure).toBe('missing');
  });

  it('the service name is ms365-cli:<8 hex of sha256(configDir)> and differs per config dir', () => {
    expect(serviceName('/a')).toMatch(/^ms365-cli:[0-9a-f]{8}$/);
    expect(serviceName('/a')).not.toBe(serviceName('/b'));
  });

  it('platform backends: darwin → macos, win32 → dpapi, linux → libsecret; override validated', () => {
    expect(platformBackend({}, 'darwin')).toBe('macos');
    expect(platformBackend({}, 'win32')).toBe('dpapi');
    expect(platformBackend({}, 'linux')).toBe('libsecret');
    expect(platformBackend({}, 'aix')).toBeNull();
    expect(platformBackend({ MS365_CLI_KEYCHAIN_BACKEND: 'dpapi' }, 'linux')).toBe('dpapi');
    expect(() => platformBackend({ MS365_CLI_KEYCHAIN_BACKEND: 'kwallet' }, 'linux')).toThrow(/Invalid/);
  });

  it('keychain_unavailable suggests unlocking on macOS', () => {
    expect(keychainUnavailable('locked', 'macos').fix).toMatch(/security unlock-keychain/);
  });

  it('lint: the backends never put a value after -w in argv, and spawn only through runTool', () => {
    const src = fs.readFileSync(path.resolve(__dirname, '..', 'src', 'keychain.ts'), 'utf8');
    expect(src).not.toMatch(/['"]-w['"]\s*,\s*[^\]\s]/);
    expect(src.match(/spawnSync\(/g)).toHaveLength(1);
    expect(src).not.toMatch(/(?<![.\w])(execSync|execFileSync|spawn|exec|execFile)\(/);
  });
});

describe('storage.json and the storage choice (ID-4, E1, E3)', () => {
  const noEnv = (extra: Record<string, string> = {}) => ({ PATH: process.env.PATH ?? '', ...extra });

  it('MS365_CLI_STORAGE wins over storage.json; invalid values are exit 2', () => {
    fs.writeFileSync(storageConfigPath(dir), JSON.stringify({ storage: 'keychain' }), { mode: 0o600 });
    expect(configuredStorage(dir, noEnv())).toEqual({ setting: 'keychain', source: 'file' });
    expect(configuredStorage(dir, noEnv({ MS365_CLI_STORAGE: 'file' }))).toEqual({ setting: 'file', source: 'env' });
    expect(configuredStorage(path.join(dir, 'none'), noEnv())).toEqual({ setting: 'auto', source: 'default' });
    expect(() => configuredStorage(dir, noEnv({ MS365_CLI_STORAGE: 'vault' }))).toThrow(expect.objectContaining({ exitCode: ExitCode.USAGE }));
  });

  it('E1: auto resolves to the file store everywhere, records nothing and prints nothing', () => {
    const notice = vi.fn();
    for (const platform of ['darwin', 'win32', 'linux'] as const) {
      const r = resolveStorage({ configDir: dir, phase: 'write', env: noEnv(), platform, writeFileAtomic, notice });
      expect(r.kind).toBe('file');
    }
    expect(notice).not.toHaveBeenCalled();
    expect(fs.existsSync(storageConfigPath(dir))).toBe(false);
  });

  it.skipIf(!posix)('explicit keychain with a working (stub) backend → keychain; storage.json storage=keychain works too', () => {
    const stubs = keychainStubs('libsecret');
    expect(resolveStorage({ configDir: dir, phase: 'read', env: stubs.env, writeFileAtomic }).kind).toBe('keychain');
    fs.writeFileSync(storageConfigPath(dir), JSON.stringify({ storage: 'keychain' }), { mode: 0o600 });
    const env = { ...stubs.env, MS365_CLI_STORAGE: '' };
    expect(resolveStorage({ configDir: dir, phase: 'read', env, writeFileAtomic }).store?.backend).toBe('libsecret');
  });

  it('explicit keychain with no usable backend is exit 3 keychain_unavailable', () => {
    const env = noEnv({ MS365_CLI_STORAGE: 'keychain', MS365_CLI_KEYCHAIN_BACKEND: 'libsecret', DBUS_SESSION_BUS_ADDRESS: '' });
    expect(() => resolveStorage({ configDir: dir, phase: 'read', env, writeFileAtomic }))
      .toThrow(expect.objectContaining({ code: 'keychain_unavailable', exitCode: ExitCode.AUTH }));
  });

  it.skipIf(!posix)('auto with the probe enabled: resolves once on the first write, sticky, recorded in storage.json', () => {
    const stubs = keychainStubs('libsecret');
    const env = { ...stubs.env, MS365_CLI_STORAGE: 'auto' };
    const opts = { configDir: dir, env, writeFileAtomic, autoPrefersKeychain: true };
    expect(resolveStorage({ ...opts, phase: 'read' }).kind).toBe('file'); // reads never probe
    expect(fs.existsSync(storageConfigPath(dir))).toBe(false);
    expect(resolveStorage({ ...opts, phase: 'write' }).kind).toBe('keychain');
    const cfg = readStorageConfig(dir);
    expect(cfg).toMatchObject({ resolved: 'keychain', backend: 'libsecret', reason: expect.stringMatching(/probe passed/) });
    expect(cfg.probedAt).toBeTruthy();
    expect(Object.keys(stubs.items(serviceName(dir)))).toEqual([]); // the probe item was deleted
    const calls = stubs.argv().split('\n').length;
    expect(resolveStorage({ ...opts, phase: 'write' }).kind).toBe('keychain'); // sticky: no second probe
    expect(stubs.argv().split('\n').length).toBe(calls);
  });

  it('auto fallback prints the notice once per config directory and sticks to the file store', () => {
    const env = noEnv({ MS365_CLI_STORAGE: 'auto', MS365_CLI_KEYCHAIN_BACKEND: 'libsecret', DBUS_SESSION_BUS_ADDRESS: '' });
    const notice = vi.fn();
    const opts = { configDir: dir, env, writeFileAtomic, autoPrefersKeychain: true, notice, phase: 'write' as const };
    expect(resolveStorage(opts).kind).toBe('file');
    expect(resolveStorage(opts).kind).toBe('file');
    expect(notice).toHaveBeenCalledTimes(1);
    expect(notice.mock.calls[0][0]).toMatch(/^ms365: keychain unavailable \(no D-Bus session.*\); using the 0600 file — ms365 auth storage show$/);
    expect(readStorageConfig(dir)).toMatchObject({ resolved: 'file', noticeShown: true });
    // A re-probe (storage.json `resolved` cleared, as `auth storage use auto` will do) keeps noticeShown: no second notice.
    const cfg = readStorageConfig(dir);
    delete cfg.resolved;
    fs.writeFileSync(storageConfigPath(dir), JSON.stringify(cfg));
    resolveStorage(opts);
    expect(notice).toHaveBeenCalledTimes(1);
  });

  it.skipIf(!posix)('E3: on Windows/DPAPI, auto resolves to file when PowerShell takes > 400 ms to start', () => {
    const stubs = keychainStubs('dpapi', { MS365_CLI_STORAGE: 'auto', STUB_PS_START_S: '0.6' });
    const notice = vi.fn();
    const r = resolveStorage({ configDir: dir, phase: 'write', env: stubs.env, writeFileAtomic, autoPrefersKeychain: true, notice });
    expect(r.kind).toBe('file');
    expect(r.reason).toMatch(/PowerShell took \d+ ms to start \(limit 400 ms\)/);
    expect(notice).toHaveBeenCalledTimes(1);
  });

  it.skipIf(!posix)('E3: a fast PowerShell start lets auto use DPAPI', () => {
    const stubs = keychainStubs('dpapi', { MS365_CLI_STORAGE: 'auto' });
    const r = resolveStorage({ configDir: dir, phase: 'write', env: stubs.env, writeFileAtomic, autoPrefersKeychain: true });
    expect(r.kind).toBe('keychain');
  });

  it('auto that resolved to keychain but finds it unavailable later is exit 3 (no silent split-brain)', () => {
    fs.writeFileSync(storageConfigPath(dir), JSON.stringify({ resolved: 'keychain', backend: 'libsecret' }), { mode: 0o600 });
    const env = noEnv({ MS365_CLI_STORAGE: 'auto', MS365_CLI_KEYCHAIN_BACKEND: 'libsecret', DBUS_SESSION_BUS_ADDRESS: '' });
    expect(() => resolveStorage({ configDir: dir, phase: 'read', env, writeFileAtomic }))
      .toThrow(expect.objectContaining({ code: 'keychain_unavailable' }));
  });
});

describe.skipIf(!posix)('ID-19: file modes are checked on read', () => {
  it('a group/world-readable credential file is chmod 0600 with one stderr line', () => {
    const f = path.join(dir, 'token-cache.json');
    fs.writeFileSync(f, '{}', { mode: 0o644 });
    fs.chmodSync(f, 0o644);
    const warn = vi.fn();
    ensurePrivateMode(f, warn);
    expect(fs.statSync(f).mode & 0o777).toBe(0o600);
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0][0]).toMatch(/mode 0644\); changed it to 0600/);
    ensurePrivateMode(f, warn);
    expect(warn).toHaveBeenCalledTimes(1);
    ensurePrivateMode(path.join(dir, 'missing.json'), warn);
    expect(warn).toHaveBeenCalledTimes(1);
  });

  it('storage.json is checked when read', () => {
    fs.writeFileSync(storageConfigPath(dir), '{"storage":"file"}');
    fs.chmodSync(storageConfigPath(dir), 0o664);
    const spy = vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
    expect(readStorageConfig(dir).storage).toBe('file');
    expect(fs.statSync(storageConfigPath(dir)).mode & 0o777).toBe(0o600);
    expect(spy).toHaveBeenCalledWith(expect.stringMatching(/storage\.json was accessible to other users/));
  });
});
