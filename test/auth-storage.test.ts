import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { spawn } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import type { Command } from 'commander';
import { AuthManager, configDir, splitTokenCache, writeFileAtomic } from '../src/auth.js';
import { registerAuthCommands } from '../src/commands/auth.js';
import { commandMeta } from '../src/catalog.js';
import { buildTools } from '../src/mcp-server.js';
import { itemKey, type MsalCacheJson } from '../src/credential-store.js';
import { createKeychainStore, serviceName, type KeychainBackend } from '../src/keychain.js';
import { ExitCode } from '../src/errors.js';
import { runCommand } from './helpers/cli.js';
import { keychainStubs, type KeychainStubs } from './helpers/keychain-stubs.js';

/** IDT-3: `auth storage show|use` and migration (ID-5, ID-15), logout/remove cleanup (ID-13), catalog + MCP (ID-16). */

const posix = process.platform !== 'win32';
const FIXTURE = path.resolve(__dirname, 'fixtures', 'msal-cache-3-accounts.json');
const fixture = (): MsalCacheJson => JSON.parse(fs.readFileSync(FIXTURE, 'utf8'));
const CID = '11111111-1111-1111-1111-111111111111';
const ALICE = '00000000-0000-0000-0000-00000000000a.72f988bf-86f1-41af-91ab-2d7cd011db47';
const BOB = '00000000-0000-0000-0000-00000000000b.72f988bf-86f1-41af-91ab-2d7cd011db47';
const CAROL = '00000000-0000-0000-6a1e-5a0d3c3a9b2c.9188040d-6c67-4c5b-b112-36a304b66dad';
const ALL = [ALICE, BOB, CAROL];

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-storage-'));
  vi.stubEnv('MS365_CLI_CONFIG_DIR', dir);
  vi.stubEnv('MS365_CLI_TOKEN', '');
  vi.stubEnv('MS365_CLI_CLIENT_ID', CID);
});
afterEach(() => { vi.unstubAllEnvs(); vi.restoreAllMocks(); });

const indexFile = () => path.join(dir, 'token-cache.json');
const storageJson = () => JSON.parse(fs.readFileSync(path.join(dir, 'storage.json'), 'utf8'));
const register = (p: Command) => registerAuthCommands(p, () => new AuthManager(['User.Read']));
const run = (argv: string[], opts: Parameters<typeof runCommand>[2] = {}) => runCommand(register, argv, opts);

/**
 * Stubbed keychain driven by process.env, with MS365_CLI_STORAGE unset so storage.json decides (as for a user).
 * `items()` is what the stub keychain holds for this config directory.
 */
function stubbed(backend: KeychainBackend = 'libsecret', extra: Record<string, string> = {}) {
  const stubs = keychainStubs(backend, extra);
  for (const [k, v] of Object.entries(stubs.env)) vi.stubEnv(k, v);
  vi.stubEnv('MS365_CLI_STORAGE', '');
  const items = () => backend === 'dpapi'
    ? (() => { try { return fs.readdirSync(path.join(dir, 'credentials')).filter(f => f.endsWith('.dpapi')).map(f => f.slice(0, -6)); } catch { return []; } })()
    : Object.keys(stubs.items(serviceName(dir)));
  return { stubs, items };
}

async function accountIds(): Promise<string[]> {
  const m = new AuthManager(['User.Read']);
  await m.initialize();
  return [...new Set((await m.listAccounts()).map(a => a.homeAccountId))].sort();
}

describe.skipIf(!posix)('auth storage use: migration both ways with no re-login (ID-5, AC 4)', () => {
  it.each(['libsecret', 'macos', 'dpapi'] as KeychainBackend[])('%s: file → keychain → file round-trips every account deep-equal', async backend => {
    const { stubs, items } = stubbed(backend);
    fs.writeFileSync(indexFile(), JSON.stringify(fixture()), { mode: 0o600 });
    const before = JSON.parse(fs.readFileSync(indexFile(), 'utf8'));
    expect(await accountIds()).toEqual([...ALL].sort());

    const toKc = await run(['auth', 'storage', 'use', 'keychain']);
    expect(toKc.stderr).toMatch(/3 account\(s\) moved from the file, no sign-in needed/);
    expect(toKc.exitCode).toBe(0);
    expect(storageJson().storage).toBe('keychain');
    expect(fs.readFileSync(indexFile(), 'utf8')).not.toContain('"secret"');
    expect(items().sort()).toEqual([...ALL].sort());
    expect(stubs.argv()).not.toContain('0.AXXX-refresh');
    expect(await accountIds()).toEqual([...ALL].sort());

    const toFile = await run(['auth', 'storage', 'use', 'file']);
    expect(toFile.stderr).toMatch(/3 account\(s\) moved from the keychain/);
    expect(storageJson().storage).toBe('file');
    expect(JSON.parse(fs.readFileSync(indexFile(), 'utf8'))).toEqual(before);
    expect(items()).toEqual([]);
    expect(await accountIds()).toEqual([...ALL].sort());
  });

  it('re-running use keychain is idempotent; use file when already file touches no keychain tool', async () => {
    const { stubs, items } = stubbed();
    fs.writeFileSync(indexFile(), JSON.stringify(fixture()), { mode: 0o600 });
    await run(['auth', 'storage', 'use', 'keychain']);
    const again = await run(['auth', 'storage', 'use', 'keychain']);
    expect(again.exitCode).toBe(0);
    expect(items().sort()).toEqual([...ALL].sort());
    expect(await accountIds()).toEqual([...ALL].sort());

    await run(['auth', 'storage', 'use', 'file']);
    const calls = stubs.argv();
    const fileAgain = await run(['auth', 'storage', 'use', 'file']);
    expect(fileAgain.exitCode).toBe(0);
    expect(stubs.argv()).toBe(calls);
  });

  it('a keychain that does not read back what was written: nothing switched, the file untouched, written items removed', async () => {
    const { stubs, items } = stubbed();
    // Wrap the stub: store works, lookup returns something else.
    const real = path.join(stubs.bin, 'secret-tool.real');
    fs.renameSync(path.join(stubs.bin, 'secret-tool'), real);
    fs.writeFileSync(path.join(stubs.bin, 'secret-tool'), `#!/bin/bash\nif [ "$1" = lookup ]; then echo garbage; exit 0; fi\nexec "${real}" "$@"\n`, { mode: 0o755 });
    const original = JSON.stringify(fixture());
    fs.writeFileSync(indexFile(), original, { mode: 0o600 });
    const r = await run(['auth', 'storage', 'use', 'keychain']);
    expect(r.exitCode).toBe(ExitCode.GENERIC);
    expect(r.stderr).toMatch(/did not read back/);
    expect(fs.readFileSync(indexFile(), 'utf8')).toBe(original);
    expect(fs.existsSync(path.join(dir, 'storage.json')) ? storageJson().storage : undefined).not.toBe('keychain');
    expect(items()).toEqual([]);
  });

  it('an unusable keychain is exit 3 keychain_unavailable and changes nothing', async () => {
    stubbed('libsecret', { STUB_MODE: 'fail' });
    const original = JSON.stringify(fixture());
    fs.writeFileSync(indexFile(), original, { mode: 0o600 });
    const r = await run(['auth', 'storage', 'use', 'keychain']);
    expect(r.exitCode).toBe(ExitCode.AUTH);
    expect(r.stderr).toMatch(/keychain/i);
    expect(fs.readFileSync(indexFile(), 'utf8')).toBe(original);
    expect(fs.existsSync(path.join(dir, 'storage.json')) ? storageJson().storage : undefined).not.toBe('keychain');

    vi.stubEnv('DBUS_SESSION_BUS_ADDRESS', '');
    const noBus = await run(['auth', 'storage', 'use', 'keychain']);
    expect(noBus.exitCode).toBe(ExitCode.AUTH);
    expect(noBus.stderr).toMatch(/D-Bus/);
  });

  it('use auto (E1) moves keychain credentials back to the file and records no resolution', async () => {
    const { items } = stubbed();
    fs.writeFileSync(indexFile(), JSON.stringify(fixture()), { mode: 0o600 });
    await run(['auth', 'storage', 'use', 'keychain']);
    const r = await run(['auth', 'storage', 'use', 'auto']);
    expect(r.exitCode).toBe(0);
    expect(storageJson()).toEqual({ storage: 'auto' });
    expect(items()).toEqual([]);
    expect(JSON.parse(fs.readFileSync(indexFile(), 'utf8'))).toEqual(fixture());
  });

  it('MS365_CLI_STORAGE overriding storage.json is a usage error (exit 2), not a silent no-op', async () => {
    stubbed();
    vi.stubEnv('MS365_CLI_STORAGE', 'file');
    const r = await run(['auth', 'storage', 'use', 'keychain']);
    expect(r.exitCode).toBe(ExitCode.USAGE);
    expect(r.stderr).toMatch(/MS365_CLI_STORAGE/);
  });

  it('an invalid mode is a usage error', async () => {
    expect((await run(['auth', 'storage', 'use', 'vault'])).exitCode).toBe(ExitCode.USAGE);
  });

  it('--dry-run and --read-only change nothing and spawn no keychain tool', async () => {
    const { stubs, items } = stubbed();
    const original = JSON.stringify(fixture());
    fs.writeFileSync(indexFile(), original, { mode: 0o600 });
    const dry = await run(['auth', 'storage', 'use', 'keychain'], { dryRun: true });
    expect(dry.exitCode).toBe(0);
    expect(dry.stderr).toMatch(/\[dry-run\] would set storage to keychain and move 3 account\(s\) from the file to the keychain/);
    const ro = await run(['auth', 'storage', 'use', 'keychain'], { env: { MS365_CLI_READ_ONLY: '1' } });
    expect(ro.exitCode).toBe(ExitCode.PERMISSION);
    expect(fs.readFileSync(indexFile(), 'utf8')).toBe(original);
    expect(fs.existsSync(path.join(dir, 'storage.json'))).toBe(false);
    expect(items()).toEqual([]);
    expect(stubs.argv()).toBe('');
  });
});

// ── An interrupted migration (the CLI killed mid-move) leaves a usable store ──

const tsx = path.resolve(__dirname, '..', 'node_modules', '.bin', 'tsx');
const cli = path.resolve(__dirname, '..', 'src', 'index.ts');

function spawnCli(args: string[], env: NodeJS.ProcessEnv) {
  return new Promise<{ code: number | null; signal: string | null; stderr: string }>((resolve, reject) => {
    const p = spawn(tsx, [cli, ...args], { env, cwd: dir, stdio: ['ignore', 'pipe', 'pipe'] });
    let stderr = '';
    p.stderr.on('data', d => { stderr += d; });
    p.on('exit', (code, signal) => resolve({ code, signal, stderr }));
    p.on('error', reject);
  });
}

/** Replace the secret-tool stub with one that SIGKILLs the CLI (its parent) on the 2nd call of `op`. */
function killOnSecond(stubs: KeychainStubs, op: string) {
  const real = path.join(stubs.bin, 'secret-tool.real');
  if (!fs.existsSync(real)) fs.renameSync(path.join(stubs.bin, 'secret-tool'), real);
  const counter = path.join(stubs.db, `count-${op}`);
  fs.writeFileSync(path.join(stubs.bin, 'secret-tool'), `#!/bin/bash
if [ "$1" = ${op} ]; then
  n=$(( $(cat "${counter}" 2>/dev/null || echo 0) + 1 )); echo $n > "${counter}"
  if [ $n -eq 2 ]; then kill -9 $PPID; exit 1; fi
fi
exec "${real}" "$@"
`, { mode: 0o755 });
  return () => fs.writeFileSync(path.join(stubs.bin, 'secret-tool'), `#!/bin/bash\nexec "${real}" "$@"\n`, { mode: 0o755 });
}

/** A crashed process's lock is stale: age it past LOCK_STALE_MS so the next run breaks it. */
function ageLocks() {
  for (const f of fs.readdirSync(dir).filter(n => n.endsWith('.lock'))) {
    const old = new Date(Date.now() - 60_000);
    fs.utimesSync(path.join(dir, f), old, old);
  }
}

describe.skipIf(!posix)('interrupted migration (ID-5)', () => {
  it('killed while writing items into the keychain: every account still signs in; re-running completes the move', async () => {
    const { stubs, items } = stubbed();
    fs.writeFileSync(indexFile(), JSON.stringify(fixture()), { mode: 0o600 });
    const restore = killOnSecond(stubs, 'store');
    const env = { ...process.env, MS365_CLI_CONFIG_DIR: dir };
    const r = await spawnCli(['auth', 'storage', 'use', 'keychain'], env);
    expect(r.signal ?? r.code, r.stderr).toBeOneOf(['SIGKILL', 137]);
    restore();
    ageLocks();
    expect(items().length).toBe(1);
    expect(await accountIds()).toEqual([...ALL].sort());

    const again = await run(['auth', 'storage', 'use', 'keychain']);
    expect(again.exitCode).toBe(0);
    expect(fs.readFileSync(indexFile(), 'utf8')).not.toContain('"secret"');
    expect(items().sort()).toEqual([...ALL].sort());
    expect(await accountIds()).toEqual([...ALL].sort());
  }, 30_000);

  it('killed while deleting keychain items after the file was written: still usable; re-running completes the move', async () => {
    const { stubs, items } = stubbed();
    fs.writeFileSync(indexFile(), JSON.stringify(fixture()), { mode: 0o600 });
    await run(['auth', 'storage', 'use', 'keychain']);
    const restore = killOnSecond(stubs, 'clear');
    const r = await spawnCli(['auth', 'storage', 'use', 'file'], { ...process.env, MS365_CLI_CONFIG_DIR: dir });
    expect(r.signal ?? r.code, r.stderr).toBeOneOf(['SIGKILL', 137]);
    restore();
    ageLocks();
    expect(storageJson().storage).toBe('keychain');
    expect(items().length).toBe(2);
    expect(await accountIds()).toEqual([...ALL].sort());

    const again = await run(['auth', 'storage', 'use', 'file']);
    expect(again.exitCode).toBe(0);
    expect(items()).toEqual([]);
    expect(JSON.parse(fs.readFileSync(indexFile(), 'utf8'))).toEqual(fixture());
    expect(await accountIds()).toEqual([...ALL].sort());
  }, 30_000);
});

// ── auth storage show (ID-15) ──

describe('auth storage show (ID-15)', () => {
  it('file store (test default): JSON with every field; the probe is data; exit 0', async () => {
    fs.writeFileSync(indexFile(), JSON.stringify(fixture()), { mode: 0o600 });
    const r = await run(['auth', 'storage', 'show', '-o', 'json']);
    expect(r.exitCode).toBe(0);
    const s = JSON.parse(r.stdout);
    expect(Object.keys(s).sort()).toEqual(['accounts', 'available', 'backend', 'configured', 'files', 'noticeShown', 'probeMs', 'reason', 'resolved', 'source'].sort());
    expect(s).toMatchObject({ configured: 'file', source: 'env', resolved: 'file', backend: 'file', accounts: { count: 3, withItem: 3 }, noticeShown: false });
    expect(typeof s.available).toBe('boolean');
    expect(s.files.map((f: { name: string }) => f.name)).toEqual(['token-cache.json', 'storage.json', 'credentials/']);
    const tc = s.files[0];
    expect(tc).toMatchObject({ path: indexFile(), exists: true });
    if (posix) expect(tc.mode).toBe('0600');
    expect(s.files[1]).toMatchObject({ exists: false, mode: null });
  });

  it.skipIf(!posix)('keychain: reports the backend, a live probe and accounts with items', async () => {
    const { stubs } = stubbed();
    fs.writeFileSync(indexFile(), JSON.stringify(fixture()), { mode: 0o600 });
    await run(['auth', 'storage', 'use', 'keychain']);
    const s = JSON.parse((await run(['auth', 'storage', 'show', '-o', 'json'])).stdout);
    expect(s).toMatchObject({ configured: 'keychain', source: 'file', resolved: 'keychain', backend: 'libsecret', available: true, accounts: { count: 3, withItem: 3 } });
    expect(s.probeMs).toBeGreaterThanOrEqual(0);
    // The probe item was deleted again.
    expect(Object.keys(stubs.items(serviceName(dir))).sort()).toEqual([...ALL].sort());
  });

  it.skipIf(!posix)('an unusable configured keychain is still exit 0: resolved null, available false, reason given', async () => {
    stubbed('libsecret', { STUB_MODE: 'fail' });
    vi.stubEnv('MS365_CLI_STORAGE', 'keychain');
    const r = await run(['auth', 'storage', 'show', '-o', 'json']);
    expect(r.exitCode).toBe(0);
    const s = JSON.parse(r.stdout);
    // The tool is there but fails: the store in effect is still the keychain; the probe says it does not work.
    expect(s).toMatchObject({ configured: 'keychain', resolved: 'keychain', backend: 'libsecret', available: false, accounts: { withItem: null } });
    expect(s.reason).toMatch(/secret-tool/);

    vi.stubEnv('DBUS_SESSION_BUS_ADDRESS', '');
    const noBus = await run(['auth', 'storage', 'show', '-o', 'json']);
    expect(noBus.exitCode).toBe(0);
    expect(JSON.parse(noBus.stdout)).toMatchObject({ configured: 'keychain', resolved: null, backend: 'libsecret', available: false });
    expect(JSON.parse(noBus.stdout).reason).toMatch(/D-Bus/);
  });

  it.skipIf(!posix)('--no-probe, --dry-run and --read-only spawn no keychain tool', async () => {
    const { stubs } = stubbed();
    for (const [argv, opts] of [
      [['auth', 'storage', 'show', '--no-probe', '-o', 'json'], {}],
      [['auth', 'storage', 'show', '-o', 'json'], { dryRun: true }],
      [['auth', 'storage', 'show', '-o', 'json'], { env: { MS365_CLI_READ_ONLY: '1' } }],
    ] as const) {
      const r = await run([...argv], opts);
      expect(r.exitCode).toBe(0);
      expect(JSON.parse(r.stdout)).toMatchObject({ available: null, probeMs: null });
    }
    expect(stubs.argv()).toBe('');
  });

  it('the default table output is readable', async () => {
    const r = await run(['auth', 'storage', 'show', '--no-probe', '-o', 'table']);
    expect(r.exitCode).toBe(0);
    expect(r.stdout).toMatch(/configured/);
  });
});

// ── logout and remove delete keychain items (ID-13, AC 8) ──

describe.skipIf(!posix)('logout and remove clean the keychain (ID-13)', () => {
  it.each(['libsecret', 'macos', 'dpapi'] as KeychainBackend[])('%s: logout deletes every item of the service (orphans too), credentials/, and storage.json resolved', async backend => {
    const { items } = stubbed(backend);
    vi.stubEnv('MS365_CLI_STORAGE', 'keychain');
    const store = createKeychainStore({ backend, configDir: dir, env: process.env, writeFileAtomic });
    splitTokenCache(store).save(JSON.stringify(fixture()));
    store.set(itemKey('orphan-account'), '{}'); // not in the index
    fs.writeFileSync(path.join(dir, 'storage.json'), JSON.stringify({ storage: 'auto', resolved: 'keychain', backend, probedAt: 'x', reason: 'y', noticeShown: true }));
    expect(items().length).toBe(4);

    const r = await run(['auth', 'logout']);
    expect(r.exitCode).toBe(0);
    expect(items()).toEqual([]);
    expect(fs.existsSync(path.join(dir, 'credentials'))).toBe(false);
    expect(fs.existsSync(indexFile())).toBe(false);
    expect(storageJson()).toEqual({ storage: 'auto', noticeShown: true });
  });

  it('logout under --dry-run leaves the items', async () => {
    const { items } = stubbed();
    vi.stubEnv('MS365_CLI_STORAGE', 'keychain');
    splitTokenCache(createKeychainStore({ backend: 'libsecret', configDir: dir, env: process.env, writeFileAtomic })).save(JSON.stringify(fixture()));
    await run(['auth', 'logout'], { dryRun: true });
    expect(items().length).toBe(3);
  });

  it('a file-mode logout never spawns a keychain tool', async () => {
    const { stubs } = stubbed();
    vi.stubEnv('MS365_CLI_STORAGE', 'file');
    fs.writeFileSync(indexFile(), JSON.stringify(fixture()), { mode: 0o600 });
    expect((await run(['auth', 'logout'])).exitCode).toBe(0);
    expect(fs.existsSync(indexFile())).toBe(false);
    expect(stubs.argv()).toBe('');
  });

  it('auth remove deletes that account’s item only; a missing item is not an error', async () => {
    const { items } = stubbed();
    vi.stubEnv('MS365_CLI_STORAGE', 'keychain');
    splitTokenCache(createKeychainStore({ backend: 'libsecret', configDir: dir, env: process.env, writeFileAtomic })).save(JSON.stringify(fixture()));
    const r = await run(['auth', 'remove', BOB]);
    expect(r.exitCode).toBe(0);
    expect(items().sort()).toEqual([ALICE, CAROL].sort());
    expect(fs.readFileSync(indexFile(), 'utf8')).not.toContain(BOB);
    expect(await accountIds()).toEqual([ALICE, CAROL].sort());
  });
});

// ── Catalog and MCP (ID-16) ──

describe('catalog and MCP exposure (ID-16)', () => {
  it('auth storage show is a read, auth storage use an update; neither needs a scope', () => {
    expect(commandMeta('auth storage show')).toMatchObject({ scopes: [], destructive: false, idempotent: true });
    expect(commandMeta('auth storage use')).toMatchObject({ scopes: [], destructive: false });
  });

  it('MCP exposes auth storage show (read-only) but not auth storage use', async () => {
    const { Command } = await import('commander');
    const program = new Command();
    registerAuthCommands(program, () => { throw new Error('no auth'); });
    const names = buildTools(program).map(t => t.name);
    expect(names).toContain('auth_storage_show');
    expect(names).not.toContain('auth_storage_use');
    expect(buildTools(program, { readOnly: true }).map(t => t.name)).toContain('auth_storage_show');
  });
});

it('configDir is the per-test sandbox', () => { expect(configDir()).toBe(dir); });
