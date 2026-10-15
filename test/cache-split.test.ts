import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { spawn } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import {
  AuthManager, FileTokenCachePlugin, fileCredentialStore, splitTokenCache, withFileLock, writeFileAtomic,
} from '../src/auth.js';
import {
  createFileStore, itemKey, reassembleMsalCache, splitMsalCache, stableStringify, SplitTokenCache,
  type CredentialStore, type MsalCacheJson,
} from '../src/credential-store.js';
import { createKeychainStore, serviceName, KEYCHAIN_BACKENDS, type KeychainBackend } from '../src/keychain.js';
import { CliError, ExitCode } from '../src/errors.js';
import { keychainStubs, type KeychainStubs } from './helpers/keychain-stubs.js';

const posix = process.platform !== 'win32';
const FIXTURE = path.resolve(__dirname, 'fixtures', 'msal-cache-3-accounts.json');
const fixture = (): MsalCacheJson => JSON.parse(fs.readFileSync(FIXTURE, 'utf8'));
const CID = '11111111-1111-1111-1111-111111111111';
const ALICE = '00000000-0000-0000-0000-00000000000a.72f988bf-86f1-41af-91ab-2d7cd011db47';
const BOB = '00000000-0000-0000-0000-00000000000b.72f988bf-86f1-41af-91ab-2d7cd011db47';
const CAROL = '00000000-0000-0000-6a1e-5a0d3c3a9b2c.9188040d-6c67-4c5b-b112-36a304b66dad';

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-split-'));
  vi.stubEnv('MS365_CLI_CONFIG_DIR', dir);
  vi.stubEnv('MS365_CLI_TOKEN', '');
  vi.stubEnv('MS365_CLI_CLIENT_ID', CID);
});
afterEach(() => { vi.unstubAllEnvs(); vi.restoreAllMocks(); });

const indexFile = () => path.join(dir, 'token-cache.json');
const readIndex = () => JSON.parse(fs.readFileSync(indexFile(), 'utf8'));

type Kind = 'file' | KeychainBackend;
/** An item store of each kind: the file store (in `items/`) or a stubbed keychain backend. */
function itemStore(kind: Kind, extra: Record<string, string> = {}): { store: CredentialStore; stubs?: KeychainStubs } {
  if (kind === 'file') {
    fs.mkdirSync(path.join(dir, 'items'), { recursive: true });
    return { store: createFileStore({ dir: () => path.join(dir, 'items'), writeFileAtomic, withFileLock }) };
  }
  const stubs = keychainStubs(kind, extra);
  return { store: createKeychainStore({ backend: kind, configDir: dir, env: stubs.env, writeFileAtomic }), stubs };
}
/** Run the stubbed backend through process.env, as the CLI does. */
function useStubs(stubs: KeychainStubs) {
  for (const [k, v] of Object.entries(stubs.env)) vi.stubEnv(k, v);
}

/** A store wrapper counting calls. */
function counting(inner: CredentialStore) {
  const calls = { get: 0, set: [] as string[], delete: [] as string[] };
  const store: CredentialStore = {
    kind: inner.kind,
    get: k => { calls.get++; return inner.get(k); },
    set: (k, v) => { calls.set.push(k); inner.set(k, v); },
    delete: k => { calls.delete.push(k); inner.delete(k); },
    list: () => inner.list(),
    available: () => inner.available(),
  };
  return { store, calls };
}

describe('split/reassemble is lossless (ID-6)', () => {
  it('the fixture is a real MSAL serialization with 3 accounts (one in two tenants) and every section', () => {
    const c = fixture();
    expect(Object.keys(c).sort()).toEqual(['AccessToken', 'Account', 'AppMetadata', 'IdToken', 'RefreshToken']);
    expect(Object.keys(c.Account)).toHaveLength(4);
    expect(new Set(Object.values(c.Account).map(a => a.home_account_id))).toEqual(new Set([ALICE, BOB, CAROL]));
    for (const s of ['IdToken', 'AccessToken', 'RefreshToken', 'AppMetadata']) expect(Object.keys(c[s]).length).toBeGreaterThan(0);
  });

  it('splits into one item per homeAccountId plus a secret-free index, and reassembles deep-equal', () => {
    const { index, items } = splitMsalCache(fixture());
    expect([...items.keys()].sort()).toEqual([ALICE, BOB, CAROL].sort());
    expect(Object.keys(index).sort()).toEqual(['Account', 'AppMetadata']);
    expect(JSON.stringify(index)).not.toContain('"secret"');
    for (const [hid, item] of items) {
      for (const entries of Object.values(item)) for (const e of Object.values(entries)) expect(e.home_account_id).toBe(hid);
    }
    expect(reassembleMsalCache(index, items.values())).toEqual(fixture());
  });

  it('stableStringify ignores key order', () => {
    expect(stableStringify({ b: 1, a: { d: [2, { y: 1, x: 2 }], c: 3 } })).toBe(stableStringify({ a: { c: 3, d: [2, { x: 2, y: 1 }] }, b: 1 }));
  });
});

describe.skipIf(!posix).each(['file', ...KEYCHAIN_BACKENDS] as Kind[])('[critical-path gate 1] multi-account cache through the %s item store', kind => {
  it('split → store → reassemble in a fresh instance is deep-equal; the index holds no secret', () => {
    const { store } = itemStore(kind);
    const a = splitTokenCache(store);
    expect(a.load()).toBeUndefined();
    a.save(JSON.stringify(fixture()));
    expect(fs.readFileSync(indexFile(), 'utf8')).not.toContain('"secret"');
    expect(Object.keys(readIndex()).sort()).toEqual(['Account', 'AppMetadata']);
    expect(store.list().sort()).toEqual([ALICE, BOB, CAROL].map(itemKey).sort());
    const b = splitTokenCache(store);
    expect(JSON.parse(b.load()!)).toEqual(fixture());
  });

  it('only accounts whose JSON changed are written; a removed account’s item is deleted', () => {
    const { store: inner } = itemStore(kind);
    splitTokenCache(inner).save(JSON.stringify(fixture()));
    const { store, calls } = counting(inner);
    const c = splitTokenCache(store);
    const cache = JSON.parse(c.load()!);
    const at = Object.keys(cache.AccessToken).find(k => cache.AccessToken[k].home_account_id === CAROL)!;
    cache.AccessToken[at].secret = 'refreshed';
    for (const s of ['Account', 'IdToken', 'AccessToken', 'RefreshToken']) {
      for (const k of Object.keys(cache[s])) if (cache[s][k].home_account_id === BOB) delete cache[s][k];
    }
    c.save(JSON.stringify(cache));
    expect(calls.set).toEqual([itemKey(CAROL)]);
    expect(calls.delete).toEqual([itemKey(BOB)]);
    expect(JSON.parse(splitTokenCache(inner).load()!)).toEqual(cache);
  });
});

describe.skipIf(!posix)('the plugin and AuthManager with a keychain item store', () => {
  it('a file cache moves into per-account items on the first write and MSAL still sees every account', async () => {
    const stubs = keychainStubs('libsecret');
    useStubs(stubs);
    fs.writeFileSync(indexFile(), JSON.stringify(fixture()), { mode: 0o600 });
    const a = new AuthManager(['User.Read']);
    await a.initialize();
    const ids = (await a.listAccounts()).map(x => x.homeAccountId);
    expect(new Set(ids)).toEqual(new Set([ALICE, BOB, CAROL]));
    expect(await a.removeAccount(BOB)).toBe(true);
    expect(fs.readFileSync(indexFile(), 'utf8')).not.toContain('"secret"');
    expect(Object.keys(stubs.items(serviceName(dir))).sort()).toEqual([ALICE, CAROL].sort());
    expect(stubs.argv()).not.toContain('0.AXXX-refresh');

    const b = new AuthManager(['User.Read']);
    await b.initialize();
    expect(new Set((await b.listAccounts()).map(x => x.homeAccountId))).toEqual(new Set([ALICE, CAROL]));
    // Exactly one read per item per process (ID-12), however many cache accesses MSAL makes.
    const before = stubs.argv().split('\n').filter(l => l.startsWith('secret-tool lookup')).length;
    await b.listAccounts();
    await b.activeAccountId();
    expect(stubs.argv().split('\n').filter(l => l.startsWith('secret-tool lookup')).length).toBe(before);
  });

  it('under explicit keychain, a failing backend is exit 3 keychain_unavailable from the CLI path, not a crash', async () => {
    const stubs = keychainStubs('libsecret', { STUB_MODE: 'fail' });
    useStubs(stubs);
    const one = fixture();
    fs.writeFileSync(indexFile(), JSON.stringify({ Account: one.Account, AppMetadata: one.AppMetadata }), { mode: 0o600 });
    const m = new AuthManager(['User.Read']);
    await m.initialize();
    const err = await m.listAccounts().catch(e => e);
    expect(err).toBeInstanceOf(CliError);
    expect((err as CliError).code).toBe('keychain_unavailable');
    expect((err as CliError).exitCode).toBe(ExitCode.AUTH);
  });

  it('with the file store (the default), the plugin keeps today’s single token-cache.json', async () => {
    const plugin = new FileTokenCachePlugin(fileCredentialStore(), () => null);
    await plugin.afterCacheAccess({ cacheHasChanged: true, tokenCache: { serialize: () => JSON.stringify(fixture()) } } as never);
    expect(plugin.mode()).toBe('file');
    expect(readIndex()).toEqual(fixture());
  });
});

describe.skipIf(!posix)('[critical-path gate 4] missing and corrupt items (ID-18)', () => {
  const seed = (kind: Kind) => {
    const r = itemStore(kind);
    splitTokenCache(r.store).save(JSON.stringify(fixture()));
    return r;
  };
  const accountsIn = (json: string | undefined) => new Set(Object.values(JSON.parse(json ?? '{"Account":{}}').Account).map(a => (a as { home_account_id: string }).home_account_id));

  it.each(['file', 'libsecret'] as Kind[])('%s: a corrupt item is deleted and its account dropped; others intact; the next write cleans the index', kind => {
    const { store } = seed(kind);
    store.set(itemKey(BOB), '{not json');
    const c = splitTokenCache(store);
    const loaded = c.load();
    expect(accountsIn(loaded)).toEqual(new Set([ALICE, CAROL]));
    expect(store.get(itemKey(BOB))).toBeUndefined();
    expect(store.list()).not.toContain(itemKey(BOB));
    c.save(loaded!);
    expect(accountsIn(fs.readFileSync(indexFile(), 'utf8'))).toEqual(new Set([ALICE, CAROL]));
  });

  it('an item holding another account’s entries is corrupt', () => {
    const { store } = seed('file');
    store.set(itemKey(BOB), store.get(itemKey(ALICE))!);
    expect(accountsIn(splitTokenCache(store).load())).toEqual(new Set([ALICE, CAROL]));
    expect(store.get(itemKey(BOB))).toBeUndefined();
  });

  it('a missing item drops its account; if another process wrote it meanwhile, the next save keeps it', () => {
    const { store } = seed('file');
    const bobItem = store.get(itemKey(BOB))!;
    store.delete(itemKey(BOB));
    const c = splitTokenCache(store);
    const loaded = c.load();
    expect(accountsIn(loaded)).toEqual(new Set([ALICE, CAROL]));
    store.set(itemKey(BOB), bobItem); // re-login elsewhere (it also rewrote the index entry, still present)
    c.save(loaded!);
    expect(accountsIn(splitTokenCache(store).load())).toEqual(new Set([ALICE, BOB, CAROL]));
  });

  it('macOS (non-base64) and DPAPI (does not decrypt) corrupt items are treated as missing', () => {
    for (const kind of ['macos', 'dpapi'] as const) {
      dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-split-'));
      vi.stubEnv('MS365_CLI_CONFIG_DIR', dir);
      const { store, stubs } = seed(kind);
      if (kind === 'macos') fs.writeFileSync(path.join(stubs!.db, 'macos', serviceName(dir), BOB), '!!garbage!!');
      else fs.writeFileSync(path.join(dir, 'credentials', `${BOB}.dpapi`), Buffer.from('foreign').toString('base64'));
      expect(accountsIn(splitTokenCache(store).load())).toEqual(new Set([ALICE, CAROL]));
      expect(store.list()).not.toContain(itemKey(BOB));
    }
  });

  it('the only account’s item corrupt → today’s "Not logged in" exit 3', async () => {
    const stubs = keychainStubs('libsecret');
    useStubs(stubs);
    const c = fixture();
    const one: MsalCacheJson = { Account: {}, IdToken: {}, AccessToken: {}, RefreshToken: {}, AppMetadata: c.AppMetadata };
    for (const s of ['Account', 'IdToken', 'AccessToken', 'RefreshToken']) {
      for (const [k, e] of Object.entries(c[s])) if (e.home_account_id === CAROL) one[s][k] = e;
    }
    splitTokenCache(createKeychainStore({ backend: 'libsecret', configDir: dir, env: stubs.env, writeFileAtomic })).save(JSON.stringify(one));
    fs.writeFileSync(path.join(stubs.db, 'libsecret', serviceName(dir), CAROL), 'garbage');
    const m = new AuthManager(['User.Read']);
    await m.initialize();
    const err = await m.getToken().catch(e => e);
    expect(err).toBeInstanceOf(CliError);
    expect((err as CliError).message).toBe('Not logged in. Run: ms365 auth login');
    expect((err as CliError).exitCode).toBe(ExitCode.AUTH);
  });

  it('a corrupt index file starts empty instead of crashing', () => {
    const { store } = itemStore('file');
    fs.writeFileSync(indexFile(), '{not json', { mode: 0o600 });
    expect(splitTokenCache(store).load()).toBeUndefined();
  });
});

const tsx = path.resolve(__dirname, '..', 'node_modules', '.bin', 'tsx');

function spawnCollect(cmd: string, args: string[], env: NodeJS.ProcessEnv, cwd = os.tmpdir()) {
  return new Promise<{ code: number; stdout: string; stderr: string; ms: number }>((resolve, reject) => {
    const start = Date.now();
    const p = spawn(cmd, args, { env, cwd, stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '';
    let stderr = '';
    p.stdout.on('data', d => { stdout += d; });
    p.stderr.on('data', d => { stderr += d; });
    p.on('exit', code => resolve({ code: code ?? 1, stdout, stderr, ms: Date.now() - start }));
    p.on('error', reject);
  });
}

describe.skipIf(!posix)('[critical-path gate 2] 8 concurrent processes writing different accounts', () => {
  it.each(['libsecret', 'file-items'])('%s: the store ends valid and complete, with no temp or lock files', async mode => {
    const stubs = keychainStubs('libsecret');
    const env = { ...process.env, ...stubs.env, MS365_CLI_CONFIG_DIR: dir };
    const script = path.resolve(__dirname, 'fixtures', 'split-writer.ts');
    const ROUNDS = 4;
    const barrier = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-barrier-'));
    const running = Promise.all(Array.from({ length: 8 }, (_, i) => spawnCollect(tsx, [script, String(i), String(ROUNDS), mode, barrier], env)));
    const deadline = Date.now() + 20_000;
    while (fs.readdirSync(barrier).length < 8 && Date.now() < deadline) await new Promise(r => setTimeout(r, 10));
    fs.writeFileSync(path.join(barrier, 'go'), '');
    const results = await running;
    for (const r of results) expect(r.stderr).toBe('');
    expect(results.map(r => r.code)).toEqual(Array(8).fill(0));

    const index = readIndex();
    expect(JSON.stringify(index)).not.toContain('"secret"');
    const ids = Array.from({ length: 8 }, (_, i) => Array.from({ length: ROUNDS }, (_, r) => `uid${i}x${r}.utid`)).flat();
    expect(new Set(Object.values(index.Account).map(a => (a as { home_account_id: string }).home_account_id))).toEqual(new Set(ids));
    expect(Object.keys(index.AppMetadata)).toHaveLength(1);

    const store = mode === 'file-items'
      ? createFileStore({ dir: () => path.join(dir, 'items'), writeFileAtomic, withFileLock })
      : createKeychainStore({ backend: 'libsecret', configDir: dir, env: stubs.env, writeFileAtomic });
    expect(store.list().sort()).toEqual(ids.map(itemKey).sort());
    const all = JSON.parse(splitTokenCache(store).load()!);
    const rtOf = (hid: string) => (Object.values(all.RefreshToken).find(e => (e as { home_account_id: string }).home_account_id === hid) as { secret: string }).secret;
    for (let i = 0; i < 8; i++) {
      expect(rtOf(`uid${i}x0.utid`)).toBe(`rt-uid${i}x0.utid-${ROUNDS - 1}`); // refreshed every round, latest kept
      for (let r = 1; r < ROUNDS; r++) expect(rtOf(`uid${i}x${r}.utid`)).toBe(`rt-uid${i}x${r}.utid-${r}`);
    }
    const leftovers = (d: string) => fs.readdirSync(d).filter(f => /\.(lock|tmp)$|\.tmp\./.test(f));
    expect(leftovers(dir)).toEqual([]);
    if (mode === 'file-items') expect(leftovers(path.join(dir, 'items'))).toEqual([]);
  }, 60_000);
});

describe.skipIf(!posix)('[critical-path gate 3] a hanging keychain tool never delays CLI exit', () => {
  it('the spawned CLI exits 3 keychain_unavailable shortly after the timeout; the stub never saw the secret', async () => {
    const stubs = keychainStubs('libsecret', { STUB_MODE: 'hang', MS365_CLI_KEYCHAIN_TIMEOUT_MS: '400' });
    const c = fixture();
    fs.writeFileSync(indexFile(), JSON.stringify({ Account: c.Account, AppMetadata: c.AppMetadata }), { mode: 0o600 });
    const env = { ...process.env, ...stubs.env, MS365_CLI_CONFIG_DIR: dir, MS365_CLI_CLIENT_ID: CID, MS365_CLI_TOKEN: '' };
    const cli = path.resolve(__dirname, '..', 'src', 'index.ts');
    const r = await spawnCollect(tsx, [cli, 'auth', 'accounts'], env, dir);
    expect(r.code).toBe(ExitCode.AUTH);
    expect(r.stderr).toMatch(/keychain/i);
    expect(r.stderr).toMatch(/killed/);
    // tsx start-up (~1 s) + one 400 ms timeout; a hung child would hold the process for 30 s.
    expect(r.ms).toBeLessThan(8_000);
    expect(stubs.argv()).toMatch(/^secret-tool lookup service ms365-cli:[0-9a-f]{8} account /m);
    expect(stubs.argv()).not.toContain('refresh');
  }, 30_000);
});

describe('SplitTokenCache invalidate()', () => {
  it('re-reads the stores after invalidate()', () => {
    const items = new Map<string, string>();
    const store: CredentialStore = {
      kind: 'memory', get: k => items.get(k), set: (k, v) => { items.set(k, v); }, delete: k => { items.delete(k); },
      list: () => [...items.keys()], available: () => true,
    };
    let index: string | undefined;
    const c = new SplitTokenCache({ index: { read: () => index, write: d => { index = d; }, remove: () => { index = undefined; } }, items: store, lock: fn => fn() });
    expect(c.load()).toBeUndefined();
    new SplitTokenCache({ index: { read: () => index, write: d => { index = d; }, remove: () => { index = undefined; } }, items: store, lock: fn => fn() })
      .save(JSON.stringify(fixture()));
    expect(c.load()).toBeUndefined(); // read once
    c.invalidate();
    expect(JSON.parse(c.load()!)).toEqual(fixture());
  });
});
