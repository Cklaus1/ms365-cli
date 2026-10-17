import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { spawn } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { cachedStore, createFileStore, TOKEN_CACHE_KEY, type CredentialStore } from '../src/credential-store.js';
import { FileTokenCachePlugin, fileCredentialStore, withFileLock, writeFileAtomic } from '../src/auth.js';

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-store-'));
  vi.stubEnv('MS365_CLI_CONFIG_DIR', dir);
});
afterEach(() => { vi.unstubAllEnvs(); vi.restoreAllMocks(); });

const cacheFile = () => path.join(dir, 'token-cache.json');

/** In-memory store that counts backend calls. */
function memoryStore(): CredentialStore & { calls: Record<string, number>; data: Map<string, string> } {
  const data = new Map<string, string>();
  const calls: Record<string, number> = { get: 0, set: 0, delete: 0, list: 0 };
  return {
    kind: 'memory', calls, data,
    get(k) { calls.get++; return data.get(k) || undefined; },
    set(k, v) { calls.set++; data.set(k, v); },
    delete(k) { calls.delete++; data.delete(k); },
    list() { calls.list++; return [...data.keys()].sort(); },
    available() { return true; },
  };
}

describe('file credential store (ID-2: today\'s behaviour)', () => {
  it('token-cache is exactly configDir()/token-cache.json, mode 0600, no temp or lock left behind', () => {
    const store = fileCredentialStore();
    expect(store.kind).toBe('file');
    expect(store.available()).toBe(true);
    expect(store.get(TOKEN_CACHE_KEY)).toBeUndefined();
    store.set(TOKEN_CACHE_KEY, '{"a":1}');
    expect(fs.readFileSync(cacheFile(), 'utf8')).toBe('{"a":1}');
    if (process.platform !== 'win32') expect(fs.statSync(cacheFile()).mode & 0o777).toBe(0o600);
    expect(fs.readdirSync(dir)).toEqual(['token-cache.json']);
    expect(store.get(TOKEN_CACHE_KEY)).toBe('{"a":1}');
  });

  it('follows MS365_CLI_CONFIG_DIR on every call and creates the directory 0700', () => {
    const store = fileCredentialStore();
    const nested = path.join(dir, 'nested', 'cfg');
    vi.stubEnv('MS365_CLI_CONFIG_DIR', nested);
    store.set(TOKEN_CACHE_KEY, 'x');
    expect(fs.readFileSync(path.join(nested, 'token-cache.json'), 'utf8')).toBe('x');
    if (process.platform !== 'win32') expect(fs.statSync(nested).mode & 0o777).toBe(0o700);
  });

  it('an empty file reads as no value; a missing key deletes quietly', () => {
    const store = fileCredentialStore();
    fs.writeFileSync(cacheFile(), '');
    expect(store.get(TOKEN_CACHE_KEY)).toBeUndefined();
    expect(store.list()).toEqual([]);
    fs.unlinkSync(cacheFile());
    expect(() => store.delete(TOKEN_CACHE_KEY)).not.toThrow();
    vi.stubEnv('MS365_CLI_CONFIG_DIR', path.join(dir, 'does-not-exist'));
    expect(() => store.delete(TOKEN_CACHE_KEY)).not.toThrow();
    expect(store.list()).toEqual([]);
  });

  it('list() enumerates only prefixed keys from the directory itself', () => {
    const store = fileCredentialStore();
    store.set(TOKEN_CACHE_KEY, 'a');
    store.set(`${TOKEN_CACHE_KEY}.uid1`, 'b');
    for (const f of ['config.json', 'selected-account.json', 'token-cache.json.123.ab.tmp', 'token-cache.json.lock']) {
      fs.writeFileSync(path.join(dir, f), 'x');
    }
    expect(store.list()).toEqual(['token-cache', 'token-cache.uid1']);
    store.delete(`${TOKEN_CACHE_KEY}.uid1`);
    expect(store.list()).toEqual(['token-cache']);
    expect(fs.existsSync(path.join(dir, 'token-cache.uid1.json'))).toBe(false);
  });

  it('rejects keys outside its prefix or that could escape the directory', () => {
    const store = fileCredentialStore();
    for (const bad of ['config', 'selected-account', '../token-cache', 'token-cache/../x', 'token-cache..x', '', 'token-cache.']) {
      expect(() => store.set(bad, 'x')).toThrow(/Invalid credential store key/);
      expect(() => store.get(bad)).toThrow(/Invalid credential store key/);
      expect(() => store.delete(bad)).toThrow(/Invalid credential store key/);
    }
    expect(fs.readdirSync(dir)).toEqual([]);
  });

  it('set and delete go through the cross-process lock', () => {
    const lock = vi.fn(withFileLock);
    const write = vi.fn(writeFileAtomic);
    const store = createFileStore({ dir: () => dir, writeFileAtomic: write, withFileLock: lock });
    store.set(TOKEN_CACHE_KEY, 'v');
    expect(write).toHaveBeenCalledWith(cacheFile(), 'v');
    store.delete(TOKEN_CACHE_KEY);
    expect(lock).toHaveBeenCalledWith(cacheFile(), expect.any(Function));
    expect(fs.readdirSync(dir)).toEqual([]);
  });

  it('delete waits for a lock held by another process', async () => {
    const store = fileCredentialStore();
    store.set(TOKEN_CACHE_KEY, 'v');
    const lock = `${cacheFile()}.lock`;
    const holder = spawn(process.execPath, ['-e',
      `const fs=require('fs');fs.closeSync(fs.openSync(${JSON.stringify(lock)},'wx'));setTimeout(()=>fs.unlinkSync(${JSON.stringify(lock)}),400);`]);
    const exited = new Promise(r => holder.on('exit', r));
    const until = Date.now() + 5000;
    while (!fs.existsSync(lock) && Date.now() < until) await new Promise(r => setTimeout(r, 5));
    expect(fs.existsSync(lock)).toBe(true);
    const t0 = Date.now();
    store.delete(TOKEN_CACHE_KEY);
    expect(Date.now() - t0).toBeGreaterThanOrEqual(200);
    expect(fs.readdirSync(dir)).toEqual([]);
    await exited;
  });

  it('a stale lock is broken by delete', () => {
    const store = fileCredentialStore();
    store.set(TOKEN_CACHE_KEY, 'v');
    const lock = `${cacheFile()}.lock`;
    fs.writeFileSync(lock, '');
    const old = new Date(Date.now() - 60_000);
    fs.utimesSync(lock, old, old);
    store.delete(TOKEN_CACHE_KEY);
    expect(fs.readdirSync(dir)).toEqual([]);
  });
});

describe('once-per-process read cache (ID-12)', () => {
  it('reads the backend once; own writes and deletes update the cache without re-reading', () => {
    const inner = memoryStore();
    inner.data.set(TOKEN_CACHE_KEY, 'v1');
    const store = cachedStore(inner);
    expect(store.kind).toBe('memory');
    expect(store.get(TOKEN_CACHE_KEY)).toBe('v1');
    expect(store.get(TOKEN_CACHE_KEY)).toBe('v1');
    expect(inner.calls.get).toBe(1);

    inner.data.set(TOKEN_CACHE_KEY, 'other-process'); // not seen: only own writes invalidate
    expect(store.get(TOKEN_CACHE_KEY)).toBe('v1');

    store.set(TOKEN_CACHE_KEY, 'v2');
    expect(store.get(TOKEN_CACHE_KEY)).toBe('v2');
    store.delete(TOKEN_CACHE_KEY);
    expect(store.get(TOKEN_CACHE_KEY)).toBeUndefined();
    expect(inner.calls).toEqual({ get: 1, set: 1, delete: 1, list: 0 });
  });

  it('caches a missing value too, and invalidate() forces a re-read', () => {
    const inner = memoryStore();
    const store = cachedStore(inner);
    expect(store.get(TOKEN_CACHE_KEY)).toBeUndefined();
    inner.data.set(TOKEN_CACHE_KEY, 'relogged');
    expect(store.get(TOKEN_CACHE_KEY)).toBeUndefined();
    expect(inner.calls.get).toBe(1);
    store.invalidate(TOKEN_CACHE_KEY);
    expect(store.get(TOKEN_CACHE_KEY)).toBe('relogged');
    store.invalidate();
    expect(store.get(TOKEN_CACHE_KEY)).toBe('relogged');
    expect(inner.calls.get).toBe(3);
  });

  it('list() and available() pass through uncached', () => {
    const inner = memoryStore();
    const store = cachedStore(inner);
    store.set('a', '1');
    expect(store.list()).toEqual(['a']);
    expect(store.list()).toEqual(['a']);
    expect(inner.calls.list).toBe(2);
    expect(store.available()).toBe(true);
  });

  it('the MSAL plugin reads its store once across many cache accesses and writes only on change', async () => {
    const inner = memoryStore();
    inner.data.set(TOKEN_CACHE_KEY, '{"v":1}');
    const plugin = new FileTokenCachePlugin(inner);
    const deserialized: string[] = [];
    const ctx = (changed: boolean, out = '') => ({
      cacheHasChanged: changed,
      tokenCache: { deserialize: (d: string) => { deserialized.push(d); }, serialize: () => out },
    }) as never;
    for (let i = 0; i < 3; i++) {
      await plugin.beforeCacheAccess(ctx(false));
      await plugin.afterCacheAccess(ctx(false));
    }
    expect(inner.calls.get).toBe(1);
    expect(inner.calls.set).toBe(0);
    await plugin.afterCacheAccess(ctx(true, '{"v":2}'));
    await plugin.beforeCacheAccess(ctx(false));
    expect(inner.calls).toMatchObject({ get: 1, set: 1 });
    expect(deserialized).toEqual(['{"v":1}', '{"v":1}', '{"v":1}', '{"v":2}']);
  });

  it('the default plugin store is the 0600 token-cache.json file', async () => {
    const plugin = new FileTokenCachePlugin();
    expect(plugin.store.kind).toBe('file');
    await plugin.afterCacheAccess({ cacheHasChanged: true, tokenCache: { serialize: () => '{"x":1}' } } as never);
    expect(fs.readFileSync(cacheFile(), 'utf8')).toBe('{"x":1}');
    if (process.platform !== 'win32') expect(fs.statSync(cacheFile()).mode & 0o777).toBe(0o600);
  });
});
