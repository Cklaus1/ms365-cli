import fs from 'fs';
import path from 'path';

/**
 * Storage seam for credentials (identity-profiles spec ID-2). Values are opaque UTF-8 strings; keys are short
 * identifiers chosen by the caller (today only `token-cache`). Every method is synchronous: backends either touch
 * a local file or run their OS tool in a child process with a spawnSync timeout (ID-3), never a native module.
 */
export interface CredentialStore {
  /** Which backend this is (reported by `auth storage show` later). */
  readonly kind: string;
  /** The stored value, or `undefined` when there is none (an empty value counts as none). */
  get(key: string): string | undefined;
  /** Create or replace the value for `key`. */
  set(key: string, value: string): void;
  /** Remove `key`; a missing key is not an error. */
  delete(key: string): void;
  /** Every key that currently has a value, enumerated from the backend itself (no index is trusted). */
  list(): string[];
  /** Whether the backend can be used at all on this machine. */
  available(): boolean;
}

/** The key of the MSAL token cache. */
export const TOKEN_CACHE_KEY = 'token-cache';

/**
 * ID-19: a credential-bearing file that is group/world-accessible is set back to 0600, with one stderr line.
 * POSIX only (on Windows the mode bits are synthetic and the file inherits the profile's ACL). Missing file: no-op.
 */
export function ensurePrivateMode(file: string, warn: (line: string) => void = l => { process.stderr.write(l + '\n'); }): void {
  if (process.platform === 'win32') return;
  let mode: number;
  try { mode = fs.statSync(file).mode; } catch { return; }
  if ((mode & 0o077) === 0) return;
  try {
    fs.chmodSync(file, 0o600);
    warn(`ms365: ${file} was accessible to other users (mode ${(mode & 0o777).toString(8).padStart(4, '0')}); changed it to 0600`);
  } catch { /* not ours to change; reading still works */ }
}

export const KEY_RE = /^[A-Za-z0-9][A-Za-z0-9_-]*(\.[A-Za-z0-9_-]+)*$/;

function assertKey(key: string, prefix: string): void {
  if (!KEY_RE.test(key) || (key !== prefix && !key.startsWith(`${prefix}.`))) {
    throw new Error(`Invalid credential store key "${key}"`);
  }
}

export interface FileStoreDeps {
  /** Directory holding the files; resolved on every call so env changes (tests, profiles) are honoured. */
  dir: () => string;
  /** Atomic 0600 replace under the cross-process lock (auth.ts `writeFileAtomic`). */
  writeFileAtomic: (target: string, data: string) => void;
  /** Cross-process lock next to a target (auth.ts `withFileLock`). */
  withFileLock: <T>(target: string, fn: () => T) => T;
  /** Only keys equal to this or starting with `<prefix>.` are accepted and listed (default `token-cache`). */
  prefix?: string;
}

/**
 * The `file` backend: today's behaviour. Key `k` is the file `<dir>/<k>.json`, written 0600 via temp file + rename
 * under an O_EXCL lock; so `token-cache` is exactly `configDir()/token-cache.json`.
 */
export function createFileStore(deps: FileStoreDeps): CredentialStore {
  const prefix = deps.prefix ?? TOKEN_CACHE_KEY;
  const file = (key: string) => {
    assertKey(key, prefix);
    return path.join(deps.dir(), `${key}.json`);
  };
  return {
    kind: 'file',
    get(key) {
      const target = file(key);
      ensurePrivateMode(target);
      try {
        return fs.readFileSync(target, 'utf8') || undefined;
      } catch {
        return undefined;
      }
    },
    set(key, value) {
      deps.writeFileAtomic(file(key), value);
    },
    delete(key) {
      const target = file(key);
      if (!fs.existsSync(path.dirname(target))) return;
      deps.withFileLock(target, () => {
        try { fs.unlinkSync(target); } catch (err) {
          if ((err as NodeJS.ErrnoException).code !== 'ENOENT') throw err;
        }
      });
    },
    list() {
      let names: string[];
      try { names = fs.readdirSync(deps.dir()); } catch { return []; }
      return names
        .filter(n => n.endsWith('.json'))
        .map(n => n.slice(0, -'.json'.length))
        .filter(k => KEY_RE.test(k) && (k === prefix || k.startsWith(`${prefix}.`)))
        .filter(k => {
          try { return fs.statSync(path.join(deps.dir(), `${k}.json`)).size > 0; } catch { return false; }
        })
        .sort();
    },
    available() {
      return true;
    },
  };
}

/** A store whose reads are cached in memory (ID-12). */
export interface CachedCredentialStore extends CredentialStore {
  /** Drop the cached values so the next `get` reads the backend again (e.g. after another process re-logged in). */
  invalidate(key?: string): void;
}

/**
 * Wrap a store so each key is read from the backend at most once per wrapper (one wrapper per process), and the
 * cache is changed only by this wrapper's own `set`/`delete` (ID-12). MSAL calls `beforeCacheAccess` 2–3 times per
 * command; on a keychain backend each read is a child process, so this bounds it to one.
 */
export function cachedStore(inner: CredentialStore): CachedCredentialStore {
  const cache = new Map<string, string | undefined>();
  return {
    get kind() { return inner.kind; },
    get(key) {
      if (!cache.has(key)) cache.set(key, inner.get(key));
      return cache.get(key);
    },
    set(key, value) {
      inner.set(key, value);
      cache.set(key, value || undefined);
    },
    delete(key) {
      inner.delete(key);
      cache.set(key, undefined);
    },
    list() {
      return inner.list();
    },
    available() {
      return inner.available();
    },
    invalidate(key) {
      if (key === undefined) cache.clear();
      else cache.delete(key);
    },
  };
}

// ── Per-account split of the MSAL cache (ID-6, ID-18) ───────────────

/**
 * Thrown by a backend's `get` when the item exists but cannot be decoded (e.g. a DPAPI blob that no longer
 * decrypts). The split cache deletes such an item and treats it as missing (ID-18).
 */
export class CorruptItemError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'CorruptItemError';
  }
}

type Entries = Record<string, Record<string, unknown>>;
/** A parsed MSAL cache: `{Account, IdToken, AccessToken, RefreshToken, AppMetadata, …}`, each a map of entries. */
export type MsalCacheJson = Record<string, Entries>;

/** Sections kept in the index (`token-cache.json`): no secrets. Everything else is split by `home_account_id`. */
const INDEX_SECTIONS = new Set(['Account', 'AppMetadata']);

/** JSON with object keys sorted, so equal caches give equal strings (MSAL's key order is not promised). */
export function stableStringify(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(',')}]`;
  if (value && typeof value === 'object') {
    const obj = value as Record<string, unknown>;
    return `{${Object.keys(obj).sort().filter(k => obj[k] !== undefined)
      .map(k => `${JSON.stringify(k)}:${stableStringify(obj[k])}`).join(',')}}`;
  }
  return JSON.stringify(value);
}

const isObj = (v: unknown): v is Record<string, unknown> => !!v && typeof v === 'object' && !Array.isArray(v);

/** Parse a serialized MSAL cache; anything that isn't `{section: {key: entry}}` yields null. */
export function parseMsalCache(data: string | undefined): MsalCacheJson | null {
  if (!data) return null;
  let parsed: unknown;
  try { parsed = JSON.parse(data); } catch { return null; }
  if (!isObj(parsed)) return null;
  for (const section of Object.values(parsed)) {
    if (!isObj(section) || !Object.values(section).every(isObj)) return null;
  }
  return parsed as MsalCacheJson;
}

export interface SplitCache {
  /** `Account` and `AppMetadata` (plus any secret-free entry without `home_account_id`). */
  index: MsalCacheJson;
  /** homeAccountId → that account's credential sections (`IdToken`, `AccessToken`, `RefreshToken`, …). */
  items: Map<string, MsalCacheJson>;
}

/**
 * Partition a cache by `home_account_id` (ID-6). Lossless for MSAL's shape: every Account/IdToken/AccessToken/
 * RefreshToken entry carries `home_account_id`; AppMetadata carries none and stays in the index.
 */
export function splitMsalCache(cache: MsalCacheJson): SplitCache {
  const index: MsalCacheJson = { Account: {}, AppMetadata: {} };
  const items = new Map<string, MsalCacheJson>();
  for (const [section, entries] of Object.entries(cache)) {
    if (INDEX_SECTIONS.has(section)) { index[section] = { ...entries }; continue; }
    for (const [key, entry] of Object.entries(entries)) {
      const hid = entry.home_account_id;
      if (typeof hid !== 'string' || !hid) {
        (index[section] ??= {})[key] = entry;
        continue;
      }
      let item = items.get(hid);
      if (!item) items.set(hid, item = {});
      (item[section] ??= {})[key] = entry;
    }
  }
  return { index, items };
}

/** The inverse of splitMsalCache: one object to hand to a single `deserialize()` (it replaces the whole cache). */
export function reassembleMsalCache(index: MsalCacheJson, items: Iterable<MsalCacheJson>): MsalCacheJson {
  const out: MsalCacheJson = { Account: {}, IdToken: {}, AccessToken: {}, RefreshToken: {}, AppMetadata: {} };
  for (const part of [index, ...items]) {
    for (const [section, entries] of Object.entries(part)) out[section] = { ...(out[section] ?? {}), ...entries };
  }
  return out;
}

/** Account-index homeAccountIds, in order of appearance. */
function accountIds(index: MsalCacheJson): string[] {
  const ids = new Set<string>();
  for (const e of Object.values(index.Account ?? {})) if (typeof e.home_account_id === 'string') ids.add(e.home_account_id);
  return [...ids];
}

function accountEntriesOf(index: MsalCacheJson, hid: string): Entries {
  return Object.fromEntries(Object.entries(index.Account ?? {}).filter(([, e]) => e.home_account_id === hid));
}

/** A valid item: credential sections only, every entry belonging to `hid`. */
function parseItem(data: string, hid: string): MsalCacheJson | null {
  const item = parseMsalCache(data);
  if (!item) return null;
  for (const [section, entries] of Object.entries(item)) {
    if (INDEX_SECTIONS.has(section)) return null;
    if (!Object.values(entries).every(e => e.home_account_id === hid)) return null;
  }
  return item;
}

const ITEM_SAFE = /^[A-Za-z0-9_-]+(\.[A-Za-z0-9_-]+)*$/;

/** The store key of an account's item: `token-cache.<homeAccountId>` (base64url-encoded if it has odd characters). */
export function itemKey(hid: string): string {
  return `${TOKEN_CACHE_KEY}.${ITEM_SAFE.test(hid) ? hid : `x-${Buffer.from(hid).toString('base64url')}`}`;
}

export interface SplitTokenCacheDeps {
  /** The index file (`token-cache.json`): raw reads and **unlocked** atomic writes (the caller holds `lock`). */
  index: { read(): string | undefined; write(data: string): void; remove(): void };
  /** The per-account item store (a keychain backend; the file store in tests). Reads must not be cached. */
  items: CredentialStore;
  /** The cross-process lock that serializes every read-modify-write of the index and items. */
  lock<T>(fn: () => T): T;
}

/**
 * The MSAL cache stored as one item per account plus a secret-free index (ID-6). `load()` reassembles everything
 * (dropping accounts whose item is missing or corrupt, ID-18) and is read once per instance (ID-12); `save()` is a
 * read-modify-write under the lock that writes only accounts whose JSON changed, deletes items of accounts this
 * process removed, and keeps accounts other processes added meanwhile.
 */
export class SplitTokenCache {
  /** homeAccountId → fingerprint (item JSON + its Account entries) as last loaded or saved by this process. */
  private baseline = new Map<string, string>();
  /** Accounts this process has seen (in the index at load time, or saved by it). */
  private known = new Set<string>();
  /** Index accounts that had no usable item at load time (dropped from the reassembled cache). */
  private missing = new Set<string>();
  private current: string | undefined;
  private loaded = false;

  constructor(private readonly deps: SplitTokenCacheDeps) {}

  /** The reassembled cache as a JSON string, or undefined when nothing is stored. */
  load(): string | undefined {
    if (this.loaded) return this.current;
    this.loaded = true;
    const index = parseMsalCache(this.deps.index.read());
    if (!index) return this.current = undefined;
    // Credential sections left in the index file (a pre-split file cache) count as that account's item.
    const legacy = splitMsalCache(index);
    const items: MsalCacheJson[] = [];
    for (const hid of accountIds(index)) {
      this.known.add(hid);
      const item = this.readItem(hid) ?? legacy.items.get(hid);
      if (!item) { this.missing.add(hid); continue; }
      items.push(item);
      if (legacy.items.has(hid)) continue; // not yet in the item store: the next save writes it
      this.baseline.set(hid, fingerprint(item, accountEntriesOf(index, hid)));
    }
    const keptIndex: MsalCacheJson = { ...legacy.index, Account: Object.fromEntries(
      Object.entries(index.Account ?? {}).filter(([, e]) => !this.missing.has(e.home_account_id as string))) };
    return this.current = stableStringify(reassembleMsalCache(keptIndex, items));
  }

  /** Treat `serialized` (e.g. a file cache being moved into the item store) as loaded, with nothing yet stored. */
  adopt(serialized: string | undefined): void {
    this.loaded = true;
    this.current = serialized;
    const cache = parseMsalCache(serialized);
    if (cache) for (const hid of accountIds(cache)) this.known.add(hid);
  }

  /** Forget what was loaded so the next `load()` reads the stores again (IDT-5). */
  invalidate(): void {
    this.baseline.clear();
    this.known.clear();
    this.missing.clear();
    this.current = undefined;
    this.loaded = false;
  }

  save(serialized: string): void {
    const cache = parseMsalCache(serialized);
    if (!cache) throw new Error('MSAL produced a token cache that is not JSON');
    const { index: mine, items: myItems } = splitMsalCache(cache);
    const myIds = new Set([...accountIds(mine), ...myItems.keys()]);

    this.deps.lock(() => {
      const disk = parseMsalCache(this.deps.index.read()) ?? { Account: {}, AppMetadata: {} };
      const diskSplit = splitMsalCache(disk);

      // Accounts this process removed (or found without an item): delete their item unless another process has
      // since written one (a fresh, uncached read under the lock).
      const dropped = new Set<string>();
      for (const hid of this.known) {
        if (myIds.has(hid)) continue;
        if (this.missing.has(hid) && !this.baseline.has(hid)) {
          if (this.readItem(hid)) continue; // re-logged in elsewhere meanwhile: keep it
        } else {
          this.deps.items.delete(itemKey(hid));
        }
        dropped.add(hid);
      }

      // Write only what changed.
      const next = new Map<string, string>();
      for (const hid of myIds) {
        const item = myItems.get(hid) ?? {};
        const fp = fingerprint(item, accountEntriesOf(mine, hid));
        next.set(hid, fp);
        if (this.baseline.get(hid) !== fp) this.deps.items.set(itemKey(hid), stableStringify(item));
      }
      // Credentials still sitting in the index file for accounts this process never handled (a pre-split file):
      // move them into the item store instead of dropping them.
      for (const [hid, item] of diskSplit.items) {
        if (myIds.has(hid) || dropped.has(hid) || this.readItem(hid)) continue;
        this.deps.items.set(itemKey(hid), stableStringify(item));
      }

      // Index: the disk's entries, minus accounts this process dropped or rewrote, plus this process's view.
      const index: MsalCacheJson = {};
      for (const [section, entries] of Object.entries(diskSplit.index)) index[section] = { ...entries };
      index.Account = Object.fromEntries(Object.entries(disk.Account ?? {}).filter(([, e]) => {
        const hid = e.home_account_id as string;
        return !dropped.has(hid) && !(myIds.has(hid) && this.baseline.get(hid) !== next.get(hid));
      }));
      for (const [section, entries] of Object.entries(mine)) {
        if (section === 'Account') {
          for (const [k, e] of Object.entries(entries)) {
            const hid = e.home_account_id as string;
            if (this.baseline.get(hid) !== next.get(hid) || !Object.values(index.Account).some(x => x.home_account_id === hid)) index.Account[k] = e;
          }
        } else {
          index[section] = { ...(index[section] ?? {}), ...entries };
        }
      }
      this.deps.index.write(stableStringify(index));

      this.baseline = next;
      this.known = new Set(myIds);
      this.missing.clear();
    });
    this.current = serialized;
    this.loaded = true;
  }

  /** An account's item, or undefined when missing; a corrupt item is deleted and treated as missing (ID-18). */
  private readItem(hid: string): MsalCacheJson | undefined {
    const key = itemKey(hid);
    let data: string | undefined;
    try {
      data = this.deps.items.get(key);
    } catch (err) {
      if (!(err instanceof CorruptItemError)) throw err;
      this.deps.items.delete(key);
      return undefined;
    }
    if (data === undefined) return undefined;
    const item = parseItem(data, hid);
    if (!item) {
      this.deps.items.delete(key);
      return undefined;
    }
    return item;
  }
}

function fingerprint(item: MsalCacheJson, accounts: Entries): string {
  return stableStringify({ item, accounts });
}

// ── Migration between the file layout and per-account items (ID-5) ──

export interface MigrationDeps {
  /** The index file (`token-cache.json`): raw reads and **unlocked** atomic writes (the caller's `lock` is held). */
  index: { read(): string | undefined; write(data: string): void };
  /** The per-account item store (the keychain backend). */
  items: CredentialStore;
  /** The cross-process lock held for the whole migration (the same one SplitTokenCache.save takes). */
  lock<T>(fn: () => T): T;
  /** Make the destination the configured store (storage.json); `rollback` undoes it. */
  commit(): void;
  rollback(): void;
}

export interface MigrationResult {
  /** homeAccountIds now held by the destination. */
  accounts: string[];
  /** Accounts left out because no credentials for them were found (they could not sign in silently anyway). */
  dropped: string[];
}

/** Thrown when a destination does not read back what was written; the source is left untouched. */
export class MigrationVerifyError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'MigrationVerifyError';
  }
}

/**
 * Move a file cache's credentials into per-account items (`auth storage use keychain`). Under the lock:
 * 1. commit the setting first — from here on a reader takes credentials from an item, or else from the index file
 *    (SplitTokenCache's legacy path), so a crash at any later point leaves every account usable;
 * 2. write one item per account and read each back, comparing it to the source (stable JSON);
 * 3. only then rewrite the index file without its credential sections.
 * A failure in 2 deletes the items written, rolls the setting back and leaves the file as it was.
 * Re-running it after an interruption completes the move (idempotent).
 */
export function migrateFileToItems(deps: MigrationDeps): MigrationResult {
  return deps.lock(() => {
    const disk = parseMsalCache(deps.index.read());
    deps.commit();
    if (!disk) return { accounts: [], dropped: [] };
    const { index, items: legacy } = splitMsalCache(disk);
    const written: string[] = [];
    try {
      for (const [hid, item] of legacy) {
        deps.items.set(itemKey(hid), stableStringify(item));
        written.push(hid);
      }
      for (const [hid, item] of legacy) {
        let back: string | undefined;
        try { back = deps.items.get(itemKey(hid)); } catch (err) {
          if (!(err instanceof CorruptItemError)) throw err;
        }
        const parsed = back === undefined ? null : parseItem(back, hid);
        if (!parsed || stableStringify(parsed) !== stableStringify(item)) {
          throw new MigrationVerifyError(`the keychain item for account ${hid} did not read back as written`);
        }
      }
    } catch (err) {
      for (const hid of written) { try { deps.items.delete(itemKey(hid)); } catch { /* best effort */ } }
      deps.rollback();
      throw err;
    }
    if (legacy.size > 0) deps.index.write(stableStringify(index));
    const ids = accountIds(disk);
    const have = new Set([...legacy.keys()]);
    // Accounts already moved by an earlier (interrupted) run keep their items.
    for (const hid of ids) {
      if (have.has(hid)) continue;
      try { if (deps.items.get(itemKey(hid)) !== undefined) have.add(hid); } catch { /* corrupt: treated as missing */ }
    }
    return { accounts: ids.filter(h => have.has(h)), dropped: ids.filter(h => !have.has(h)) };
  });
}

/**
 * Move per-account items back into the single 0600 file (`auth storage use file`). Under the lock:
 * 1. reassemble index + items (an item still missing falls back to credentials left in the index file);
 * 2. write the full cache to the index file and read it back, comparing (stable JSON);
 * 3. delete every item of this service (`list()`, not the index — orphans go too);
 * 4. commit the setting last — until then readers in keychain mode find the credentials in the file (legacy path).
 * A crash at any point leaves every account usable; re-running completes the move.
 */
export function migrateItemsToFile(deps: MigrationDeps): MigrationResult {
  return deps.lock(() => {
    const disk = parseMsalCache(deps.index.read());
    if (!disk) {
      for (const key of deps.items.list()) deps.items.delete(key);
      deps.commit();
      return { accounts: [], dropped: [] };
    }
    const { index, items: legacy } = splitMsalCache(disk);
    const parts: MsalCacheJson[] = [];
    const accounts: string[] = [];
    const dropped: string[] = [];
    for (const hid of accountIds(disk)) {
      let data: string | undefined;
      try { data = deps.items.get(itemKey(hid)); } catch (err) {
        if (!(err instanceof CorruptItemError)) throw err;
      }
      const item = (data === undefined ? null : parseItem(data, hid)) ?? legacy.get(hid);
      if (item) { parts.push(item); accounts.push(hid); } else dropped.push(hid);
    }
    const keptIndex: MsalCacheJson = { ...index, Account: Object.fromEntries(
      Object.entries(index.Account ?? {}).filter(([, e]) => !dropped.includes(e.home_account_id as string))) };
    const full = stableStringify(reassembleMsalCache(keptIndex, parts));
    deps.index.write(full);
    const back = parseMsalCache(deps.index.read());
    if (!back || stableStringify(back) !== full) {
      // The items are untouched and still authoritative; put the index back as it was.
      deps.index.write(stableStringify(disk));
      throw new MigrationVerifyError('token-cache.json did not read back as written');
    }
    for (const key of deps.items.list()) deps.items.delete(key);
    deps.commit();
    return { accounts, dropped };
  });
}

/** homeAccountIds in a cache's `Account` section, and how many of them have credentials in it (ID-15). */
export function cacheAccounts(data: string | undefined): { ids: string[]; withCredentials: string[] } {
  const cache = parseMsalCache(data);
  if (!cache) return { ids: [], withCredentials: [] };
  const { items } = splitMsalCache(cache);
  const ids = accountIds(cache);
  return { ids, withCredentials: ids.filter(h => items.has(h)) };
}
