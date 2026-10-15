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

const KEY_RE = /^[A-Za-z0-9][A-Za-z0-9_-]*(\.[A-Za-z0-9_-]+)*$/;

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
