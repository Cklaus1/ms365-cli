import type { AccountInfo, Configuration, ICachePlugin, PublicClientApplication, TokenCacheContext } from '@azure/msal-node';
import fs from 'fs';
import path from 'path';
import os from 'os';
import crypto from 'crypto';
import { CliError, ExitCode } from './errors.js';
import { checkExternalToken, requireExternalToken, tokenRejectedError } from './external-token.js';
import {
  cacheAccounts, cachedStore, createFileStore, ensurePrivateMode, itemKey, migrateFileToItems, migrateItemsToFile,
  MigrationVerifyError, SplitTokenCache, TOKEN_CACHE_KEY,
  type CachedCredentialStore, type CredentialStore, type MigrationDeps, type MigrationResult,
} from './credential-store.js';
import {
  AUTO_PREFERS_KEYCHAIN, BACKEND_LABEL, configuredStorage, keychainTimeoutMs, keychainUnavailable, platformKeychain,
  probeKeychain, PROBE_TIMEOUT_MS, readStorageConfig, resolveStorage, safeBackend, storageConfigPath, writeStorageConfig,
  type KeychainStore, type ResolveOptions, type StorageConfig, type StorageSetting,
} from './keychain.js';

/** Silent token requests use .default: the token carries whatever the user consented to at login. */
export const GRAPH_DEFAULT_SCOPE = 'https://graph.microsoft.com/.default';

/** Treat access tokens as expired this long before they actually expire. */
const EXPIRY_SKEW_MS = 5 * 60 * 1000;

export function configDir(): string {
  return process.env.MS365_CLI_CONFIG_DIR || path.join(os.homedir(), '.config', 'ms365-cli');
}
const selectedAccountPath = () => path.join(configDir(), 'selected-account.json');
const deviceCodePath = () => path.join(configDir(), 'device-code.txt');

const SCOPE_HIERARCHY: Record<string, string[]> = {
  'Mail.ReadWrite': ['Mail.Read'],
  'Calendars.ReadWrite': ['Calendars.Read'],
  'Files.ReadWrite': ['Files.Read'],
  'Tasks.ReadWrite': ['Tasks.Read'],
  'Contacts.ReadWrite': ['Contacts.Read'],
  'Presence.ReadWrite': ['Presence.Read'],
};

const PERSONAL_SCOPES = [
  'User.Read',
  'Mail.ReadWrite',
  'Mail.Send',
  'Calendars.ReadWrite',
  'Files.ReadWrite',
  'Tasks.ReadWrite',
  'Contacts.ReadWrite',
  'Notes.Read',
  'Notes.Create',
  'People.Read',
  'Presence.Read',
  'Presence.ReadWrite',
  'MailboxSettings.ReadWrite',
  'User.ReadWrite',
  'Files.Read.All',
  'ShortNotes.ReadWrite',
];

const ORG_SCOPES = [
  'Chat.Read',
  'ChatMessage.Read',
  'ChatMessage.Send',
  'Team.ReadBasic.All',
  'Channel.ReadBasic.All',
  'ChannelMessage.Read.All',
  'ChannelMessage.Send',
  'TeamMember.Read.All',
  'Sites.Read.All',
  'Mail.Read.Shared',
  'Mail.Send.Shared',
  'User.Read.All',
  'Group.Read.All',
  'Calendars.Read.Shared',
  'Place.Read.All',
  'OnlineMeetings.Read',
  'OnlineMeetingTranscript.Read.All',
  'OnlineMeetingRecording.Read.All',
  'Reports.Read.All',
  // audit-trail (AO-15): per-workload audit search, directory audits, app grants (Directory.Read.All: AO-28, OQ-9)
  'AuditLogsQuery-Exchange.Read.All',
  'AuditLog.Read.All',
  'Application.Read.All',
  'Directory.Read.All',
];

function ensureConfigDir(): void {
  fs.mkdirSync(configDir(), { recursive: true, mode: 0o700 });
}

export const configFilePath = () => path.join(configDir(), 'config.json');

/** Scopes saved by `ms365 auth setup` in config.json, or null when none are saved (or the file is unreadable). */
export function readSavedScopes(): string[] | null {
  const data = readFileIfExists(configFilePath());
  if (!data) return null;
  try {
    const scopes = JSON.parse(data).scopes;
    return Array.isArray(scopes) && scopes.length > 0 && scopes.every(s => typeof s === 'string') ? scopes : null;
  } catch {
    return null;
  }
}

/** Scopes requested at login: the wizard's saved choice (D14), else the built-in personal (+ org) set. */
export function buildScopes(orgMode: boolean): string[] {
  const scopes = readSavedScopes() ?? [...PERSONAL_SCOPES, ...(orgMode ? ORG_SCOPES : [])];

  // Optimize scope hierarchy
  const scopeSet = new Set(scopes);
  for (const [higher, lowers] of Object.entries(SCOPE_HIERARCHY)) {
    if (scopeSet.has(higher)) {
      for (const lower of lowers) {
        scopeSet.delete(lower);
      }
    }
  }

  return Array.from(scopeSet);
}

// ── File persistence with cross-process locking ─────────────────────

const LOCK_TIMEOUT_MS = 10_000;
const LOCK_STALE_MS = 30_000;

function sleepSync(ms: number): void {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}

/** Run fn while holding an O_EXCL lock file next to `target`. Stale locks are broken. */
export function withFileLock<T>(target: string, fn: () => T): T {
  const lock = `${target}.lock`;
  const deadline = Date.now() + LOCK_TIMEOUT_MS;
  let fd: number | null = null;
  while (fd === null) {
    try {
      fd = fs.openSync(lock, 'wx', 0o600);
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code !== 'EEXIST') throw err;
      try {
        if (Date.now() - fs.statSync(lock).mtimeMs > LOCK_STALE_MS) fs.unlinkSync(lock);
      } catch { /* lock released meanwhile */ }
      if (Date.now() > deadline) throw new CliError(`Timed out waiting for lock ${lock}`, { code: 'lock_timeout' });
      sleepSync(10 + Math.floor(Math.random() * 20));
    }
  }
  try {
    return fn();
  } finally {
    fs.closeSync(fd);
    try { fs.unlinkSync(lock); } catch { /* already gone */ }
  }
}

/** Replace `target` atomically (unique 0600 temp file + rename). The caller holds the lock. */
function replaceFileAtomic(target: string, data: string): void {
  const tmp = `${target}.${process.pid}.${crypto.randomBytes(4).toString('hex')}.tmp`;
  try {
    fs.writeFileSync(tmp, data, { mode: 0o600 });
    fs.renameSync(tmp, target);
  } catch (err) {
    try { fs.unlinkSync(tmp); } catch { /* not created */ }
    throw err;
  }
}

/** Atomically replace `target` (unique temp file + rename) under a cross-process lock. */
export function writeFileAtomic(target: string, data: string): void {
  ensureConfigDir();
  withFileLock(target, () => replaceFileAtomic(target, data));
}

function readFileIfExists(p: string): string | undefined {
  try {
    const data = fs.readFileSync(p, 'utf8');
    return data || undefined;
  } catch {
    return undefined;
  }
}

/** Today's credential store: the 0600 `configDir()/token-cache.json` (ID-2 `file` backend). */
export function fileCredentialStore(): CredentialStore {
  return createFileStore({ dir: configDir, writeFileAtomic, withFileLock });
}

/** The split cache over `items`, with `configDir()/token-cache.json` as its index (ID-6). */
export function splitTokenCache(items: CredentialStore): SplitTokenCache {
  const indexFile = () => path.join(configDir(), `${TOKEN_CACHE_KEY}.json`);
  return new SplitTokenCache({
    index: {
      read: () => { ensurePrivateMode(indexFile()); return readFileIfExists(indexFile()); },
      write: data => { ensureConfigDir(); replaceFileAtomic(indexFile(), data); },
      remove: () => { try { fs.unlinkSync(indexFile()); } catch { /* none */ } },
    },
    items,
    lock: fn => { ensureConfigDir(); return withFileLock(indexFile(), fn); },
  });
}

/**
 * Where the cache's credentials go for a given phase: null = the whole cache in `store` (the 0600 file), else the
 * per-account item store (a keychain backend). Called lazily, so `--help`/dry-run never resolve storage.
 */
export type ItemStoreResolver = (phase: 'read' | 'write') => CredentialStore | null;

/** The production resolver (ID-4): MS365_CLI_STORAGE / storage.json / auto (E1: auto = file). */
export function storageResolver(overrides: Partial<ResolveOptions> = {}): ItemStoreResolver {
  return phase => {
    const r = resolveStorage({ configDir: configDir(), phase, writeFileAtomic, ...overrides });
    return r.kind === 'keychain' ? r.store ?? null : null;
  };
}

// ── Switching storage and reporting it (ID-5, ID-15) ───────────────

const indexFilePath = () => path.join(configDir(), `${TOKEN_CACHE_KEY}.json`);

/** storage.json without what `auto` resolved (logout, explicit switches). Keeps `storage` and `noticeShown`. */
function withoutResolution(cfg: StorageConfig): StorageConfig {
  const { resolved: _r, backend: _b, probedAt: _p, reason: _x, ...rest } = cfg;
  return rest;
}

export interface StorageOptions {
  env?: NodeJS.ProcessEnv;
  platform?: NodeJS.Platform;
  /** Default AUTO_PREFERS_KEYCHAIN (E1: false). */
  autoPrefersKeychain?: boolean;
  /** Plan only (`--dry-run`): never probe, since the probe writes an item. */
  planOnly?: boolean;
}

export interface StoragePlan {
  target: StorageSetting;
  /** Where the credentials are now, and where they go. */
  from: 'file' | 'keychain';
  to: 'file' | 'keychain';
  /** Accounts in the index (token-cache.json). */
  accounts: number;
  /** The keychain item store involved (from or to), when either side is the keychain. */
  store?: KeychainStore;
  /** The storage.json to write once the destination is verified. */
  next: StorageConfig;
}

/** Work out what `auth storage use <target>` would do, touching nothing (dry-run safe: no tool is spawned). */
export function planStorageSwitch(target: StorageSetting, opts: StorageOptions = {}): StoragePlan {
  const env = opts.env ?? process.env;
  const dir = configDir();
  const configured = configuredStorage(dir, env);
  if (configured.source === 'env' && configured.setting !== target) {
    throw new CliError(`MS365_CLI_STORAGE=${configured.setting} is set and overrides storage.json, so switching to "${target}" would not take effect.`,
      { code: 'usage', exitCode: ExitCode.USAGE, fix: `unset MS365_CLI_STORAGE (or set it to ${target}) and retry` });
  }
  const cfg = readStorageConfig(dir);
  // Where the credentials are now. An unusable keychain is still "keychain": its items cannot be read back.
  let from: 'file' | 'keychain';
  let fromStore: KeychainStore | undefined;
  try {
    const r = resolveStorage({ configDir: dir, phase: 'read', env, platform: opts.platform, writeFileAtomic, autoPrefersKeychain: opts.autoPrefersKeychain });
    from = r.kind;
    fromStore = r.store;
  } catch (err) {
    if (!(err instanceof CliError && err.code === 'keychain_unavailable')) throw err;
    from = 'keychain';
  }

  let to: 'file' | 'keychain';
  let toStore: KeychainStore | undefined;
  let next: StorageConfig = { ...withoutResolution(cfg), storage: target };
  if (target === 'file') {
    to = 'file';
  } else if (target === 'keychain') {
    const kc = platformKeychain({ configDir: dir, env, platform: opts.platform, writeFileAtomic });
    if (typeof kc === 'string') throw keychainUnavailable(kc, safeBackend(env, opts.platform));
    to = 'keychain';
    toStore = kc;
  } else {
    // auto: E1 makes it the file store (nothing recorded); with the probe enabled it re-probes now (ID-4).
    to = 'file';
    if ((opts.autoPrefersKeychain ?? AUTO_PREFERS_KEYCHAIN) && opts.planOnly) {
      next = { ...next, reason: 'not probed (--dry-run)' };
    } else if (opts.autoPrefersKeychain ?? AUTO_PREFERS_KEYCHAIN) {
      const kc = platformKeychain({ configDir: dir, env, platform: opts.platform, writeFileAtomic, timeoutMs: Math.min(PROBE_TIMEOUT_MS, keychainTimeoutMs(env)) });
      const probed = typeof kc === 'string' ? { ok: false, reason: kc } : probeKeychain(kc);
      const full = probed.ok ? platformKeychain({ configDir: dir, env, platform: opts.platform, writeFileAtomic }) : probed.reason;
      next = { ...next, resolved: 'file', probedAt: new Date().toISOString(), reason: probed.reason };
      if (typeof full !== 'string') {
        to = 'keychain';
        toStore = full;
        next = { ...next, resolved: 'keychain', backend: full.backend };
      }
    }
  }
  return { target, from, to, accounts: cacheAccounts(readFileIfExists(indexFilePath())).ids.length, store: toStore ?? fromStore, next };
}

export interface StorageSwitchResult extends MigrationResult {
  target: StorageSetting;
  from: 'file' | 'keychain';
  to: 'file' | 'keychain';
  backend: string;
}

/**
 * `auth storage use <target>` (ID-5): move every cached account between the 0600 file and the keychain with no
 * re-login, holding the token-cache lock throughout; the source is removed only after the destination reads back
 * what was written, and an interrupted run leaves a usable store (see migrateFileToItems / migrateItemsToFile).
 */
export function switchStorage(target: StorageSetting, opts: StorageOptions = {}): StorageSwitchResult {
  const plan = planStorageSwitch(target, opts);
  const env = opts.env ?? process.env;
  const dir = configDir();
  const before = readStorageConfig(dir);
  const label = plan.to === 'keychain' ? BACKEND_LABEL[plan.store!.backend] : 'file';
  const done = (r: MigrationResult): StorageSwitchResult => ({ ...r, target, from: plan.from, to: plan.to, backend: label });
  const commit = () => writeStorageConfig(dir, plan.next, writeFileAtomic);
  const rollback = () => {
    if (fs.existsSync(storageConfigPath(dir)) || Object.keys(before).length > 0) writeStorageConfig(dir, before, writeFileAtomic);
  };

  if (plan.from === 'file' && plan.to === 'file') {
    commit();
    const ids = cacheAccounts(readFileIfExists(indexFilePath())).ids;
    return done({ accounts: ids, dropped: [] });
  }
  if (plan.from === 'keychain' && !plan.store) {
    // The keychain holding the credentials cannot be used, so they cannot be moved.
    const kc = platformKeychain({ configDir: dir, env, platform: opts.platform, writeFileAtomic });
    throw new CliError(`Cannot read the keychain that holds the credentials (${typeof kc === 'string' ? kc : 'unavailable'}), so they cannot be moved to the ${plan.to === 'file' ? 'file' : 'new store'}.`, {
      code: 'keychain_unavailable', exitCode: ExitCode.AUTH,
      fix: 'make the keychain available (on macOS: security unlock-keychain) and retry, or sign in again after: ms365 auth logout',
    });
  }
  const deps: MigrationDeps = {
    index: {
      read: () => readFileIfExists(indexFilePath()),
      write: data => { ensureConfigDir(); replaceFileAtomic(indexFilePath(), data); },
    },
    items: plan.store!,
    lock: fn => { ensureConfigDir(); return withFileLock(indexFilePath(), fn); },
    commit,
    rollback,
  };
  try {
    // keychain → keychain re-runs the import, which also completes an interrupted move into the keychain.
    return done(plan.to === 'keychain' ? migrateFileToItems(deps) : migrateItemsToFile(deps));
  } catch (err) {
    if (err instanceof MigrationVerifyError) {
      throw new CliError(`Storage not switched: ${err.message}. Nothing was removed; the credentials stay where they were.`,
        { code: 'migration_failed', exitCode: ExitCode.GENERIC, fix: 'retry; if it keeps failing: ms365 auth storage show' });
    }
    throw err;
  }
}

export interface StorageStatus {
  configured: StorageSetting;
  /** Where `configured` came from: MS365_CLI_STORAGE, storage.json, or the default. */
  source: 'env' | 'file' | 'default';
  /** Where credentials are read from now (null: the configured keychain cannot be used). */
  resolved: 'keychain' | 'file' | null;
  backend: string;
  /** Live probe of this platform's keychain (null with --no-probe, --dry-run or --read-only). */
  available: boolean | null;
  probeMs: number | null;
  reason: string | null;
  /** Accounts in the index, and how many have credentials in the store in effect (an item, or file entries). */
  accounts: { count: number; withItem: number | null };
  files: Array<{ name: string; path: string; exists: boolean; mode: string | null }>;
  noticeShown: boolean;
}

function fileMode(p: string): { exists: boolean; mode: string | null } {
  try {
    const st = fs.statSync(p);
    return { exists: true, mode: process.platform === 'win32' ? 'inherited' : (st.mode & 0o777).toString(8).padStart(4, '0') };
  } catch {
    return { exists: false, mode: null };
  }
}

/** `auth storage show` (ID-15). Never throws for keychain trouble: the probe result is data. */
export function storageStatus(opts: StorageOptions & { probe?: boolean; noProbeReason?: string } = {}): StorageStatus {
  const env = opts.env ?? process.env;
  const dir = configDir();
  const { setting, source } = configuredStorage(dir, env);
  const cfg = readStorageConfig(dir);
  let resolved: StorageStatus['resolved'];
  let store: KeychainStore | undefined;
  let reason: string | null = null;
  try {
    const r = resolveStorage({ configDir: dir, phase: 'read', env, platform: opts.platform, writeFileAtomic, autoPrefersKeychain: opts.autoPrefersKeychain });
    resolved = r.kind;
    store = r.store;
    reason = r.reason ?? null;
  } catch (err) {
    if (!(err instanceof CliError && err.code === 'keychain_unavailable')) throw err;
    resolved = null;
    reason = err.message;
  }
  const backendOf = (b: string | undefined) => (b ? BACKEND_LABEL[b as keyof typeof BACKEND_LABEL] ?? b : 'keychain');
  const backend = resolved === 'file' ? 'file' : backendOf(store?.backend ?? safeBackend(env, opts.platform));

  let available: boolean | null = null;
  let probeMs: number | null = null;
  if (opts.probe !== false) {
    const started = Date.now();
    const kc = platformKeychain({ configDir: dir, env, platform: opts.platform, writeFileAtomic, timeoutMs: Math.min(PROBE_TIMEOUT_MS, keychainTimeoutMs(env)) });
    const p = typeof kc === 'string' ? { ok: false, reason: kc } : probeKeychain(kc);
    available = p.ok;
    probeMs = Date.now() - started;
    reason ??= p.reason;
    if (!p.ok && resolved !== 'keychain') reason = `${reason === p.reason ? '' : `${reason}; `}keychain: ${p.reason}`;
  } else if (opts.noProbeReason) {
    reason ??= opts.noProbeReason;
  }

  const { ids, withCredentials } = cacheAccounts(readFileIfExists(indexFilePath()));
  let withItem: number | null = withCredentials.length;
  if (store) {
    try {
      const keys = new Set(store.list());
      withItem = ids.filter(h => keys.has(itemKey(h))).length;
    } catch { withItem = null; }
  } else if (resolved === null) {
    withItem = null;
  }

  const files = [
    ['token-cache.json', indexFilePath()],
    ['storage.json', storageConfigPath(dir)],
    ['credentials/', path.join(dir, 'credentials')],
  ].map(([name, p]) => ({ name, path: p, ...fileMode(p) }));
  return {
    configured: setting, source, resolved, backend, available, probeMs, reason,
    accounts: { count: ids.length, withItem }, files, noticeShown: !!cfg.noticeShown,
  };
}

/**
 * ID-13: delete every keychain item of this config directory's service (enumerated with `list()`, not from the
 * index), the DPAPI `credentials/` directory, and storage.json's resolution (keeps `storage` and `noticeShown`).
 * The keychain is only consulted when it is the store in effect, so file-mode logouts never spawn a keychain tool.
 */
function removeStoredCredentials(): void {
  const dir = configDir();
  let firstErr: unknown;
  try {
    const r = resolveStorage({ configDir: dir, phase: 'read', writeFileAtomic });
    if (r.store) for (const key of r.store.list()) r.store.delete(key);
  } catch (err) { firstErr = err; }
  try { fs.rmSync(path.join(dir, 'credentials'), { recursive: true, force: true }); } catch (err) { firstErr ??= err; }
  if (fs.existsSync(storageConfigPath(dir))) {
    const cfg = readStorageConfig(dir);
    const next = withoutResolution(cfg);
    if (Object.keys(next).length !== Object.keys(cfg).length) writeStorageConfig(dir, next, writeFileAtomic);
  }
  if (firstErr) throw firstErr;
}

/**
 * MSAL cache plugin backed by a CredentialStore. In file mode (the default; E1) the whole cache is one 0600
 * `token-cache.json`, as before. In keychain mode the credentials are split into one item per account and the file
 * keeps only the secret-free `Account`/`AppMetadata` index (ID-6), reassembled before MSAL's single `deserialize()`.
 * Either way the stores are read once per plugin (one per process) and changed only by this plugin's writes
 * (ID-12). MSAL calls afterCacheAccess after every token operation and sets cacheHasChanged when a token was
 * fetched from the network, so silent refreshes (and rotated refresh tokens) are persisted without extra code paths.
 */
export class FileTokenCachePlugin implements ICachePlugin {
  /** The token-cache store: the whole cache in file mode. */
  readonly store: CachedCredentialStore;
  private readonly resolver: ItemStoreResolver;
  /** Set once keychain mode is in effect. */
  private split: SplitTokenCache | null = null;
  private readResolved = false;

  /**
   * @param store the file store (default `configDir()/token-cache.json`)
   * @param items the per-account item store, or a resolver deciding it lazily (default: file mode only)
   */
  constructor(store: CredentialStore = fileCredentialStore(), items?: CredentialStore | ItemStoreResolver) {
    this.store = cachedStore(store);
    this.resolver = typeof items === 'function' ? items : () => items ?? null;
  }

  /** Which layout is in effect (`split` once credentials live in per-account items). */
  mode(): 'file' | 'split' {
    return this.split ? 'split' : 'file';
  }

  /** Drop what was read so the next access reads the stores again (IDT-5). */
  invalidate(): void {
    this.store.invalidate();
    this.split?.invalidate();
  }

  async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
    if (!this.readResolved) {
      const items = this.resolver('read');
      if (items) this.split = splitTokenCache(items);
      this.readResolved = true;
    }
    const data = this.split ? this.split.load() : this.store.get(TOKEN_CACHE_KEY);
    if (data) {
      try {
        ctx.tokenCache.deserialize(data);
      } catch {
        // Corrupt cache file: start empty (next login rewrites it).
      }
    }
  }

  async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
    if (!ctx.cacheHasChanged) return;
    if (!this.split) {
      // `auto` resolves on the first write (ID-4); with E1 it stays file.
      const items = this.resolver('write');
      if (items) {
        this.split = splitTokenCache(items);
        this.split.adopt(this.readResolved ? this.store.get(TOKEN_CACHE_KEY) : undefined);
      }
    }
    if (this.split) this.split.save(ctx.tokenCache.serialize());
    else this.store.set(TOKEN_CACHE_KEY, ctx.tokenCache.serialize());
  }
}

/** Decode an access token's JWT payload (no signature check: the claims are only reported, never trusted). */
export function decodeAccessToken(token: string): Record<string, unknown> {
  const opaque = () => new CliError(
    'The access token is not a JWT (personal Microsoft account tokens are opaque), so its claims cannot be read locally.',
    { code: 'opaque_token' },
  );
  const parts = token.split('.');
  if (parts.length !== 3) throw opaque();
  try {
    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));
    if (payload && typeof payload === 'object' && !Array.isArray(payload)) return payload;
  } catch { /* fall through */ }
  throw opaque();
}

export interface TokenInfo {
  scopes: string[];
  audience?: string;
  tenantId?: string;
  user?: string;
  expiresOn?: string;
  expired?: boolean;
}

/** The consent-relevant claims of an access token (spec §6): granted scopes (`scp`), audience, tenant, user, expiry. */
export function tokenInfo(token: string, now = new Date()): TokenInfo {
  const claims = decodeAccessToken(token);
  const str = (v: unknown) => (typeof v === 'string' ? v : undefined);
  const info: TokenInfo = {
    scopes: (str(claims.scp) ?? '').split(' ').filter(Boolean),
    audience: str(claims.aud),
    tenantId: str(claims.tid),
    user: str(claims.upn) ?? str(claims.preferred_username),
  };
  if (typeof claims.exp === 'number') {
    info.expiresOn = new Date(claims.exp * 1000).toISOString();
    info.expired = claims.exp * 1000 <= now.getTime();
  }
  return info;
}

/**
 * Turn a sign-in failure into an actionable CliError. Covers the cases that otherwise surface
 * as raw MSAL/AADSTS text: admin approval required or requested, consent declined, and an
 * expired device code. `orgMode` decides whether the suggested retry includes --org-mode.
 */
export function explainLoginError(err: unknown, orgMode = false): CliError {
  const raw = err instanceof Error ? err.message : String(err);
  const retry = orgMode ? 'ms365 --org-mode auth login' : 'ms365 auth login';
  if (/AADSTS90094|AADSTS90095|AADSTS65001|admin.?(approval|consent)|consent_required|approval/i.test(raw)) {
    return new CliError(
      `Sign-in needs an admin to approve some of the requested permissions (${raw}). ` +
      'A pending approval request does not complete this login.',
      {
        code: 'admin_consent_required', exitCode: ExitCode.AUTH,
        fix: 'an admin approves it in Entra (Enterprise applications → Admin consent requests, or the app\'s API permissions → Grant admin consent); then run: ' + retry,
      },
    );
  }
  if (/AADSTS65004|declined|access_denied/i.test(raw)) {
    return new CliError(`Sign-in was declined (${raw}).`,
      { code: 'consent_declined', exitCode: ExitCode.AUTH, fix: `run ${retry} again and accept the permissions` });
  }
  if (/device_code_expired|expired_token|code_expired|AADSTS70020/i.test(raw)) {
    return new CliError(
      `The sign-in code expired before the login was completed (${raw}). ` +
      'If you clicked "Request approval", this login cannot finish: the approval is recorded for later, not for this code.',
      {
        code: 'login_expired', exitCode: ExitCode.AUTH,
        fix: `if approval was requested, an admin approves it in Entra (Enterprise applications → Admin consent requests); then run ${retry} and finish signing in within 15 minutes`,
      },
    );
  }
  return new CliError(`Login failed: ${raw}`, { code: 'login_failed', exitCode: ExitCode.AUTH });
}

const noClientIdError = () => new CliError(
  'No client ID configured. Set MS365_CLI_CLIENT_ID in your .env file or run: ms365 auth setup',
  { code: 'no_client_id', exitCode: ExitCode.AUTH, fix: 'ms365 auth setup' },
);

/** `auth login`/`logout`/`accounts`/`select`/`remove` under an admin profile (ID-8): exit 2. */
export function adminProfileUsageError(profileName?: string | null): CliError {
  const which = profileName ? `"${profileName}" is an admin profile` : 'This is an admin profile';
  return new CliError(`${which}: admin profiles sign in with az login, not ms365 (nothing is cached by ms365).`, {
    code: 'usage', exitCode: ExitCode.USAGE,
    fix: 'az login --tenant <tenant> for the admin profile, or drop --profile to manage ms365\'s own sign-in',
  });
}

const LOGIN_PAGE = (message: string) =>
  `<!doctype html><html><head><meta charset="utf-8"><title>ms365</title></head><body><p>${message}</p></body></html>`;
const LOGIN_SUCCESS_HTML = LOGIN_PAGE('Signed in to ms365. You can close this window and return to the terminal.');
const LOGIN_ERROR_HTML = LOGIN_PAGE('ms365 sign-in failed. Return to the terminal for details.');

export interface AuthManagerOptions {
  clientId?: string;
  tenantId?: string;
  /** Use this cached account (username, case-insensitive, or homeAccountId) for this process only. */
  account?: string;
  /**
   * Where access tokens come from: `msal` (ms365's own sign-in, refresh tokens cached), `az` (ID-8: an admin
   * profile borrows `az account get-access-token --tenant <tenantId>` on demand, held in memory only) or `external`
   * (AH-2: the host supplies MS365_CLI_TOKEN; nothing is cached and MSAL is never loaded). Default: `external` when
   * MS365_CLI_TOKEN is set, else `msal`.
   */
  source?: TokenSource;
  /** A `type: host` profile is in effect: external mode is mandatory and the token must be a verifiable JWT. */
  host?: boolean;
  /** The --profile name in effect, for messages. */
  profileName?: string;
}

export type TokenSource = 'msal' | 'az' | 'external';

/**
 * A silent token request MSAL cannot satisfy without the user: `InteractionRequiredAuthError` (interaction_required,
 * login_required, no_tokens_found, refresh_token_expired, …) or a server `invalid_grant` (revoked/expired refresh token).
 */
function needsInteraction(err: unknown): boolean {
  if (!err || typeof err !== 'object') return false;
  const e = err as { name?: unknown; errorCode?: unknown };
  return e.name === 'InteractionRequiredAuthError' || e.errorCode === 'invalid_grant' || e.errorCode === 'interaction_required';
}

export interface LoginOptions {
  /** Sign in with the system browser (auth code + PKCE, loopback redirect) instead of device code. */
  openBrowser?: (url: string) => Promise<void>;
}

export class AuthManager {
  private msalConfig: Configuration;
  private appP?: Promise<PublicClientApplication>;
  /** One plugin (and so one read-cached store) per manager, i.e. per process (ID-12). */
  private readonly cachePlugin = new FileTokenCachePlugin(fileCredentialStore(), storageResolver());
  private scopes: string[];
  private accessToken: string | null = null;
  private tokenExpiry: number | null = null;
  private selectedAccountId: string | null = null;
  private tokenRefreshPromise: Promise<string> | null = null;
  private accountOverride: string | null;
  private source: TokenSource;
  private profileName: string | null;
  private readonly host: boolean;
  /** The pinned tenant (profile, else MS365_CLI_TENANT_ID), compared with an external token's `tid`. */
  private readonly pinnedTenant: string | undefined;

  /** @param scopes scopes requested (and consented) at login */
  constructor(scopes: string[], options: AuthManagerOptions = {}) {
    this.scopes = scopes;
    this.host = !!options.host;
    this.source = options.source ?? (process.env.MS365_CLI_TOKEN || this.host ? 'external' : 'msal');
    this.profileName = options.profileName ?? null;
    // An az-backed manager has no cached accounts to pick from; an external one checks the token against it.
    this.accountOverride = this.source === 'az' ? null : options.account || null;
    this.pinnedTenant = options.tenantId || process.env.MS365_CLI_TENANT_ID || undefined;
    // Only an msal-backed manager signs anything in itself, so only it needs a client ID (grant-scopes asks for one too).
    const clientId = options.clientId
      || process.env.MS365_CLI_CLIENT_ID
      || '';
    if (!clientId && this.source === 'msal') throw noClientIdError();
    const tenantId = this.pinnedTenant || 'common';
    this.msalConfig = {
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
      cache: { cachePlugin: this.cachePlugin },
    };
  }

  /** MSAL costs ~45 ms to import, so it loads on first real use, not for --help/describe/dry-run (X-1). */
  private app(): Promise<PublicClientApplication> {
    this.appP ??= import('@azure/msal-node').then(m => new m.PublicClientApplication(this.msalConfig));
    return this.appP;
  }

  async initialize(): Promise<void> {
    if (this.source !== 'msal') return;
    await this.loadSelectedAccount();
  }

  /** `msal` (ms365's own sign-in), `az` (an admin profile's borrowed Azure CLI token) or `external` (supplied by a host). */
  tokenSource(): TokenSource {
    return this.source;
  }

  /** The --profile name in effect, if any. */
  getProfileName(): string | null {
    return this.profileName;
  }

  /** Under an admin profile there is no ms365 sign-in to manage: login/logout/accounts/select/remove are usage errors (ID-8). */
  assertMsalSource(): void {
    if (this.source === 'az') throw adminProfileUsageError(this.profileName);
  }

  private async loadSelectedAccount(): Promise<void> {
    const data = readFileIfExists(selectedAccountPath());
    try {
      if (data) this.selectedAccountId = JSON.parse(data).accountId ?? null;
    } catch {
      this.selectedAccountId = null;
    }
  }

  /** The selected account ID is not a secret: it lives in a plain 0600 file only. */
  private saveSelectedAccount(): void {
    if (this.selectedAccountId) {
      writeFileAtomic(selectedAccountPath(), JSON.stringify({ accountId: this.selectedAccountId }));
    } else {
      try { fs.unlinkSync(selectedAccountPath()); } catch { /* none */ }
    }
  }

  async getToken(forceRefresh = false): Promise<string> {
    // AH-2: a supplied token is checked on every use (a long run can outlive it) and never refreshed: a 401 asks for
    // a fresh one (forceRefresh comes only from GraphClient's 401 retry; cache-touching auth commands refuse earlier).
    if (this.source === 'external') {
      if (forceRefresh) throw tokenRejectedError();
      const token = requireExternalToken(this.profileName);
      checkExternalToken(token, { host: this.host, tenantId: this.pinnedTenant, account: this.accountOverride ?? undefined, profileName: this.profileName });
      return token;
    }

    if (!forceRefresh && this.accessToken && this.tokenExpiry && this.tokenExpiry - EXPIRY_SKEW_MS > Date.now()) {
      return this.accessToken;
    }

    // Coalesce concurrent refresh calls into a single request
    if (this.tokenRefreshPromise) {
      return this.tokenRefreshPromise;
    }

    this.tokenRefreshPromise = this.doTokenRefresh(forceRefresh);
    try {
      return await this.tokenRefreshPromise;
    } finally {
      this.tokenRefreshPromise = null;
    }
  }

  private async doTokenRefresh(forceRefresh: boolean): Promise<string> {
    if (this.source === 'az') return this.borrowAzToken();
    const account = await this.getCurrentAccount();
    if (!account) {
      throw new CliError('Not logged in. Run: ms365 auth login', { code: 'auth_required', exitCode: ExitCode.AUTH, fix: 'ms365 auth login' });
    }

    const silent = async (acct: AccountInfo) => (await this.app()).acquireTokenSilent({
      account: acct,
      scopes: [GRAPH_DEFAULT_SCOPE],
      forceRefresh,
    });
    let response;
    try {
      response = await silent(account);
    } catch (err) {
      if (!needsInteraction(err)) throw err;
      // ID-12 / IDT-5: this process read the store once; another process may since have refreshed or signed in
      // again. Re-read it (and the selection) and retry silently once before the normal re-login error.
      this.cachePlugin.invalidate();
      await this.loadSelectedAccount();
      const again = await this.getCurrentAccount();
      if (!again) throw err;
      response = await silent(again);
    }

    this.accessToken = response.accessToken;
    this.tokenExpiry = response.expiresOn ? new Date(response.expiresOn).getTime() : null;
    return this.accessToken;
  }

  /**
   * ID-8: an admin profile's token is az's signed-in user through the Azure CLI app, fetched on demand and kept in
   * memory only (until exp − 5 min); nothing is written. az refreshes its own session, so expiry means calling az again.
   */
  private async borrowAzToken(): Promise<string> {
    const tenant = this.getTenantId();
    // Loaded lazily: keeps grant-scopes (and GraphClient) out of auth.ts's import graph.
    const { azGraphToken } = await import('./grant-scopes.js');
    const az = azGraphToken(tenant);
    if (!az.ok) {
      const who = this.profileName ? `Admin profile "${this.profileName}"` : 'This admin profile';
      const why = az.reason === 'missing' ? 'the Azure CLI (az) is not installed'
        : az.reason === 'not_logged_in' ? `the Azure CLI is not signed in to tenant ${tenant} (${az.detail})`
          : `the Azure CLI could not provide a token (${az.detail})`;
      throw new CliError(`${who} borrows its token from the Azure CLI, but ${why}.`, {
        code: 'az_required', exitCode: ExitCode.AUTH,
        fix: az.reason === 'missing' ? `install the Azure CLI, then: az login --tenant ${tenant}` : `az login --tenant ${tenant}`,
      });
    }
    let exp: number | null = null;
    try {
      const e = decodeAccessToken(az.token.accessToken).exp;
      if (typeof e === 'number') exp = e * 1000;
    } catch { /* not a JWT: no expiry known, so it is borrowed again on the next call */ }
    this.accessToken = az.token.accessToken;
    this.tokenExpiry = exp;
    return this.accessToken;
  }

  /** The account in effect: the --account override, else the saved selection, else the first cached one. `undefined` when the override matches nothing. */
  private async pickAccount(): Promise<AccountInfo | null | undefined> {
    const accounts = await (await this.app()).getTokenCache().getAllAccounts();
    if (this.accountOverride) {
      const wanted = this.accountOverride.toLowerCase();
      return accounts.find(a => a.homeAccountId === this.accountOverride || a.username.toLowerCase() === wanted);
    }
    if (accounts.length === 0) return null;
    return accounts.find(a => a.homeAccountId === this.selectedAccountId) ?? accounts[0];
  }

  async getCurrentAccount(): Promise<AccountInfo | null> {
    const account = await this.pickAccount();
    if (account === undefined) {
      throw new CliError(`Account "${this.accountOverride}" is not signed in. List cached accounts with: ms365 auth accounts`,
        { code: 'account_not_found', exitCode: ExitCode.AUTH, fix: 'ms365 auth accounts' });
    }
    return account;
  }

  /** homeAccountId of the account this invocation uses (same precedence as getCurrentAccount()), or null when none matches. */
  async activeAccountId(): Promise<string | null> {
    return (await this.pickAccount())?.homeAccountId ?? null;
  }

  /** The --account / MS365_CLI_ACCOUNT / profile account for this process, if any. */
  getAccountOverride(): string | null {
    return this.accountOverride;
  }

  /** The app (client) ID this manager signs in with. */
  getClientId(): string {
    if (!this.msalConfig.auth.clientId) throw noClientIdError();
    return this.msalConfig.auth.clientId;
  }

  /** The tenant from the authority ('common' unless one is configured). */
  getTenantId(): string {
    return this.msalConfig.auth.authority?.split('/').pop() || 'common';
  }

  /** The scopes requested (and consented) at login. */
  requestedScopes(): string[] {
    return [...this.scopes];
  }

  /** True when this manager requests the org-mode scopes at login. */
  requestsOrgScopes(): boolean {
    return ORG_SCOPES.some(s => this.scopes.includes(s));
  }

  /** Granted (consented) scopes of the current access token, from its `scp` claim. */
  async grantedScopes(): Promise<string[]> {
    return tokenInfo(await this.getToken()).scopes;
  }

  async login(options: LoginOptions = {}): Promise<string> {
    this.assertMsalSource();
    const response = options.openBrowser
      ? await (await this.app()).acquireTokenInteractive({
        scopes: this.scopes,
        openBrowser: options.openBrowser,
        successTemplate: LOGIN_SUCCESS_HTML,
        errorTemplate: LOGIN_ERROR_HTML,
      })
      : await this.deviceCodeLogin();

    if (!response?.accessToken) {
      throw new CliError('Login did not return an access token.', { code: 'login_failed', exitCode: ExitCode.AUTH });
    }
    this.accessToken = response.accessToken;
    this.tokenExpiry = response.expiresOn ? new Date(response.expiresOn).getTime() : null;

    // The account that just signed in becomes the active one.
    if (response.account) {
      this.selectedAccountId = response.account.homeAccountId;
      this.saveSelectedAccount();
    }
    return this.accessToken;
  }

  private async deviceCodeLogin() {
    const response = await (await this.app()).acquireTokenByDeviceCode({
      scopes: this.scopes,
      deviceCodeCallback: (resp) => {
        // Write to stderr AND a file in user config dir so device code is always visible
        process.stderr.write('\n' + resp.message + '\n\n');
        try {
          ensureConfigDir();
          fs.writeFileSync(deviceCodePath(), resp.message + '\n', { mode: 0o600 });
        } catch { /* best effort */ }
      },
    });
    try { fs.unlinkSync(deviceCodePath()); } catch { /* not written */ }
    return response;
  }

  /** Remove every account and all cache, selection, device-code, temp and lock files. */
  async logout(): Promise<void> {
    this.assertMsalSource();
    // Every local file and keychain item goes even if MSAL cannot read the cache (e.g. the keychain is locked).
    let firstErr: unknown;
    try {
      const cache = (await this.app()).getTokenCache();
      for (const account of await cache.getAllAccounts()) {
        await cache.removeAccount(account);
      }
    } catch (err) { firstErr = err; }
    this.accessToken = null;
    this.tokenExpiry = null;
    this.selectedAccountId = null;
    this.cachePlugin.store.delete(TOKEN_CACHE_KEY);
    try { removeStoredCredentials(); } catch (err) { firstErr ??= err; }

    const dir = configDir();
    const names = fs.existsSync(dir) ? fs.readdirSync(dir) : [];
    for (const name of names) {
      if (/^(token-cache\.json|selected-account\.json|device-code\.txt)(\..*\.tmp|\.tmp|\.lock)?$/.test(name)) {
        try { fs.unlinkSync(path.join(dir, name)); } catch { /* already gone */ }
      }
    }
    if (firstErr) throw firstErr;
  }

  async listAccounts(): Promise<AccountInfo[]> {
    this.assertMsalSource();
    return (await this.app()).getTokenCache().getAllAccounts();
  }

  async selectAccount(accountId: string): Promise<boolean> {
    this.assertMsalSource();
    const accounts = await this.listAccounts();
    const account = accounts.find(a => a.homeAccountId === accountId);
    if (!account) return false;

    this.selectedAccountId = accountId;
    this.accessToken = null;
    this.tokenExpiry = null;
    this.saveSelectedAccount();
    return true;
  }

  async removeAccount(accountId: string): Promise<boolean> {
    this.assertMsalSource();
    const accounts = await this.listAccounts();
    const account = accounts.find(a => a.homeAccountId === accountId);
    if (!account) return false;

    await (await this.app()).getTokenCache().removeAccount(account);
    if (this.selectedAccountId === accountId) {
      this.selectedAccountId = null;
      this.accessToken = null;
      this.tokenExpiry = null;
      this.saveSelectedAccount();
    }
    return true;
  }

  async verifyLogin(): Promise<{ success: boolean; user?: { displayName: string; email: string } }> {
    try {
      const token = await this.getToken();
      const response = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: { Authorization: `Bearer ${token}` },
      });

      if (response.ok) {
        const data = await response.json();
        return {
          success: true,
          user: { displayName: data.displayName, email: data.userPrincipalName },
        };
      }
      return { success: false };
    } catch {
      return { success: false };
    }
  }

  getSelectedAccountId(): string | null {
    return this.selectedAccountId;
  }
}
