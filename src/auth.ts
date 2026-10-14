import type { AccountInfo, Configuration, ICachePlugin, TokenCacheContext } from '@azure/msal-node';
import { PublicClientApplication } from '@azure/msal-node';
import fs from 'fs';
import path from 'path';
import os from 'os';
import crypto from 'crypto';
import { spawnSync } from 'child_process';
import { createRequire } from 'module';
import { CliError, ExitCode } from './errors.js';

// libsecret needs a D-Bus session bus. Without one (headless Linux, WSLg with DISPLAY set),
// it tries D-Bus autolaunch via X11 and keytar hangs in a native thread.
const isHeadless = process.platform === 'linux' && !process.env.DBUS_SESSION_BUS_ADDRESS;

const SERVICE_NAME = 'ms365-cli';
const TOKEN_CACHE_ACCOUNT = 'msal-token-cache';
const SELECTED_ACCOUNT_KEY = 'selected-account';

/** Silent token requests use .default: the token carries whatever the user consented to at login. */
export const GRAPH_DEFAULT_SCOPE = 'https://graph.microsoft.com/.default';

/** Treat access tokens as expired this long before they actually expire. */
const EXPIRY_SKEW_MS = 5 * 60 * 1000;

export function configDir(): string {
  return process.env.MS365_CLI_CONFIG_DIR || path.join(os.homedir(), '.config', 'ms365-cli');
}
const tokenCachePath = () => path.join(configDir(), 'token-cache.json');
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

/** Atomically replace `target` (unique temp file + rename) under a cross-process lock. */
export function writeFileAtomic(target: string, data: string): void {
  ensureConfigDir();
  withFileLock(target, () => {
    const tmp = `${target}.${process.pid}.${crypto.randomBytes(4).toString('hex')}.tmp`;
    try {
      fs.writeFileSync(tmp, data, { mode: 0o600 });
      fs.renameSync(tmp, target);
    } catch (err) {
      try { fs.unlinkSync(tmp); } catch { /* not created */ }
      throw err;
    }
  });
}

function readFileIfExists(p: string): string | undefined {
  try {
    const data = fs.readFileSync(p, 'utf8');
    return data || undefined;
  } catch {
    return undefined;
  }
}

// ── keytar: optional, used only to migrate entries written by older versions ──
//
// keytar calls can hang inside a native thread, and a hung thread also blocks
// process.exit(). So keytar never loads in this process: each call runs in a short-lived
// child process that is SIGKILLed on timeout.

const KEYTAR_TIMEOUT_MS = 3000;

function keyringDisabled(): boolean {
  return isHeadless || process.env.MS365_CLI_NO_KEYRING === '1';
}

let keytarPath: string | null | undefined;
function resolveKeytar(): string | null {
  if (keytarPath === undefined) {
    try {
      keytarPath = createRequire(import.meta.url).resolve('keytar');
    } catch {
      keytarPath = null; // optional dependency not installed
    }
  }
  return keytarPath;
}

/** Run one keytar call in a child process. Returns undefined if keytar is unavailable, failed or hung. */
function keytarCall(op: 'get' | 'delete', account: string): string | undefined {
  if (keyringDisabled()) return undefined;
  const mod = resolveKeytar();
  if (!mod) return undefined;
  const script = `const k=require(${JSON.stringify(mod)});` +
    `const p=${op === 'get' ? 'k.getPassword' : 'k.deletePassword'}(${JSON.stringify(SERVICE_NAME)},${JSON.stringify(account)});` +
    `p.then(v=>{process.stdout.write(v===null||v===undefined?'':String(v));process.exit(0)},()=>process.exit(2));`;
  const r = spawnSync(process.execPath, ['-e', script], { timeout: KEYTAR_TIMEOUT_MS, killSignal: 'SIGKILL', encoding: 'utf8' });
  return r.status === 0 ? r.stdout : undefined;
}

const migrationMarker = () => path.join(configDir(), '.keyring-migrated');

/** Read a legacy keyring entry at most once per config dir (the marker prevents repeated probes). */
function keytarMigrate(account: string): string | undefined {
  if (keyringDisabled() || fs.existsSync(migrationMarker())) return undefined;
  const value = keytarCall('get', account) || undefined;
  return value;
}

function markMigrated(): void {
  try {
    ensureConfigDir();
    fs.writeFileSync(migrationMarker(), '', { mode: 0o600 });
  } catch { /* best effort */ }
}

/** Delete a legacy keyring entry; false only when keytar is present but the delete failed or hung. */
function keytarDelete(account: string): boolean {
  if (keyringDisabled() || !resolveKeytar()) return true;
  return keytarCall('delete', account) !== undefined;
}

/**
 * MSAL cache plugin backed by a 0600 file. MSAL calls afterCacheAccess after every token
 * operation and sets cacheHasChanged when a token was fetched from the network, so silent
 * refreshes (and rotated refresh tokens) are persisted without extra code paths.
 */
export class FileTokenCachePlugin implements ICachePlugin {
  async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
    let data = readFileIfExists(tokenCachePath());
    if (!data) {
      // One-time migration from the OS keyring used by older versions.
      data = keytarMigrate(TOKEN_CACHE_ACCOUNT);
      if (data) writeFileAtomic(tokenCachePath(), data);
      markMigrated();
    }
    if (data) {
      try {
        ctx.tokenCache.deserialize(data);
      } catch {
        // Corrupt cache file: start empty (next login rewrites it).
      }
    }
  }

  async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
    if (ctx.cacheHasChanged) {
      writeFileAtomic(tokenCachePath(), ctx.tokenCache.serialize());
    }
  }
}

export interface AuthManagerOptions {
  clientId?: string;
  tenantId?: string;
}

export class AuthManager {
  private msalApp: PublicClientApplication;
  private scopes: string[];
  private accessToken: string | null = null;
  private tokenExpiry: number | null = null;
  private selectedAccountId: string | null = null;
  private tokenRefreshPromise: Promise<string> | null = null;

  /** @param scopes scopes requested (and consented) at login */
  constructor(scopes: string[], options: AuthManagerOptions = {}) {
    this.scopes = scopes;
    // Accept both MS365_CLI_* and MS365_MCP_* env var names (MCP server compat)
    const clientId = options.clientId
      || process.env.MS365_CLI_CLIENT_ID
      || process.env.MS365_MCP_CLIENT_ID;
    if (!clientId) {
      throw new CliError(
        'No client ID configured. Set MS365_CLI_CLIENT_ID in your .env file or run: ms365 auth setup',
        { code: 'no_client_id', exitCode: ExitCode.AUTH, fix: 'ms365 auth setup' },
      );
    }
    const tenantId = options.tenantId
      || process.env.MS365_CLI_TENANT_ID
      || process.env.MS365_MCP_TENANT_ID
      || 'common';
    const config: Configuration = {
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
      cache: { cachePlugin: new FileTokenCachePlugin() },
    };
    this.msalApp = new PublicClientApplication(config);
  }

  async initialize(): Promise<void> {
    await this.loadSelectedAccount();
  }

  private async loadSelectedAccount(): Promise<void> {
    let data = readFileIfExists(selectedAccountPath());
    if (!data && !fs.existsSync(migrationMarker())) {
      data = keytarMigrate(SELECTED_ACCOUNT_KEY);
      if (data) writeFileAtomic(selectedAccountPath(), data);
    }
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
    // Check env var token first (accept both CLI and MCP server names)
    const envToken = process.env.MS365_CLI_TOKEN || process.env.MS365_MCP_OAUTH_TOKEN;
    if (envToken) return envToken;

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
    const account = await this.getCurrentAccount();
    if (!account) {
      throw new CliError('Not logged in. Run: ms365 auth login', { code: 'auth_required', exitCode: ExitCode.AUTH, fix: 'ms365 auth login' });
    }

    const response = await this.msalApp.acquireTokenSilent({
      account,
      scopes: [GRAPH_DEFAULT_SCOPE],
      forceRefresh,
    });

    this.accessToken = response.accessToken;
    this.tokenExpiry = response.expiresOn ? new Date(response.expiresOn).getTime() : null;
    return this.accessToken;
  }

  async getCurrentAccount(): Promise<AccountInfo | null> {
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();
    if (accounts.length === 0) return null;

    if (this.selectedAccountId) {
      const selected = accounts.find(a => a.homeAccountId === this.selectedAccountId);
      if (selected) return selected;
    }

    return accounts[0];
  }

  async login(): Promise<string> {
    const response = await this.msalApp.acquireTokenByDeviceCode({
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

  /** Remove every account, the cache files, and any legacy keyring entries. Returns warnings. */
  async logout(): Promise<string[]> {
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();
    for (const account of accounts) {
      await this.msalApp.getTokenCache().removeAccount(account);
    }
    this.accessToken = null;
    this.tokenExpiry = null;
    this.selectedAccountId = null;

    // Files first, so a hanging keyring can't leave the file cache behind.
    const dir = configDir();
    const names = fs.existsSync(dir) ? fs.readdirSync(dir) : [];
    for (const name of names) {
      if (/^(token-cache\.json|selected-account\.json|device-code\.txt)(\..*\.tmp|\.tmp|\.lock)?$/.test(name)) {
        try { fs.unlinkSync(path.join(dir, name)); } catch { /* already gone */ }
      }
    }

    const warnings: string[] = [];
    for (const key of [TOKEN_CACHE_ACCOUNT, SELECTED_ACCOUNT_KEY]) {
      if (!keytarDelete(key)) {
        warnings.push(`Could not remove the "${key}" entry from the OS keyring; remove "${SERVICE_NAME}" entries manually.`);
      }
    }
    return warnings;
  }

  async listAccounts(): Promise<AccountInfo[]> {
    return this.msalApp.getTokenCache().getAllAccounts();
  }

  async selectAccount(accountId: string): Promise<boolean> {
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
    const accounts = await this.listAccounts();
    const account = accounts.find(a => a.homeAccountId === accountId);
    if (!account) return false;

    await this.msalApp.getTokenCache().removeAccount(account);
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
