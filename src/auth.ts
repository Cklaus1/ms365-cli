import type { AccountInfo, Configuration, ICachePlugin, PublicClientApplication, TokenCacheContext } from '@azure/msal-node';
import fs from 'fs';
import path from 'path';
import os from 'os';
import crypto from 'crypto';
import { CliError, ExitCode } from './errors.js';

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

/**
 * MSAL cache plugin backed by a single 0600 file (no OS keyring; see spec D28 / N6). MSAL calls afterCacheAccess after every token
 * operation and sets cacheHasChanged when a token was fetched from the network, so silent
 * refreshes (and rotated refresh tokens) are persisted without extra code paths.
 */
export class FileTokenCachePlugin implements ICachePlugin {
  async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
    const data = readFileIfExists(tokenCachePath());
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

const LOGIN_PAGE = (message: string) =>
  `<!doctype html><html><head><meta charset="utf-8"><title>ms365</title></head><body><p>${message}</p></body></html>`;
const LOGIN_SUCCESS_HTML = LOGIN_PAGE('Signed in to ms365. You can close this window and return to the terminal.');
const LOGIN_ERROR_HTML = LOGIN_PAGE('ms365 sign-in failed. Return to the terminal for details.');

export interface AuthManagerOptions {
  clientId?: string;
  tenantId?: string;
  /** Use this cached account (username, case-insensitive, or homeAccountId) for this process only. */
  account?: string;
}

export interface LoginOptions {
  /** Sign in with the system browser (auth code + PKCE, loopback redirect) instead of device code. */
  openBrowser?: (url: string) => Promise<void>;
}

export class AuthManager {
  private msalConfig: Configuration;
  private appP?: Promise<PublicClientApplication>;
  private scopes: string[];
  private accessToken: string | null = null;
  private tokenExpiry: number | null = null;
  private selectedAccountId: string | null = null;
  private tokenRefreshPromise: Promise<string> | null = null;
  private accountOverride: string | null;

  /** @param scopes scopes requested (and consented) at login */
  constructor(scopes: string[], options: AuthManagerOptions = {}) {
    this.scopes = scopes;
    this.accountOverride = options.account || null;
    const clientId = options.clientId
      || process.env.MS365_CLI_CLIENT_ID;
    if (!clientId) {
      throw new CliError(
        'No client ID configured. Set MS365_CLI_CLIENT_ID in your .env file or run: ms365 auth setup',
        { code: 'no_client_id', exitCode: ExitCode.AUTH, fix: 'ms365 auth setup' },
      );
    }
    const tenantId = options.tenantId
      || process.env.MS365_CLI_TENANT_ID
      || 'common';
    this.msalConfig = {
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
      cache: { cachePlugin: new FileTokenCachePlugin() },
    };
  }

  /** MSAL costs ~45 ms to import, so it loads on first real use, not for --help/describe/dry-run (X-1). */
  private app(): Promise<PublicClientApplication> {
    this.appP ??= import('@azure/msal-node').then(m => new m.PublicClientApplication(this.msalConfig));
    return this.appP;
  }

  async initialize(): Promise<void> {
    await this.loadSelectedAccount();
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
    // A token supplied via env (e.g. from another tool) bypasses the cache
    const envToken = process.env.MS365_CLI_TOKEN;
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

    const response = await (await this.app()).acquireTokenSilent({
      account,
      scopes: [GRAPH_DEFAULT_SCOPE],
      forceRefresh,
    });

    this.accessToken = response.accessToken;
    this.tokenExpiry = response.expiresOn ? new Date(response.expiresOn).getTime() : null;
    return this.accessToken;
  }

  async getCurrentAccount(): Promise<AccountInfo | null> {
    const accounts = await (await this.app()).getTokenCache().getAllAccounts();
    if (this.accountOverride) {
      const wanted = this.accountOverride;
      const match = accounts.find(a => a.homeAccountId === wanted || a.username.toLowerCase() === wanted.toLowerCase());
      if (!match) {
        throw new CliError(`Account "${wanted}" is not signed in. List cached accounts with: ms365 auth accounts`,
          { code: 'account_not_found', exitCode: ExitCode.AUTH, fix: 'ms365 auth accounts' });
      }
      return match;
    }
    if (accounts.length === 0) return null;

    if (this.selectedAccountId) {
      const selected = accounts.find(a => a.homeAccountId === this.selectedAccountId);
      if (selected) return selected;
    }

    return accounts[0];
  }

  /** Granted (consented) scopes of the current access token, from its `scp` claim. */
  async grantedScopes(): Promise<string[]> {
    return tokenInfo(await this.getToken()).scopes;
  }

  async login(options: LoginOptions = {}): Promise<string> {
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
    const cache = (await this.app()).getTokenCache();
    for (const account of await cache.getAllAccounts()) {
      await cache.removeAccount(account);
    }
    this.accessToken = null;
    this.tokenExpiry = null;
    this.selectedAccountId = null;

    const dir = configDir();
    const names = fs.existsSync(dir) ? fs.readdirSync(dir) : [];
    for (const name of names) {
      if (/^(token-cache\.json|selected-account\.json|device-code\.txt)(\..*\.tmp|\.tmp|\.lock)?$/.test(name)) {
        try { fs.unlinkSync(path.join(dir, name)); } catch { /* already gone */ }
      }
    }
  }

  async listAccounts(): Promise<AccountInfo[]> {
    return (await this.app()).getTokenCache().getAllAccounts();
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
