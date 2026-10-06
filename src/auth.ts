import fs from 'node:fs';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { randomBytes } from 'node:crypto';
import { AuthError, InteractionRequiredAuthError, PublicClientApplication } from '@azure/msal-node';
import type {
  AccountInfo,
  AuthenticationResult,
  ICachePlugin,
  INetworkModule,
  NetworkRequestOptions,
  NetworkResponse,
} from '@azure/msal-node';
import logger from './logger.js';

/**
 * Public client ID of "Microsoft Graph Command Line Tools" (now Microsoft Graph PowerShell), used when no
 * MS365_CLI_CLIENT_ID is configured. Microsoft controls this app; tenants may block it or require admin consent.
 */
export const DEFAULT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';

const TOKEN_CACHE_FILE = 'token-cache.json';
const SELECTED_ACCOUNT_FILE = 'selected-account.json';
const DEVICE_CODE_FILE = 'device-code.txt';

/** `~/.config/ms365-cli`, resolved per call so HOME is read at use time. */
export function configDir(): string {
  return path.join(os.homedir(), '.config', 'ms365-cli');
}

/** Scopes requested per command group. Setup and `buildScopes` both derive from this table. */
export const SCOPE_GROUPS: Readonly<Record<string, { label: string; org: boolean; scopes: readonly string[] }>> = {
  user: { label: 'User profile', org: false, scopes: ['User.Read'] },
  mail: { label: 'Mail, mail rules, mailbox settings / auto-replies', org: false, scopes: ['Mail.ReadWrite', 'Mail.Send', 'MailboxSettings.ReadWrite'] },
  calendar: { label: 'Calendar', org: false, scopes: ['Calendars.ReadWrite'] },
  files: { label: 'OneDrive / Excel', org: false, scopes: ['Files.ReadWrite'] },
  tasks: { label: 'To Do / Planner', org: false, scopes: ['Tasks.ReadWrite'] },
  contacts: { label: 'Contacts', org: false, scopes: ['Contacts.ReadWrite'] },
  onenote: { label: 'OneNote', org: false, scopes: ['Notes.Read', 'Notes.Create'] },
  people: { label: 'People search', org: false, scopes: ['People.Read'] },
  presence: { label: 'Presence', org: false, scopes: ['Presence.Read', 'Presence.ReadWrite'] },
  teams: {
    label: 'Teams and channel messages',
    org: true,
    scopes: ['Team.ReadBasic.All', 'Channel.ReadBasic.All', 'TeamMember.Read.All', 'ChannelMessage.Read.All', 'ChannelMessage.Send'],
  },
  chat: { label: 'Chats', org: true, scopes: ['Chat.Read', 'ChatMessage.Read', 'ChatMessage.Send'] },
  sharepoint: { label: 'SharePoint', org: true, scopes: ['Sites.Read.All'] },
  shared: { label: 'Shared mailboxes and calendars', org: true, scopes: ['Mail.Read.Shared', 'Mail.Send.Shared', 'Calendars.Read.Shared'] },
  directory: { label: 'Directory: users and groups', org: true, scopes: ['User.Read.All', 'Group.Read.All'] },
};

const SCOPE_HIERARCHY: Record<string, string[]> = {
  'Mail.ReadWrite': ['Mail.Read'],
  'Calendars.ReadWrite': ['Calendars.Read'],
  'Files.ReadWrite': ['Files.Read'],
  'Tasks.ReadWrite': ['Tasks.Read'],
  'Contacts.ReadWrite': ['Contacts.Read'],
  'Presence.ReadWrite': ['Presence.Read'],
  'MailboxSettings.ReadWrite': ['MailboxSettings.Read'],
};

export function buildScopes(orgMode: boolean): string[] {
  const scopeSet = new Set(
    Object.values(SCOPE_GROUPS).filter((g) => orgMode || !g.org).flatMap((g) => g.scopes),
  );
  for (const [higher, lowers] of Object.entries(SCOPE_HIERARCHY)) {
    if (scopeSet.has(higher)) {
      for (const lower of lowers) scopeSet.delete(lower);
    }
  }
  return Array.from(scopeSet);
}

/** No usable login: no cached account, or the cached one needs a new interactive sign-in. */
export class NotLoggedInError extends Error {
  constructor(message = 'Not logged in. Run: ms365 auth login') {
    super(message);
    this.name = 'NotLoggedInError';
  }
}

/**
 * Write `data` to `file` (mode 0600, inside a 0700 dir) via a unique temp file and rename, so concurrent
 * writers never leave a partial file: readers see one complete version, and the last rename wins.
 */
export async function writeSecureFile(file: string, data: string): Promise<void> {
  await fsp.mkdir(path.dirname(file), { recursive: true, mode: 0o700 });
  const tmp = `${file}.${process.pid}.${randomBytes(6).toString('hex')}`;
  try {
    await fsp.writeFile(tmp, data, { mode: 0o600, flag: 'wx' });
    await fsp.rename(tmp, file);
  } catch (err) {
    await fsp.rm(tmp, { force: true });
    throw err;
  }
}

/** MSAL cache plugin over a JSON file: read before every cache access, write back when MSAL changed it. */
export function fileCachePlugin(file: string): ICachePlugin {
  return {
    async beforeCacheAccess(context) {
      let data: string;
      try {
        data = await fsp.readFile(file, 'utf8');
      } catch (err) {
        if (err instanceof Error && 'code' in err && err.code === 'ENOENT') return;
        throw err;
      }
      try {
        JSON.parse(data);
      } catch {
        // MSAL keeps the raw string as a merge snapshot, so a corrupt file must never reach deserialize().
        logger.warn(`Ignoring unreadable token cache ${file}; run: ms365 auth login`);
        return;
      }
      context.tokenCache.deserialize(data);
    },
    async afterCacheAccess(context) {
      if (context.cacheHasChanged) await writeSecureFile(file, context.tokenCache.serialize());
    },
  };
}

/** Tenant-side refusals of the built-in app; the fix is registering a custom app (setup option 2). */
const CUSTOM_APP_CODES = new Set([65001, 50105, 700016, 90094, 53003, 7000112, 50020, 500208]);

const CUSTOM_APP_HINT =
  '\nYour organization blocks or restricts the built-in app (Microsoft Graph Command Line Tools).' +
  '\nRegister your own app instead: run `ms365 auth setup` and choose [2], then set MS365_CLI_CLIENT_ID.';

interface AadFailure {
  description: string;
  codes: number[];
}

/**
 * MSAL network client over global fetch. Behaves like MSAL's default client, but remembers the first identity
 * platform error of a flow: the device-code flow otherwise reduces it to e.g. "invalid_grant", dropping the
 * AADSTS code users need.
 */
class AadNetworkClient implements INetworkModule {
  failure: AadFailure | null = null;

  sendGetRequestAsync<T>(url: string, options?: NetworkRequestOptions, timeout?: number): Promise<NetworkResponse<T>> {
    return this.send<T>(url, { method: 'GET', headers: options?.headers }, timeout);
  }

  async sendPostRequestAsync<T>(url: string, options?: NetworkRequestOptions): Promise<NetworkResponse<T>> {
    const response = await this.send<T>(url, { method: 'POST', headers: options?.headers, body: options?.body ?? '' });
    // MSAL does not check the device-code response and would poll with an undefined code; fail here instead.
    if (url.includes('/devicecode') && response.status >= 400) {
      const description = this.failure?.description ?? `device code request failed with HTTP ${response.status}`;
      throw new AuthError('device_code_request_failed', options?.correlationId ?? '', description);
    }
    return response;
  }

  private async send<T>(url: string, init: RequestInit, timeout?: number): Promise<NetworkResponse<T>> {
    let response: Response;
    try {
      response = await fetch(url, { ...init, signal: timeout ? AbortSignal.timeout(timeout) : undefined });
    } catch (err) {
      this.failure ??= { description: `Network request to ${new URL(url).host} failed: ${fetchFailureReason(err)}`, codes: [] };
      throw err;
    }
    const body: unknown = await response.json();
    if (body && typeof body === 'object' && 'error' in body && typeof body.error === 'string' && body.error !== 'authorization_pending') {
      const description = 'error_description' in body && typeof body.error_description === 'string'
        ? body.error_description.split(/\r?\n/)[0]
        : body.error;
      const codes = 'error_codes' in body && Array.isArray(body.error_codes)
        ? body.error_codes.filter((c): c is number => typeof c === 'number')
        : [];
      this.failure ??= { description, codes };
    }
    // MSAL's INetworkModule is generic over the body type it expects; MSAL validates the fields it reads.
    const typedBody = body as T;
    return { headers: Object.fromEntries(response.headers), body: typedBody, status: response.status };
  }
}

/** Why a fetch() call threw: undici reports the real reason (DNS, refused, TLS) in `cause`. */
function fetchFailureReason(err: unknown): string {
  if (err instanceof Error) return err.cause instanceof Error ? err.cause.message : err.message;
  return String(err);
}

export interface AuthManagerOptions {
  /** Overrides MS365_CLI_CLIENT_ID / MS365_MCP_CLIENT_ID / the built-in default (used by the setup wizard). */
  clientId?: string;
}

/** A pre-obtained token from the environment (CLI or MCP server name); it bypasses the token cache entirely. */
export function envToken(): string | undefined {
  return process.env.MS365_CLI_TOKEN || process.env.MS365_MCP_OAUTH_TOKEN || undefined;
}

export class AuthManager {
  readonly clientId: string;
  private readonly msalApp: PublicClientApplication;
  private readonly network = new AadNetworkClient();
  private readonly scopes: string[];
  private accessToken: string | null = null;
  private tokenExpiry: number | null = null;
  private selectedAccountId: string | null = null;
  private tokenRefreshPromise: Promise<string> | null = null;

  constructor(scopes: string[], options: AuthManagerOptions = {}) {
    this.scopes = scopes;
    // Accept both MS365_CLI_* and MS365_MCP_* env var names (MCP server compat)
    this.clientId = options.clientId || process.env.MS365_CLI_CLIENT_ID || process.env.MS365_MCP_CLIENT_ID || DEFAULT_CLIENT_ID;
    const tenantId = process.env.MS365_CLI_TENANT_ID || process.env.MS365_MCP_TENANT_ID || 'common';
    this.msalApp = new PublicClientApplication({
      auth: { clientId: this.clientId, authority: `https://login.microsoftonline.com/${tenantId}` },
      cache: { cachePlugin: fileCachePlugin(path.join(configDir(), TOKEN_CACHE_FILE)) },
      system: { networkClient: this.network },
    });
  }

  /** Restrict the config dir (0700) and credential files (0600), even if they already exist; load the selected account. */
  async initialize(): Promise<void> {
    const dir = configDir();
    fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
    fs.chmodSync(dir, 0o700);
    for (const name of [TOKEN_CACHE_FILE, SELECTED_ACCOUNT_FILE]) {
      try {
        fs.chmodSync(path.join(dir, name), 0o600);
      } catch (err) {
        if (!(err instanceof Error && 'code' in err && err.code === 'ENOENT')) throw err;
      }
    }
    try {
      const accountId: unknown = JSON.parse(fs.readFileSync(path.join(dir, SELECTED_ACCOUNT_FILE), 'utf8')).accountId;
      this.selectedAccountId = typeof accountId === 'string' ? accountId : null;
    } catch {
      this.selectedAccountId = null; // missing or unreadable: fall back to the first cached account
    }
  }

  private saveSelectedAccount(): Promise<void> {
    return writeSecureFile(path.join(configDir(), SELECTED_ACCOUNT_FILE), JSON.stringify({ accountId: this.selectedAccountId }));
  }

  async getToken(forceRefresh = false): Promise<string> {
    const token = envToken();
    if (token) return token;

    if (this.accessToken && this.tokenExpiry && this.tokenExpiry > Date.now() && !forceRefresh) {
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
    if (!account) throw new NotLoggedInError();

    this.network.failure = null;
    let response: AuthenticationResult;
    try {
      // forceRefresh skips MSAL's cached access token (used after Graph rejected it with 401).
      response = await this.msalApp.acquireTokenSilent({ account, scopes: this.scopes, forceRefresh });
    } catch (err) {
      throw this.explain(err, 'Could not refresh the access token');
    }
    this.remember(response);
    return response.accessToken;
  }

  private remember(response: AuthenticationResult): void {
    this.accessToken = response.accessToken;
    this.tokenExpiry = response.expiresOn ? new Date(response.expiresOn).getTime() : null;
  }

  /** Turn an MSAL failure into a user-facing error, keeping the AADSTS detail and suggesting the next step. */
  private explain(err: unknown, context: string): Error {
    const failure = this.network.failure;
    const detail = failure?.description ?? (err instanceof Error ? err.message : String(err));
    const codes = new Set([...(failure?.codes ?? []), ...[...detail.matchAll(/AADSTS(\d+)/g)].map((m) => Number(m[1]))]);
    const hint = this.clientId === DEFAULT_CLIENT_ID && [...codes].some((c) => CUSTOM_APP_CODES.has(c)) ? CUSTOM_APP_HINT : '';
    if (err instanceof InteractionRequiredAuthError) {
      return new NotLoggedInError(`Permissions changed or sign-in expired (${detail}) — run: ms365 auth login${hint}`);
    }
    return new Error(`${context}: ${detail}${hint}`);
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

  /** Device-code login; the message goes to stderr only. A device-code.txt left by older versions is removed. */
  async login(): Promise<void> {
    const codeFile = path.join(configDir(), DEVICE_CODE_FILE);
    this.network.failure = null;
    let response: AuthenticationResult | null;
    try {
      response = await this.msalApp.acquireTokenByDeviceCode({
        scopes: this.scopes,
        deviceCodeCallback: (resp) => {
          process.stderr.write('\n' + resp.message + '\n\n');
        },
      });
    } catch (err) {
      throw this.explain(err, 'Sign-in failed');
    } finally {
      fs.rmSync(codeFile, { force: true });
    }
    if (!response) throw new Error('Sign-in returned no token.');
    this.remember(response);

    if (!this.selectedAccountId && response.account) {
      this.selectedAccountId = response.account.homeAccountId;
      await this.saveSelectedAccount();
    }
  }

  async logout(): Promise<void> {
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();
    for (const account of accounts) {
      await this.msalApp.getTokenCache().removeAccount(account);
    }
    this.accessToken = null;
    this.tokenExpiry = null;
    this.selectedAccountId = null;
    for (const name of [TOKEN_CACHE_FILE, SELECTED_ACCOUNT_FILE, DEVICE_CODE_FILE]) {
      fs.rmSync(path.join(configDir(), name), { force: true });
    }
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
    await this.saveSelectedAccount();
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
      await this.saveSelectedAccount();
    }
    return true;
  }

  /**
   * The signed-in user per Microsoft Graph `/me`. Throws NotLoggedInError when there is no usable login,
   * and a plain Error when Graph or the network fails (the login may still be fine).
   */
  async whoAmI(): Promise<{ displayName: string; email: string }> {
    const token = await this.getToken();
    let response: Response;
    try {
      response = await fetch('https://graph.microsoft.com/v1.0/me', { headers: { Authorization: `Bearer ${token}` } });
    } catch (err) {
      throw new Error(`Could not verify the login: Microsoft Graph is unreachable (${fetchFailureReason(err)})`);
    }
    if (response.status === 401) {
      throw new NotLoggedInError('Microsoft Graph rejected the saved login (401). Run: ms365 auth login');
    }
    if (!response.ok) {
      throw new Error(`Could not verify the login: Microsoft Graph returned ${response.status} ${response.statusText}`.trimEnd());
    }
    const me: Record<string, unknown> = Object(await response.json());
    const text = (v: unknown) => (typeof v === 'string' ? v : '');
    return { displayName: text(me.displayName), email: text(me.userPrincipalName) || text(me.mail) };
  }

  getSelectedAccountId(): string | null {
    return this.selectedAccountId;
  }
}
