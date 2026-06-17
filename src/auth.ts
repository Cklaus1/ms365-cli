import type { AccountInfo, Configuration } from '@azure/msal-node';
import { PublicClientApplication } from '@azure/msal-node';
import keytar from 'keytar';
import fs from 'fs';
import path from 'path';
import os from 'os';

// Detect headless Linux where keytar hangs indefinitely trying to access GNOME Keyring.
// Skip keytar entirely in this case — use file-based fallback instead.
const isHeadless = process.platform === 'linux' && !process.env.DISPLAY && !process.env.DBUS_SESSION_BUS_ADDRESS;

/** Wrap keytar calls with a timeout — keytar hangs on headless Linux without a keyring */
function withTimeout<T>(promise: Promise<T>, ms: number = 2000): Promise<T | null> {
  if (isHeadless) return Promise.resolve(null);
  return Promise.race([
    promise,
    new Promise<null>((resolve) => {
      const timer = setTimeout(() => resolve(null), ms);
      // Don't let this timer keep the event loop alive
      if (timer.unref) timer.unref();
    }),
  ]);
}

const SERVICE_NAME = 'ms365-cli';
const TOKEN_CACHE_ACCOUNT = 'msal-token-cache';
const SELECTED_ACCOUNT_KEY = 'selected-account';

const CONFIG_DIR = path.join(os.homedir(), '.config', 'ms365-cli');
const FALLBACK_TOKEN_PATH = path.join(CONFIG_DIR, 'token-cache.json');
const SELECTED_ACCOUNT_PATH = path.join(CONFIG_DIR, 'selected-account.json');

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
];

function ensureConfigDir(): void {
  if (!fs.existsSync(CONFIG_DIR)) {
    fs.mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
  }
}

export function buildScopes(orgMode: boolean): string[] {
  const scopes = [...PERSONAL_SCOPES];
  if (orgMode) {
    scopes.push(...ORG_SCOPES);
  }

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

export class AuthManager {
  private msalApp: PublicClientApplication;
  private scopes: string[];
  private accessToken: string | null = null;
  private tokenExpiry: number | null = null;
  private selectedAccountId: string | null = null;
  private tokenRefreshPromise: Promise<string> | null = null;

  constructor(scopes: string[]) {
    this.scopes = scopes;
    // Accept both MS365_CLI_* and MS365_MCP_* env var names (MCP server compat)
    const clientId = process.env.MS365_CLI_CLIENT_ID
      || process.env.MS365_MCP_CLIENT_ID;
    if (!clientId) {
      throw new Error(
        'No client ID configured. Set MS365_CLI_CLIENT_ID in your .env file or run: ms365 auth setup'
      );
    }
    const tenantId = process.env.MS365_CLI_TENANT_ID
      || process.env.MS365_MCP_TENANT_ID
      || 'common';
    const config: Configuration = {
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
    };
    this.msalApp = new PublicClientApplication(config);
  }

  async initialize(): Promise<void> {
    await this.loadTokenCache();
    await this.loadSelectedAccount();
  }

  private async loadTokenCache(): Promise<void> {
    try {
      let cacheData: string | undefined;

      // Try file fallback first — avoids keytar which hangs on some Linux setups
      if (fs.existsSync(FALLBACK_TOKEN_PATH)) {
        cacheData = fs.readFileSync(FALLBACK_TOKEN_PATH, 'utf8');
      }

      // Only try keytar if no file cache exists
      if (!cacheData) {
        try {
          const data = await withTimeout(keytar.getPassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT));
          if (data) cacheData = data;
        } catch {
          // keytar failed
        }
      }

      if (cacheData) {
        this.msalApp.getTokenCache().deserialize(cacheData);
      }
    } catch {
      // ignore cache load errors
    }
  }

  private async saveTokenCache(): Promise<void> {
    ensureConfigDir();
    const cacheData = this.msalApp.getTokenCache().serialize();
    try {
      await withTimeout(keytar.setPassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT, cacheData));
    } catch {
      // Atomic write: write to temp file then rename to prevent corruption
      const tmpPath = FALLBACK_TOKEN_PATH + '.tmp';
      fs.writeFileSync(tmpPath, cacheData, { mode: 0o600 });
      fs.renameSync(tmpPath, FALLBACK_TOKEN_PATH);
    }
  }

  private async loadSelectedAccount(): Promise<void> {
    try {
      let data: string | undefined;

      // Try file fallback first — avoids keytar which hangs on some Linux setups
      if (fs.existsSync(SELECTED_ACCOUNT_PATH)) {
        data = fs.readFileSync(SELECTED_ACCOUNT_PATH, 'utf8');
      }

      // Only try keytar if no file exists
      if (!data) {
        try {
          const d = await withTimeout(keytar.getPassword(SERVICE_NAME, SELECTED_ACCOUNT_KEY));
          if (d) data = d;
        } catch {
          // keytar failed
        }
      }

      if (data) {
        this.selectedAccountId = JSON.parse(data).accountId;
      }
    } catch {
      // ignore
    }
  }

  private async saveSelectedAccount(): Promise<void> {
    ensureConfigDir();
    const data = JSON.stringify({ accountId: this.selectedAccountId });
    try {
      await withTimeout(keytar.setPassword(SERVICE_NAME, SELECTED_ACCOUNT_KEY, data));
    } catch {
      const tmpPath = SELECTED_ACCOUNT_PATH + '.tmp';
      fs.writeFileSync(tmpPath, data, { mode: 0o600 });
      fs.renameSync(tmpPath, SELECTED_ACCOUNT_PATH);
    }
  }

  async getToken(forceRefresh = false): Promise<string> {
    // Check env var token first (accept both CLI and MCP server names)
    const envToken = process.env.MS365_CLI_TOKEN || process.env.MS365_MCP_OAUTH_TOKEN;
    if (envToken) return envToken;

    if (this.accessToken && this.tokenExpiry && this.tokenExpiry > Date.now() && !forceRefresh) {
      return this.accessToken;
    }

    // Coalesce concurrent refresh calls into a single request
    if (this.tokenRefreshPromise) {
      return this.tokenRefreshPromise;
    }

    this.tokenRefreshPromise = this.doTokenRefresh();
    try {
      return await this.tokenRefreshPromise;
    } finally {
      this.tokenRefreshPromise = null;
    }
  }

  private async doTokenRefresh(): Promise<string> {
    const account = await this.getCurrentAccount();
    if (!account) {
      throw new Error('Not logged in. Run: ms365 auth login');
    }

    const response = await this.msalApp.acquireTokenSilent({
      account,
      scopes: this.scopes,
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
          fs.writeFileSync(path.join(CONFIG_DIR, 'device-code.txt'), resp.message + '\n', { mode: 0o600 });
        } catch {}
      },
    });

    this.accessToken = response?.accessToken || null;
    this.tokenExpiry = response?.expiresOn ? new Date(response.expiresOn).getTime() : null;

    if (!this.selectedAccountId && response?.account) {
      this.selectedAccountId = response.account.homeAccountId;
      await this.saveSelectedAccount();
    }

    await this.saveTokenCache();
    return this.accessToken!;
  }

  async logout(): Promise<void> {
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();
    for (const account of accounts) {
      await this.msalApp.getTokenCache().removeAccount(account);
    }
    this.accessToken = null;
    this.tokenExpiry = null;
    this.selectedAccountId = null;

    try {
      await withTimeout(keytar.deletePassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT));
      await withTimeout(keytar.deletePassword(SERVICE_NAME, SELECTED_ACCOUNT_KEY));
    } catch {
      // ignore keytar errors
    }

    for (const p of [FALLBACK_TOKEN_PATH, SELECTED_ACCOUNT_PATH]) {
      if (fs.existsSync(p)) fs.unlinkSync(p);
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
    await this.saveTokenCache();
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
