import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const SAVED = ['HOME', 'DISPLAY', 'DBUS_SESSION_BUS_ADDRESS', 'MS365_CLI_CLIENT_ID'] as const;

describe('AuthManager token cache persistence (headless Linux)', () => {
  let saved: Record<string, string | undefined>;
  let home: string;

  beforeEach(() => {
    saved = Object.fromEntries(SAVED.map((k) => [k, process.env[k]]));
    home = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-home-'));
    process.env.HOME = home;
    process.env.MS365_CLI_CLIENT_ID = '00000000-0000-0000-0000-000000000000';
    delete process.env.DISPLAY;
    delete process.env.DBUS_SESSION_BUS_ADDRESS;
    vi.resetModules(); // auth.ts reads HOME and headless state at import time
  });

  afterEach(() => {
    for (const k of SAVED) {
      if (saved[k] === undefined) delete process.env[k];
      else process.env[k] = saved[k];
    }
  });

  it.runIf(process.platform === 'linux')('writes the cache and selected account to files with mode 0600', async () => {
    // Dynamic import: auth.ts captures HOME and headless state at module load, so it must load after env setup.
    const { AuthManager } = await import('../src/auth.js');
    // Private persistence methods are the unit under test; no public API writes the cache without a real login.
    const auth = new AuthManager(['User.Read']) as unknown as {
      saveTokenCache(): Promise<void>;
      saveSelectedAccount(): Promise<void>;
      selectedAccountId: string | null;
    };
    await auth.saveTokenCache();
    auth.selectedAccountId = 'acct-1';
    await auth.saveSelectedAccount();

    const dir = path.join(home, '.config', 'ms365-cli');
    for (const file of ['token-cache.json', 'selected-account.json']) {
      const st = fs.statSync(path.join(dir, file));
      expect(st.mode & 0o777).toBe(0o600);
    }
    expect(JSON.parse(fs.readFileSync(path.join(dir, 'selected-account.json'), 'utf8')).accountId).toBe('acct-1');
  });
});
