import fs from 'fs';
import os from 'os';
import path from 'path';
import type { KeychainBackend } from '../../src/keychain.js';

const STUBS = path.resolve(__dirname, '..', 'fixtures', 'keychain-stubs');

export interface KeychainStubs {
  backend: KeychainBackend;
  /** Where the stubs keep their items and `argv.log`. */
  db: string;
  bin: string;
  /** Env for a process using the stubbed backend (MS365_CLI_STORAGE=keychain, stub dir first on PATH). */
  env: Record<string, string>;
  /** Every argv the stubs were called with, one line per call. */
  argv(): string;
  /** The stubbed items, account → value, of the given service. */
  items(service: string): Record<string, string>;
}

/**
 * Stub `security`, `secret-tool` and `powershell`/`pwsh` scripts in a temp dir (ID-14). The real tools are never
 * reachable from a stubbed name because the stub dir comes first on PATH and every name is stubbed.
 */
export function keychainStubs(backend: KeychainBackend, extra: Record<string, string> = {}): KeychainStubs {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-kc-'));
  const bin = path.join(root, 'bin');
  const db = path.join(root, 'db');
  fs.mkdirSync(bin);
  fs.mkdirSync(db);
  for (const [name, src] of [['security', 'security'], ['secret-tool', 'secret-tool'], ['powershell', 'powershell'], ['pwsh', 'powershell']]) {
    fs.copyFileSync(path.join(STUBS, src), path.join(bin, name));
    fs.chmodSync(path.join(bin, name), 0o755);
  }
  const env: Record<string, string> = {
    MS365_CLI_STORAGE: 'keychain',
    MS365_CLI_KEYCHAIN_BACKEND: backend,
    PATH: `${bin}${path.delimiter}${process.env.PATH ?? ''}`,
    STUB_DB: db,
    STUB_MODE: '',
    DBUS_SESSION_BUS_ADDRESS: 'unix:path=/nonexistent/stub-bus',
    ...extra,
  };
  const sub = backend === 'macos' ? 'macos' : backend === 'libsecret' ? 'libsecret' : null;
  return {
    backend, db, bin, env,
    argv: () => { try { return fs.readFileSync(path.join(db, 'argv.log'), 'utf8'); } catch { return ''; } },
    items: (service: string) => {
      if (!sub) return {};
      const d = path.join(db, sub, service);
      let names: string[] = [];
      try { names = fs.readdirSync(d).filter(n => !n.includes('.tmp.')); } catch { /* none */ }
      return Object.fromEntries(names.map(n => [n, fs.readFileSync(path.join(d, n), 'utf8')]));
    },
  };
}
