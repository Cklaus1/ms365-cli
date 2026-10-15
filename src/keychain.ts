import { spawnSync } from 'child_process';
import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import { CliError, ExitCode } from './errors.js';
import { CorruptItemError, ensurePrivateMode, KEY_RE, TOKEN_CACHE_KEY, type CredentialStore } from './credential-store.js';

/**
 * OS keychain backends (identity-profiles spec ID-3). No native module ever loads in the CLI process (lesson 6):
 * every backend call is one child process run by `runTool` with a spawnSync timeout and SIGKILL, and secrets reach
 * the tool only through stdin — never argv, which `ps` shows to every user.
 */

export type KeychainBackend = 'macos' | 'libsecret' | 'dpapi';
export const KEYCHAIN_BACKENDS: readonly KeychainBackend[] = ['macos', 'libsecret', 'dpapi'];

/** Per-operation timeout; `MS365_CLI_KEYCHAIN_TIMEOUT_MS` overrides (tests). */
const DEFAULT_TIMEOUT_MS = 5_000;
/** The `auto` probe's budget (ID-4). */
export const PROBE_TIMEOUT_MS = 1_500;
/** E3 / ID-17: on Windows, `auto` uses the file store when PowerShell takes longer than this to start. */
export const POWERSHELL_START_LIMIT_MS = 400;
/**
 * E1 (OQ-1 → b): `auto` resolves to the file store on every platform until the keychain backends are verified on
 * real macOS/Windows (IDT-7); `keychain` is opt-in. Flipping this enables the probe path below.
 */
export const AUTO_PREFERS_KEYCHAIN = false;

export function keychainTimeoutMs(env: NodeJS.ProcessEnv = process.env): number {
  const n = Number(env.MS365_CLI_KEYCHAIN_TIMEOUT_MS);
  return Number.isFinite(n) && n > 0 ? n : DEFAULT_TIMEOUT_MS;
}

// ── The child runner ────────────────────────────────────────────────

export interface ToolResult {
  status: number | null;
  stdout: string;
  stderr: string;
  /** `timeout` (killed with SIGKILL), `missing` (ENOENT), or another spawn error code. */
  failure?: 'timeout' | 'missing' | string;
  ms: number;
}

export interface RunOptions {
  /** Written to the child's stdin (the only channel a secret may use). */
  input?: string;
  timeoutMs: number;
  env?: NodeJS.ProcessEnv;
}

/** The single way any backend runs its OS tool. */
export function runTool(file: string, args: string[], opts: RunOptions): ToolResult {
  const start = Date.now();
  const r = spawnSync(file, args, {
    input: opts.input ?? '',
    encoding: 'utf8',
    timeout: opts.timeoutMs,
    killSignal: 'SIGKILL',
    stdio: ['pipe', 'pipe', 'pipe'],
    windowsHide: true,
    maxBuffer: 16 * 1024 * 1024,
    env: opts.env ?? process.env,
  });
  const ms = Date.now() - start;
  const code = (r.error as NodeJS.ErrnoException | undefined)?.code;
  const failure = code === 'ETIMEDOUT' || (r.signal === 'SIGKILL' && ms >= opts.timeoutMs) ? 'timeout'
    : code === 'ENOENT' ? 'missing'
      : code;
  return { status: r.status, stdout: r.stdout ?? '', stderr: r.stderr ?? '', failure, ms };
}

/** Resolve a command on PATH without spawning anything (PATHEXT on Windows). */
export function whichSync(cmd: string, env: NodeJS.ProcessEnv = process.env): string | null {
  const exts = process.platform === 'win32' ? ['', ...(env.PATHEXT || '.EXE;.CMD;.BAT;.COM').split(';')] : [''];
  for (const dir of (env.PATH || env.Path || '').split(path.delimiter)) {
    if (!dir) continue;
    for (const ext of exts) {
      const candidate = path.join(dir, cmd + ext);
      try {
        const st = fs.statSync(candidate);
        if (st.isFile() && (process.platform === 'win32' || (st.mode & 0o111) !== 0)) return candidate;
      } catch { /* not here */ }
    }
  }
  return null;
}

// ── Errors ──────────────────────────────────────────────────────────

/** Exit 3 `keychain_unavailable` (ID-4): the keychain was chosen but cannot be used. */
export function keychainUnavailable(reason: string, backend?: KeychainBackend): CliError {
  const unlock = backend === 'macos' ? ', or unlock it: security unlock-keychain' : '';
  return new CliError(`The OS keychain is unavailable (${reason}).`, {
    code: 'keychain_unavailable', exitCode: ExitCode.AUTH,
    fix: `use the 0600 file instead: ms365 auth storage use file${unlock}`,
  });
}

function describeFailure(tool: string, r: ToolResult, timeoutMs: number): string {
  if (r.failure === 'timeout') return `${tool} did not answer within ${timeoutMs} ms and was killed`;
  if (r.failure === 'missing') return `${tool} not found`;
  if (r.failure) return `${tool} failed to start (${r.failure})`;
  const detail = r.stderr.trim().split('\n')[0];
  return `${tool} exited ${r.status}${detail ? `: ${detail}` : ''}`;
}

// ── Backends ────────────────────────────────────────────────────────

/** `ms365-cli:<first 8 hex of sha256(configDir)>`, so two config directories never share items (ID-3). */
export function serviceName(configDir: string): string {
  return `ms365-cli:${crypto.createHash('sha256').update(configDir).digest('hex').slice(0, 8)}`;
}

/** Which backend this platform uses; `MS365_CLI_KEYCHAIN_BACKEND` overrides (so all three run on Linux CI). */
export function platformBackend(env: NodeJS.ProcessEnv = process.env, platform: NodeJS.Platform = process.platform): KeychainBackend | null {
  const forced = env.MS365_CLI_KEYCHAIN_BACKEND;
  if (forced) {
    if ((KEYCHAIN_BACKENDS as readonly string[]).includes(forced)) return forced as KeychainBackend;
    throw new CliError(`Invalid MS365_CLI_KEYCHAIN_BACKEND "${forced}" (use ${KEYCHAIN_BACKENDS.join(', ')}).`, { code: 'usage', exitCode: ExitCode.USAGE });
  }
  return platform === 'darwin' ? 'macos' : platform === 'win32' ? 'dpapi' : platform === 'linux' ? 'libsecret' : null;
}

export interface KeychainStoreOptions {
  backend: KeychainBackend;
  configDir: string;
  env?: NodeJS.ProcessEnv;
  timeoutMs?: number;
  /** Atomic 0600 write (auth.ts `writeFileAtomic`), for DPAPI blob files. */
  writeFileAtomic: (target: string, data: string) => void;
}

export interface KeychainStore extends CredentialStore {
  readonly backend: KeychainBackend;
  /** Why `available()` is false (or undefined when it is true). */
  unavailableReason(): string | undefined;
}

const PREFIX = `${TOKEN_CACHE_KEY}.`;

/** Store key `token-cache.<acct>` ↔ keychain account attribute `<acct>` (the homeAccountId). */
function accountOf(key: string): string {
  if (!key.startsWith(PREFIX) || !KEY_RE.test(key)) throw new Error(`Invalid keychain key "${key}"`);
  return key.slice(PREFIX.length);
}

export function createKeychainStore(opts: KeychainStoreOptions): KeychainStore {
  switch (opts.backend) {
    case 'macos': return macosStore(opts);
    case 'libsecret': return libsecretStore(opts);
    case 'dpapi': return dpapiStore(opts);
  }
}

/** Shared scaffolding: resolve the tool lazily, run it, turn failures into exit-3 errors. */
function toolRunner(opts: KeychainStoreOptions, names: string[], extraCheck?: () => string | undefined) {
  const env = opts.env ?? process.env;
  const timeoutMs = opts.timeoutMs ?? keychainTimeoutMs(env);
  let resolved: string | null | undefined;
  const tool = () => {
    if (resolved === undefined) resolved = names.map(n => whichSync(n, env)).find(Boolean) ?? null;
    return resolved;
  };
  const unavailableReason = () => extraCheck?.() ?? (tool() ? undefined : `${names.join(' / ')} not found on PATH`);
  const run = (args: string[], input?: string) => {
    const reason = unavailableReason();
    if (reason) throw keychainUnavailable(reason, opts.backend);
    const r = runTool(tool()!, args, { input, timeoutMs, env });
    if (r.failure) throw keychainUnavailable(describeFailure(names[0], r, timeoutMs), opts.backend);
    return r;
  };
  const fail = (r: ToolResult): never => { throw keychainUnavailable(describeFailure(names[0], r, timeoutMs), opts.backend); };
  return { run, fail, unavailableReason, env };
}

const B64_RE = /^[A-Za-z0-9+/]*={0,2}$/;

/**
 * macOS login keychain via `/usr/bin/security`†. Writes go through `security -i`, whose commands are read from
 * stdin, so the (base64) value never appears in argv. `find-generic-password -w` prints it back as one line;
 * exit 44 = no such item.
 */
function macosStore(opts: KeychainStoreOptions): KeychainStore {
  const svc = serviceName(opts.configDir);
  const t = toolRunner(opts, ['security']);
  const NOT_FOUND = 44;
  return {
    kind: 'keychain', backend: 'macos',
    unavailableReason: t.unavailableReason,
    available: () => !t.unavailableReason(),
    get(key) {
      const r = t.run(['find-generic-password', '-s', svc, '-a', accountOf(key), '-w']);
      if (r.status === NOT_FOUND) return undefined;
      if (r.status !== 0) t.fail(r);
      const b64 = r.stdout.trim();
      if (!b64) return undefined;
      if (!B64_RE.test(b64)) throw new CorruptItemError('keychain item is not base64');
      return Buffer.from(b64, 'base64').toString('utf8') || undefined;
    },
    set(key, value) {
      const b64 = Buffer.from(value, 'utf8').toString('base64');
      // Account and service are [A-Za-z0-9._:-] only and the value is base64: no quoting needed.
      const r = t.run(['-i'], `add-generic-password -U -a ${accountOf(key)} -s ${svc} -w ${b64}\n`);
      if (r.status !== 0 || /error|could not|SecKeychain/i.test(r.stderr)) t.fail(r);
    },
    delete(key) {
      const r = t.run(['delete-generic-password', '-s', svc, '-a', accountOf(key)]);
      if (r.status !== 0 && r.status !== NOT_FOUND) t.fail(r);
    },
    list() {
      // dump-keychain without -d prints attributes only (no secrets).
      const r = t.run(['dump-keychain']);
      if (r.status !== 0) t.fail(r);
      const keys: string[] = [];
      for (const block of r.stdout.split(/^keychain: /m)) {
        const svce = /"svce"<blob>="([^"]*)"/.exec(block)?.[1];
        const acct = /"acct"<blob>="([^"]*)"/.exec(block)?.[1];
        if (svce === svc && acct && KEY_RE.test(PREFIX + acct)) keys.push(PREFIX + acct);
      }
      return [...new Set(keys)].sort();
    },
  };
}

/**
 * Linux Secret Service via `secret-tool` (libsecret). `store` reads the secret from stdin until EOF; only used
 * when a D-Bus session exists (without one libsecret tries to autolaunch and can hang — lesson 6).
 */
function libsecretStore(opts: KeychainStoreOptions): KeychainStore {
  const svc = serviceName(opts.configDir);
  const t = toolRunner(opts, ['secret-tool'], () =>
    (opts.env ?? process.env).DBUS_SESSION_BUS_ADDRESS ? undefined : 'no D-Bus session (DBUS_SESSION_BUS_ADDRESS is not set)');
  const attrs = (key: string) => ['service', svc, 'account', accountOf(key)];
  return {
    kind: 'keychain', backend: 'libsecret',
    unavailableReason: t.unavailableReason,
    available: () => !t.unavailableReason(),
    get(key) {
      const r = t.run(['lookup', ...attrs(key)]);
      // secret-tool exits 1 with no output when nothing matches (an error also exits 1, but says why on stderr).
      if (r.status === 1 && !r.stdout && !r.stderr.trim()) return undefined;
      if (r.status !== 0) t.fail(r);
      return r.stdout.replace(/\n$/, '') || undefined;
    },
    set(key, value) {
      const r = t.run(['store', `--label=ms365-cli token cache (${svc})`, ...attrs(key)], value);
      if (r.status !== 0) t.fail(r);
    },
    delete(key) {
      const r = t.run(['clear', ...attrs(key)]);
      if (r.status !== 0 && !(r.status === 1 && !r.stderr.trim())) t.fail(r);
    },
    list() {
      const r = t.run(['search', '--all', 'service', svc]);
      if (r.status !== 0 && !(r.status === 1 && !r.stdout && !r.stderr.trim())) t.fail(r);
      const keys = [...`${r.stdout}\n${r.stderr}`.matchAll(/^attribute\.account = (.+)$/gm)]
        .map(m => PREFIX + m[1].trim()).filter(k => KEY_RE.test(k));
      return [...new Set(keys)].sort();
    },
  };
}

/** PowerShell for DPAPI (E2): base64 plaintext/blob in on stdin, base64 out on stdout; exit 3 = cannot decrypt. */
const DPAPI_SCRIPT = (op: 'Protect' | 'Unprotect') => [
  "$ErrorActionPreference = 'Stop'",
  'try { Add-Type -AssemblyName System.Security } catch { }',
  '$in = [Console]::In.ReadToEnd().Trim()',
  'try {',
  '  $bytes = [Convert]::FromBase64String($in)',
  `  $out = [Security.Cryptography.ProtectedData]::${op}($bytes, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)`,
  '} catch { [Console]::Error.WriteLine($_.Exception.Message); exit 3 }',
  '[Console]::Out.Write([Convert]::ToBase64String($out))',
].join('\n');

/** `-EncodedCommand` takes base64 of the UTF-16LE script; it carries no secret. */
export function encodePowerShell(script: string): string {
  return Buffer.from(script, 'utf16le').toString('base64');
}

const PS_ARGS = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand'];
const DPAPI_CORRUPT = 3;

/**
 * Windows DPAPI (`ProtectedData`, CurrentUser) through PowerShell (`pwsh` preferred, else `powershell`)†. The
 * encrypted blob lives in `configDir/credentials/<acct>.dpapi`; reading it back costs one PowerShell start (ID-17).
 */
function dpapiStore(opts: KeychainStoreOptions): KeychainStore & { measureStartMs(): number } {
  const t = toolRunner(opts, ['pwsh', 'powershell']);
  const dir = path.join(opts.configDir, 'credentials');
  const blob = (key: string) => path.join(dir, `${accountOf(key)}.dpapi`);
  const ps = (op: 'Protect' | 'Unprotect', input: string) => t.run([...PS_ARGS, encodePowerShell(DPAPI_SCRIPT(op))], input);
  return {
    kind: 'keychain', backend: 'dpapi',
    unavailableReason: t.unavailableReason,
    available: () => !t.unavailableReason(),
    get(key) {
      let data: string;
      try {
        ensurePrivateMode(blob(key));
        data = fs.readFileSync(blob(key), 'utf8').trim();
      } catch { return undefined; }
      if (!data) return undefined;
      const r = ps('Unprotect', data);
      if (r.status === DPAPI_CORRUPT) throw new CorruptItemError('DPAPI blob does not decrypt for this user');
      if (r.status !== 0) t.fail(r);
      const out = r.stdout.trim();
      if (!B64_RE.test(out)) throw new CorruptItemError('DPAPI output is not base64');
      return Buffer.from(out, 'base64').toString('utf8') || undefined;
    },
    set(key, value) {
      const r = ps('Protect', Buffer.from(value, 'utf8').toString('base64'));
      if (r.status !== 0 || !B64_RE.test(r.stdout.trim()) || !r.stdout.trim()) t.fail(r);
      fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
      if (process.platform !== 'win32') fs.chmodSync(dir, 0o700);
      opts.writeFileAtomic(blob(key), r.stdout.trim());
    },
    delete(key) {
      try { fs.unlinkSync(blob(key)); } catch (err) {
        if ((err as NodeJS.ErrnoException).code !== 'ENOENT') throw err;
      }
    },
    list() {
      let names: string[];
      try { names = fs.readdirSync(dir); } catch { return []; }
      return names.filter(n => n.endsWith('.dpapi')).map(n => PREFIX + n.slice(0, -'.dpapi'.length)).filter(k => KEY_RE.test(k)).sort();
    },
    /** E3 hook: wall time of one PowerShell start (`exit 0`). */
    measureStartMs() {
      const r = t.run([...PS_ARGS, encodePowerShell('exit 0')]);
      return r.ms;
    },
  };
}

// ── storage.json and the storage choice (ID-4) ──────────────────────

export type StorageSetting = 'auto' | 'keychain' | 'file';
const SETTINGS: readonly StorageSetting[] = ['auto', 'keychain', 'file'];

export interface StorageConfig {
  storage?: StorageSetting;
  /** What `auto` resolved to (sticky until `auth storage use auto` / a re-probe). */
  resolved?: 'keychain' | 'file';
  backend?: KeychainBackend;
  probedAt?: string;
  reason?: string;
  /** The fallback notice has been printed once for this config directory. */
  noticeShown?: boolean;
}

export const storageConfigPath = (configDir: string) => path.join(configDir, 'storage.json');

export function readStorageConfig(configDir: string): StorageConfig {
  const file = storageConfigPath(configDir);
  ensurePrivateMode(file);
  try {
    const parsed = JSON.parse(fs.readFileSync(file, 'utf8'));
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
  } catch {
    return {};
  }
}

export function writeStorageConfig(configDir: string, cfg: StorageConfig, writeFileAtomic: (t: string, d: string) => void): void {
  writeFileAtomic(storageConfigPath(configDir), JSON.stringify(cfg, null, 2) + '\n');
}

/** The configured setting: `MS365_CLI_STORAGE` wins (CI, tests), then storage.json, then `auto`. */
export function configuredStorage(configDir: string, env: NodeJS.ProcessEnv = process.env): { setting: StorageSetting; source: 'env' | 'file' | 'default' } {
  const fromEnv = env.MS365_CLI_STORAGE;
  if (fromEnv) {
    if (!(SETTINGS as readonly string[]).includes(fromEnv)) {
      throw new CliError(`Invalid MS365_CLI_STORAGE "${fromEnv}" (use auto, keychain or file).`, { code: 'usage', exitCode: ExitCode.USAGE });
    }
    return { setting: fromEnv as StorageSetting, source: 'env' };
  }
  const fromFile = readStorageConfig(configDir).storage;
  if (fromFile && (SETTINGS as readonly string[]).includes(fromFile)) return { setting: fromFile, source: 'file' };
  return { setting: 'auto', source: 'default' };
}

export interface StorageResolution {
  kind: 'file' | 'keychain';
  /** The item store when kind is `keychain`. */
  store?: KeychainStore;
  reason?: string;
}

export interface ResolveOptions {
  configDir: string;
  /** `read` never probes; `write` resolves `auto` once (sticky) on the first write. */
  phase: 'read' | 'write';
  env?: NodeJS.ProcessEnv;
  platform?: NodeJS.Platform;
  writeFileAtomic: (target: string, data: string) => void;
  /** Default AUTO_PREFERS_KEYCHAIN (E1: false). Tests set it to exercise the probe. */
  autoPrefersKeychain?: boolean;
  /** Where the one-time fallback notice goes (default stderr). */
  notice?: (line: string) => void;
}

/**
 * Decide where the MSAL cache's credentials live. `file`: the 0600 file. `keychain`: the platform backend, and
 * any failure is exit 3 `keychain_unavailable`. `auto`: E1 makes it `file` (nothing recorded); with the probe
 * enabled it resolves once, on the first write, records the outcome in storage.json and prints one notice per
 * config directory on fallback.
 */
export function resolveStorage(opts: ResolveOptions): StorageResolution {
  const env = opts.env ?? process.env;
  const { setting } = configuredStorage(opts.configDir, env);
  if (setting === 'file') return { kind: 'file' };

  const keychain = (timeoutMs?: number) => platformKeychain({ ...opts, env, timeoutMs });

  if (setting === 'keychain') {
    const kc = keychain();
    if (typeof kc === 'string') throw keychainUnavailable(kc, safeBackend(env, opts.platform));
    return { kind: 'keychain', store: kc };
  }

  // auto
  const cfg = readStorageConfig(opts.configDir);
  if (cfg.resolved === 'file') return { kind: 'file', reason: cfg.reason };
  if (cfg.resolved === 'keychain') {
    const kc = keychain();
    if (typeof kc === 'string') throw keychainUnavailable(`${kc}; this config directory keeps its credentials in the keychain`, safeBackend(env, opts.platform));
    return { kind: 'keychain', store: kc };
  }
  if (!(opts.autoPrefersKeychain ?? AUTO_PREFERS_KEYCHAIN)) {
    return { kind: 'file', reason: 'auto uses the file store until keychain storage is verified on this platform (opt in: ms365 auth storage use keychain)' };
  }
  if (opts.phase === 'read') return { kind: 'file', reason: 'not resolved yet' };

  const probeStore = keychain(Math.min(PROBE_TIMEOUT_MS, keychainTimeoutMs(env)));
  const probed = typeof probeStore === 'string' ? { ok: false, reason: probeStore } : probeKeychain(probeStore);
  const store = probed.ok ? keychain() : undefined;
  const next: StorageConfig = {
    ...cfg,
    resolved: typeof store === 'object' ? 'keychain' : 'file',
    backend: typeof store === 'object' ? store.backend : undefined,
    probedAt: new Date().toISOString(),
    reason: probed.reason,
  };
  if (next.resolved === 'file' && !cfg.noticeShown) {
    (opts.notice ?? (l => { process.stderr.write(l + '\n'); }))(
      `ms365: keychain unavailable (${probed.reason}); using the 0600 file — ms365 auth storage show`);
    next.noticeShown = true;
  }
  writeStorageConfig(opts.configDir, next, opts.writeFileAtomic);
  return typeof store === 'object' ? { kind: 'keychain', store, reason: probed.reason } : { kind: 'file', reason: probed.reason };
}

/**
 * This platform's keychain store, or why there is none (no backend for the platform, tool missing, no D-Bus).
 * Never spawns anything: availability is decided from PATH and the environment.
 */
export function platformKeychain(opts: {
  configDir: string; env?: NodeJS.ProcessEnv; platform?: NodeJS.Platform; timeoutMs?: number;
  writeFileAtomic: (target: string, data: string) => void;
}): KeychainStore | string {
  const env = opts.env ?? process.env;
  const backend = platformBackend(env, opts.platform);
  if (!backend) return `no supported keychain on ${opts.platform ?? process.platform}`;
  const store = createKeychainStore({ backend, configDir: opts.configDir, env, timeoutMs: opts.timeoutMs, writeFileAtomic: opts.writeFileAtomic });
  return store.unavailableReason() ?? store;
}

/** Backend names as `auth storage show` reports them (ID-15). */
export const BACKEND_LABEL: Record<KeychainBackend, string> = { macos: 'macos-keychain', libsecret: 'libsecret', dpapi: 'dpapi' };

export function safeBackend(env: NodeJS.ProcessEnv, platform?: NodeJS.Platform): KeychainBackend | undefined {
  try { return platformBackend(env, platform) ?? undefined; } catch { return undefined; }
}

/**
 * Write–read–delete of a probe item (call it on a store built with the PROBE_TIMEOUT_MS budget); DPAPI also checks
 * PowerShell's start time (E3). Used by the `auto` resolution and by `auth storage show` (ID-15).
 */
export function probeKeychain(kc: KeychainStore): { ok: boolean; reason: string } {
  const started = Date.now();
  try {
    if (kc.backend === 'dpapi') {
      const ms = (kc as KeychainStore & { measureStartMs(): number }).measureStartMs();
      if (ms > POWERSHELL_START_LIMIT_MS) return { ok: false, reason: `PowerShell took ${ms} ms to start (limit ${POWERSHELL_START_LIMIT_MS} ms)` };
    }
    const key = `${TOKEN_CACHE_KEY}.probe-${process.pid}-${crypto.randomBytes(4).toString('hex')}`;
    const value = crypto.randomBytes(12).toString('hex');
    kc.set(key, value);
    const back = kc.get(key);
    kc.delete(key);
    if (back !== value) return { ok: false, reason: 'the probe item did not read back' };
    return { ok: true, reason: `probe passed in ${Date.now() - started} ms` };
  } catch (err) {
    return { ok: false, reason: err instanceof CliError ? err.message.replace(/^The OS keychain is unavailable \((.*)\)\.$/, '$1') : String(err) };
  }
}
