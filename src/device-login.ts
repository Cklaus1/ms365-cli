import { spawn } from 'child_process';
import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import { configDir, writeFileAtomic, type AuthManager } from './auth.js';
import { childEnv } from './child-env.js';
import { CliError, ExitCode } from './errors.js';

/**
 * AH-8(5), OQ-12 (a): a device-code sign-in a host can relay (an MCP host without a token broker). `auth login
 * --start` writes a handle file, spawns a detached `auth login --hold <handle>` that holds MSAL's device-code call,
 * and returns once that child has the code. The child records its progress in the handle file (0600, never a
 * token); `auth login --wait <handle>` reads it until the sign-in ends.
 */
export type DeviceLoginState = 'starting' | 'pending' | 'done' | 'expired' | 'denied' | 'failed';

export interface DeviceLoginRecord {
  state: DeviceLoginState;
  /** The process holding the device-code call (0 until it starts). */
  pid: number;
  /** What the user signs in to: shown with the code, so a relayed code is recognisable. */
  app: string;
  userCode?: string;
  verificationUri?: string;
  expiresAt?: string;
  reason?: string;
}

export interface DeviceLoginStart { userCode: string; verificationUri: string; expiresAt: string; handle: string; app: string }

export interface DeviceLoginResult { state: DeviceLoginState; handle: string; expiresAt?: string; reason?: string }

/** How long `--start` waits for the child to get a code from Microsoft. */
export const START_TIMEOUT_MS = 30_000;
const START_POLL_MS = 200;
const WAIT_POLL_MS = 1000;

/** Test seams: the detached child, the clock and the poll sleep. */
export const deviceLoginDeps = {
  spawnHolder(argv: string[]): void {
    const child = spawn(process.execPath, [...process.execArgv, process.argv[1], ...argv], { detached: true, stdio: 'ignore', env: childEnv() });
    child.unref();
  },
  sleep: (ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms)),
  now: () => Date.now(),
};

const unknownHandle = (handle: string) => new CliError(`No device-code sign-in with handle "${handle}" (it ended, or was never started here).`,
  { code: 'unknown_handle', exitCode: ExitCode.USAGE, fix: 'ms365 auth login --start' });

function handleFile(handle: string): string {
  // A handle is our own random hex, so it can never name a path outside the config dir.
  if (!/^[0-9a-f]{32}$/.test(handle)) throw unknownHandle(handle);
  return path.join(configDir(), `device-login-${handle}.json`);
}

const writeRecord = (handle: string, record: DeviceLoginRecord) => writeFileAtomic(handleFile(handle), JSON.stringify(record));

function readRecord(handle: string): DeviceLoginRecord {
  const file = handleFile(handle);
  let text: string;
  try {
    text = fs.readFileSync(file, 'utf8');
  } catch {
    throw unknownHandle(handle);
  }
  const record: unknown = JSON.parse(text);
  if (typeof record !== 'object' || record === null || typeof Reflect.get(record, 'state') !== 'string') throw unknownHandle(handle);
  return record as DeviceLoginRecord;
}

function alive(pid: number): boolean {
  if (!pid) return true;
  try {
    process.kill(pid, 0);
    return true;
  } catch (err) {
    return err instanceof Error && Reflect.get(err, 'code') === 'EPERM';
  }
}

/** `auth login --start`: spawn the holder and return its code, or exit 3 `login_failed`. */
export async function startDeviceLogin(app: string, globals: string[]): Promise<DeviceLoginStart> {
  const handle = crypto.randomBytes(16).toString('hex');
  writeRecord(handle, { state: 'starting', pid: 0, app });
  deviceLoginDeps.spawnHolder([...globals, 'auth', 'login', '--hold', handle]);
  const deadline = deviceLoginDeps.now() + START_TIMEOUT_MS;
  for (;;) {
    const r = readRecord(handle);
    if (r.state === 'pending' && r.userCode && r.verificationUri && r.expiresAt) {
      return { userCode: r.userCode, verificationUri: r.verificationUri, expiresAt: r.expiresAt, handle, app };
    }
    const failed = r.state !== 'starting' && r.state !== 'pending';
    if (failed || deviceLoginDeps.now() >= deadline) {
      fs.rmSync(handleFile(handle), { force: true });
      const why = failed ? r.reason ?? r.state : `no code from Microsoft within ${START_TIMEOUT_MS / 1000} s`;
      throw new CliError(`The device-code sign-in could not start: ${why}.`, { code: 'login_failed', exitCode: ExitCode.AUTH });
    }
    await deviceLoginDeps.sleep(START_POLL_MS);
  }
}

/** The outcome of a failed MSAL device-code call. */
function failedState(err: unknown): DeviceLoginState {
  const text = err instanceof Error ? `${String(Reflect.get(err, 'errorCode') ?? '')} ${err.message}` : String(err);
  if (/expired/i.test(text)) return 'expired';
  if (/declin|denied/i.test(text)) return 'denied';
  return 'failed';
}

/** `auth login --hold <handle>` (the detached child): hold the device-code call and record each step. */
export async function holdDeviceLogin(auth: AuthManager, handle: string): Promise<void> {
  const record: DeviceLoginRecord = { ...readRecord(handle), pid: process.pid };
  const update = (change: Partial<DeviceLoginRecord>) => writeRecord(handle, Object.assign(record, change));
  update({});
  try {
    await auth.initialize();
    await auth.login({
      onDeviceCode: (c) => update({
        state: 'pending', userCode: c.userCode, verificationUri: c.verificationUri,
        expiresAt: new Date(deviceLoginDeps.now() + c.expiresIn * 1000).toISOString(),
      }),
    });
    update({ state: 'done' });
  } catch (err) {
    update({ state: failedState(err), reason: err instanceof Error ? err.message : String(err) });
  }
}

/**
 * `auth login --wait <handle> [--timeout s]`: poll the handle file until the sign-in ends (at the latest when its
 * code expires) or the timeout passes (`pending`). A finished sign-in's handle file is removed.
 */
export async function waitDeviceLogin(handle: string, timeoutMs?: number): Promise<DeviceLoginResult> {
  const deadline = timeoutMs === undefined ? Infinity : deviceLoginDeps.now() + timeoutMs;
  for (;;) {
    const r = readRecord(handle);
    let state = r.state;
    let reason = r.reason;
    const waiting = state === 'starting' || state === 'pending';
    if (waiting && r.expiresAt && Date.parse(r.expiresAt) <= deviceLoginDeps.now()) state = 'expired';
    else if (waiting && !alive(r.pid)) [state, reason] = ['failed', 'the sign-in process ended'];
    if (state !== 'starting' && state !== 'pending') {
      fs.rmSync(handleFile(handle), { force: true });
      return { state, handle, ...(reason && state !== 'done' ? { reason } : {}) };
    }
    const left = deadline - deviceLoginDeps.now();
    if (left <= 0) return { state: 'pending', handle, ...(r.expiresAt ? { expiresAt: r.expiresAt } : {}) };
    await deviceLoginDeps.sleep(Math.min(WAIT_POLL_MS, left));
  }
}
