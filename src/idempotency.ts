import { canonicalJson, planning, sha256Hex } from './plan.js';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { configDir, withFileLock, writeFileAtomic } from './auth.js';
import { CliError, ExitCode } from './errors.js';
import { formatOutput } from './formatter.js';
import type { GraphClient } from './graph-client.js';
import { info } from './utils.js';

/**
 * Idempotency-key ledger (spec §6 "Recover: retries", D23): `${configDir()}/idempotency.json` maps a caller's key
 * to the command it ran, a fingerprint of the request and a small result (ids/webLink), for 24 h. A retry with the
 * same key replays the stored result instead of sending again.
 */

export type IdempotencyResult = Record<string, unknown>;

/**
 * pending: claimed by the run in {pid, host, startedAt}. done: sent; `result` replays.
 * unknown: the send failed in a way that may still have reached Graph (CliError.maybeProcessed, R-7).
 */
export interface Entry {
  key: string;
  command: string;
  fingerprint: string;
  status: 'pending' | 'done' | 'unknown';
  result?: IdempotencyResult;
  ts: string;
  pid?: number;
  host?: string;
  startedAt?: string;
}

type Replay = { status: 'replay'; result: IdempotencyResult; ts: string };
export type KeyCheck = { status: 'new' } | Replay;

const TTL_MS = 24 * 60 * 60 * 1000;
/** A pending claim older than this is treated as abandoned (its owner may be on another host, or its pid reused). */
const PENDING_STALE_MS = 60 * 60 * 1000;

const ledgerPath = () => path.join(configDir(), 'idempotency.json');

/** Graph never sees dry-run requests, so the ledger must not either; GraphClient keeps the flag private. */
export function isDryRun(client: GraphClient): boolean {
  return client.isDryRun();
}

/** sha256 of the request body with object keys sorted, so equal requests always match. */
export function fingerprint(body: unknown): string {
  return sha256Hex(canonicalJson(body));
}

function readLedger(): Entry[] {
  let entries: Entry[];
  try {
    entries = JSON.parse(fs.readFileSync(ledgerPath(), 'utf8'));
  } catch {
    return [];
  }
  if (!Array.isArray(entries)) return [];
  const cutoff = Date.now() - TTL_MS;
  return entries.filter(e => Date.parse(e.ts) > cutoff);
}

function writeLedger(entries: Entry[]): void {
  writeFileAtomic(ledgerPath(), JSON.stringify(entries, null, 2));
}

/** Serialise read-modify-write of the ledger across processes (writeFileAtomic only locks the write itself). */
function withLedgerLock<T>(fn: () => T): T {
  fs.mkdirSync(configDir(), { recursive: true, mode: 0o700 });
  return withFileLock(`${ledgerPath()}.txn`, fn);
}

export function checkKey(key: string, command: string, fp: string): KeyCheck {
  const entry = readLedger().find(e => e.key === key);
  if (!entry) return { status: 'new' };
  if (entry.command !== command || entry.fingerprint !== fp) {
    throw new CliError(
      `Idempotency key "${key}" was already used at ${entry.ts} for a different request (${entry.command}).`,
      { code: 'idempotency_conflict', exitCode: ExitCode.USAGE, fix: 'Use a new --idempotency-key for a different request.' },
    );
  }
  if (entry.status === 'unknown' || (entry.status === 'pending' && abandoned(entry))) {
    const why = entry.status === 'unknown'
      ? 'failed in a way that may still have reached Microsoft 365'
      : 'was interrupted before it finished';
    const correlation: unknown = entry.result?.correlation;
    throw new CliError(
      `A request with idempotency key "${key}" (${entry.command}, started ${entry.startedAt ?? entry.ts}) ${why}; it may or may not have taken effect.`,
      {
        code: 'idempotency_outcome_unknown', exitCode: ExitCode.GENERIC,
        ...(typeof correlation === 'object' && correlation !== null ? { correlation: { ...correlation } } : {}),
        fix: `${entry.command.startsWith('mail ') ? `Run \`ms365 mail reconcile --idempotency-key ${key}\` (Sent Items, Drafts, Outbox)` : 'Check whether it took effect'}. If it did not, run \`ms365 idempotency release ${key}\` and retry; \`ms365 idempotency list\` shows the ledger.`,
      },
    );
  }
  if (entry.status === 'pending') {
    throw new CliError(
      `A request with idempotency key "${key}" started at ${entry.ts} and has not finished.`,
      {
        code: 'idempotency_in_progress', exitCode: ExitCode.GENERIC,
        fix: 'Wait for the other run to finish, then retry. If it crashed, check whether it took effect before using a new key.',
      },
    );
  }
  return { status: 'replay', result: entry.result ?? {}, ts: entry.ts };
}

/** The claim's owner is gone: a dead pid on this host, or a claim older than PENDING_STALE_MS (R-6). */
function abandoned(entry: Entry): boolean {
  const started = Date.parse(entry.startedAt ?? entry.ts);
  if (!(Date.now() - started < PENDING_STALE_MS)) return true;
  if (entry.host !== os.hostname() || !entry.pid) return false;
  try {
    process.kill(entry.pid, 0);
    return false;
  } catch (err) {
    return (err as NodeJS.ErrnoException).code === 'ESRCH';
  }
}

function putEntry(entry: Entry): void {
  writeLedger([...readLedger().filter(e => e.key !== entry.key), entry]);
}

export function recordKey(key: string, command: string, fp: string, result: IdempotencyResult): void {
  withLedgerLock(() => putEntry({ key, command, fingerprint: fp, status: 'done', result, ts: new Date().toISOString() }));
}

/** The live ledger entries (expired ones dropped), for `ms365 idempotency list`. */
export function listEntries(): Entry[] {
  return readLedger();
}

/** Remove one key from the ledger; false when it is not there. Never called automatically (F17). */
export function releaseKey(key: string): boolean {
  return withLedgerLock(() => {
    const entries = readLedger();
    const rest = entries.filter(e => e.key !== key);
    if (rest.length === entries.length) return false;
    writeLedger(rest);
    return true;
  });
}

/**
 * Check the key and claim it with a 'pending' entry in one locked step, run `fn`, then store its result.
 * A concurrent caller with the same key sees 'pending' and fails with idempotency_in_progress, or with
 * idempotency_outcome_unknown once the owner is gone. If `fn` throws, the claim is dropped so the key can be
 * retried, unless the error says the request may have been processed: then the entry stays as 'unknown' (R-7).
 */
export async function withIdempotency(
  key: string, command: string, fp: string, fn: () => Promise<IdempotencyResult>,
): Promise<Replay | { status: 'done'; result: IdempotencyResult }> {
  const check = withLedgerLock(() => {
    const c = checkKey(key, command, fp);
    if (c.status === 'new') {
      const ts = new Date().toISOString();
      putEntry({ key, command, fingerprint: fp, status: 'pending', ts, pid: process.pid, host: os.hostname(), startedAt: ts });
    }
    return c;
  });
  if (check.status === 'replay') return check;
  let result: IdempotencyResult;
  try {
    result = await fn();
  } catch (err) {
    withLedgerLock(() => {
      const entries = readLedger();
      const claim = entries.find(e => e.key === key);
      const rest = entries.filter(e => e.key !== key);
      // A lost send response keeps the claim as 'unknown', with the handles to reconcile it (AH-6(3)).
      writeLedger(err instanceof CliError && err.maybeProcessed && claim
        ? [...rest, { ...claim, status: 'unknown', ...(err.correlation && { result: { correlation: err.correlation } }) }]
        : rest);
    });
    throw err;
  }
  recordKey(key, command, fp, result);
  return { status: 'done', result };
}

export const IDEMPOTENCY_HELP = 'Retry-safe: if this key was used for the same request in the last 24 h, do not send again';

/**
 * Run `send` under --idempotency-key (D23) and return what it returned (stored as the key's result; default
 * `{sentAt}`). Returns null when the key replays an earlier send: the stored result is printed and Graph is not
 * called. Dry-run and --plan never touch the ledger.
 */
export async function idempotentSend(
  client: GraphClient, key: string | undefined, command: string, request: unknown, send: () => Promise<IdempotencyResult | void>,
): Promise<IdempotencyResult | null> {
  if (!key || isDryRun(client) || planning()) return (await send()) ?? {};
  const outcome = await withIdempotency(key, command, fingerprint(request), async () => (await send()) ?? { sentAt: new Date().toISOString() });
  if (outcome.status === 'done') return outcome.result;
  info(`Already sent at ${outcome.ts}; not sent again.`);
  console.log(formatOutput(outcome.result, 'json'));
  return null;
}

/**
 * Reconciliation found the send (AH-6(4)): an `unknown` entry becomes `done`, and the evidence joins its result.
 * Returns the entry's status after the update, or null when there is no such entry.
 */
export function settleKey(key: string, evidence: Record<string, unknown>): Entry['status'] | null {
  return withLedgerLock(() => {
    const entry = readLedger().find(e => e.key === key);
    if (!entry) return null;
    if (entry.status === 'pending') return entry.status;
    putEntry({ ...entry, status: 'done', result: { ...entry.result, reconciled: evidence } });
    return 'done';
  });
}
