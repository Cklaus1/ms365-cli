import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import { configDir, withFileLock, writeFileAtomic } from './auth.js';
import { CliError, ExitCode } from './errors.js';
import type { GraphClient } from './graph-client.js';

/**
 * Idempotency-key ledger (spec §6 "Recover: retries", D23): `${configDir()}/idempotency.json` maps a caller's key
 * to the command it ran, a fingerprint of the request and a small result (ids/webLink), for 24 h. A retry with the
 * same key replays the stored result instead of sending again.
 */

export type IdempotencyResult = Record<string, unknown>;

interface Entry {
  key: string;
  command: string;
  fingerprint: string;
  status: 'pending' | 'done';
  result?: IdempotencyResult;
  ts: string;
}

type Replay = { status: 'replay'; result: IdempotencyResult; ts: string };
export type KeyCheck = { status: 'new' } | Replay;

const TTL_MS = 24 * 60 * 60 * 1000;

const ledgerPath = () => path.join(configDir(), 'idempotency.json');

/** Graph never sees dry-run requests, so the ledger must not either; GraphClient keeps the flag private. */
export function isDryRun(client: GraphClient): boolean {
  return client.isDryRun();
}

function canonical(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonical);
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.keys(value).sort().map(k => [k, canonical((value as Record<string, unknown>)[k])]));
  }
  return value;
}

/** sha256 of the request body with object keys sorted, so equal requests always match. */
export function fingerprint(body: unknown): string {
  return crypto.createHash('sha256').update(JSON.stringify(canonical(body))).digest('hex');
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

function putEntry(entry: Entry): void {
  writeLedger([...readLedger().filter(e => e.key !== entry.key), entry]);
}

export function recordKey(key: string, command: string, fp: string, result: IdempotencyResult): void {
  withLedgerLock(() => putEntry({ key, command, fingerprint: fp, status: 'done', result, ts: new Date().toISOString() }));
}

/**
 * Check the key and claim it with a 'pending' entry in one locked step, run `fn`, then store its result.
 * A concurrent caller with the same key sees 'pending' and fails with idempotency_in_progress; if `fn` throws,
 * the claim is dropped so the key can be retried.
 */
export async function withIdempotency(
  key: string, command: string, fp: string, fn: () => Promise<IdempotencyResult>,
): Promise<Replay | { status: 'done'; result: IdempotencyResult }> {
  const check = withLedgerLock(() => {
    const c = checkKey(key, command, fp);
    if (c.status === 'new') putEntry({ key, command, fingerprint: fp, status: 'pending', ts: new Date().toISOString() });
    return c;
  });
  if (check.status === 'replay') return check;
  let result: IdempotencyResult;
  try {
    result = await fn();
  } catch (err) {
    withLedgerLock(() => writeLedger(readLedger().filter(e => e.key !== key)));
    throw err;
  }
  recordKey(key, command, fp, result);
  return { status: 'done', result };
}
