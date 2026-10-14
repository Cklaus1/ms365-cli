import fs from 'fs';
import path from 'path';
import { configDir } from './auth.js';

/** Rotate the log to audit.log.1 once it grows past this size (one generation is kept). */
export const AUDIT_MAX_BYTES = 5 * 1024 * 1024;

export interface AuditEntry {
  ts: string;
  actor: string;
  account?: string;
  command: string;
  method: string;
  endpoint: string;
  status: number;
  requestId?: string;
  ok: boolean;
}

let actor: string | undefined;
let commandPath: string | undefined;

/** Who is acting (global --actor flag); falls back to MS365_CLI_ACTOR, then 'cli'. */
export function setActor(name: string | undefined): void { actor = name || undefined; }

/** The resolved command path, e.g. "mail send" (set from the preAction hook). */
export function setCommandPath(name: string | undefined): void { commandPath = name || undefined; }

export function auditLogPath(): string {
  return path.join(configDir(), 'audit.log');
}

/**
 * The command as "path --opt1 --opt2": option names only, never option values or
 * positional arguments (they can hold message bodies, addresses or secrets).
 */
function describeCommand(): string {
  const args = process.argv.slice(2);
  const options = args.filter(a => /^--?[A-Za-z]/.test(a)).map(a => a.split('=')[0]);
  const firstOption = args.findIndex(a => a.startsWith('-'));
  const words = commandPath ?? (firstOption === -1 ? args : args.slice(0, firstOption))
    .filter(a => /^[a-z][a-z0-9-]*$/.test(a)).join(' ');
  return [words, ...options].filter(Boolean).join(' ');
}

/** Append one line to the audit log. Disabled by MS365_CLI_AUDIT=0. Never throws. */
export function recordAudit(entry: Omit<AuditEntry, 'ts' | 'actor' | 'command'>): void {
  if (process.env.MS365_CLI_AUDIT === '0') return;
  // Test runs must never write to the real ~/.config; they opt in via MS365_CLI_CONFIG_DIR.
  try {
    const file = auditLogPath();
    fs.mkdirSync(path.dirname(file), { recursive: true, mode: 0o700 });
    try {
      if (fs.statSync(file).size > AUDIT_MAX_BYTES) fs.renameSync(file, `${file}.1`);
    } catch { /* no log yet */ }
    const line: AuditEntry = {
      ts: new Date().toISOString(),
      actor: actor ?? (process.env.MS365_CLI_ACTOR || 'cli'),
      command: describeCommand(),
      ...entry,
      endpoint: entry.endpoint.split('?')[0],
    };
    fs.appendFileSync(file, JSON.stringify(line) + '\n', { mode: 0o600 });
  } catch {
    // Auditing must never break a command.
  }
}

/** Parse the last `n` entries of the log (malformed lines are skipped). */
export function readAuditTail(n: number): AuditEntry[] {
  let text: string;
  try {
    text = fs.readFileSync(auditLogPath(), 'utf8');
  } catch {
    return [];
  }
  return parseLines(text).slice(-n);
}

export function parseLines(text: string): AuditEntry[] {
  const out: AuditEntry[] = [];
  for (const line of text.split('\n')) {
    if (!line.trim()) continue;
    try { out.push(JSON.parse(line)); } catch { /* partial or corrupt line */ }
  }
  return out;
}
