import { Option } from 'commander';
import readline from 'readline';
import fs from 'fs';
import path from 'path';
import { CliError, ConfirmationRefusedError, ExitCode } from './errors.js';

const OUTPUT_FORMATS = ['json', 'ndjson', 'table', 'csv', 'text', 'yaml'] as const;

/** Default output format: table for interactive terminals, json for pipes/scripts. */
export const defaultFormat = process.stdout.isTTY ? 'table' : 'json';

/** Commander Option for -o/--output with validated choices and TTY-aware default. */
export function outputOption(overrideDefault?: string): Option {
  return new Option('-o, --output <format>', 'Output format: json, ndjson, table, csv, text, yaml')
    .choices([...OUTPUT_FORMATS])
    .default(overrideDefault ?? defaultFormat);
}

/** Global quiet flag — suppressed confirmation messages when true. */
export let quiet = false;
export function setQuiet(val: boolean): void { quiet = val; }

/** Print a non-data message (e.g. "Email sent.") unless --quiet is active. */
export function info(msg: string): void {
  if (!quiet) console.error(msg);
}

/**
 * Prompt user for confirmation on destructive operations.
 * Always prompts — no --yes override — so AI agents can't bypass.
 * Non-TTY stdin auto-rejects to prevent accidental deletions in scripts.
 * Resolves true when confirmed; throws ConfirmationRefusedError (exit 7) otherwise.
 */
export async function confirm(message: string): Promise<boolean> {
  if (!process.stdin.isTTY) {
    throw new ConfirmationRefusedError('Destructive operation requires interactive confirmation. Run in a terminal.');
  }
  const rl = readline.createInterface({ input: process.stdin, output: process.stderr });
  const answer = await new Promise<string>((resolve) => {
    rl.question(`${message} [y/N] `, (a) => {
      rl.close();
      resolve(a);
    });
  });
  if (answer.trim().toLowerCase() !== 'y') {
    throw new ConfirmationRefusedError('Cancelled.');
  }
  return true;
}

/** Validate and return a YYYY-MM-DD date string. */
export function parseDate(input: string): string {
  const d = new Date(input + 'T00:00:00Z');
  if (isNaN(d.getTime())) {
    throw new Error(`Invalid date: "${input}". Use YYYY-MM-DD format.`);
  }
  return input;
}

/** Truncate a string with an ellipsis if it exceeds max length. */
export function truncate(str: string, max: number): string {
  return str.length > max ? str.substring(0, max - 1) + '…' : str;
}

/** Validate a single email address. */
export function validateEmail(email: string): string {
  const trimmed = email.trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimmed)) {
    throw new Error(`Invalid email address: "${trimmed}"`);
  }
  return trimmed;
}

/** Parse a comma-separated list of emails into Graph API recipient format. */
export function parseRecipients(csv: string): Array<{ emailAddress: { address: string } }> {
  return csv.split(',').map(e => ({
    emailAddress: { address: validateEmail(e) },
  }));
}

/** Format a byte count as a human-readable size. */
export function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

/** Build OData query params from CLI options. Only includes non-undefined values. */
export function buildODataParams(
  opts: Record<string, string | undefined>,
  mapping: Record<string, string>,
): Record<string, string> {
  const params: Record<string, string> = {};
  for (const [optKey, paramKey] of Object.entries(mapping)) {
    const val = opts[optKey];
    if (val !== undefined) params[paramKey] = val;
  }
  return params;
}

// ── Paths & query quoting ─────────────────────────────────────────

/** Encode each segment of a drive/site path for use in a URL (keeps the slashes). */
export function encodePath(p: string): string {
  return p.split('/').map(encodeURIComponent).join('/');
}

/** Escape a value for use inside an OData single-quoted string literal. */
export function odataQuote(value: string): string {
  return value.replace(/'/g, "''");
}

/** Quote a term for a KQL $search expression ("…"), escaping embedded quotes. */
export function kqlQuote(term: string): string {
  return `"${term.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

/**
 * Resolve a safe local output path for a server-supplied file name: strips directories
 * and control characters (no traversal), and refuses to overwrite unless force is set.
 */
export function safeOutputPath(name: string, opts: { out?: string; force?: boolean; fallback?: string } = {}): string {
  let target: string;
  if (opts.out) {
    target = path.resolve(opts.out); // the user chose this path explicitly
  } else {
    // eslint-disable-next-line no-control-regex
    let base = path.basename(name.replace(/\\/g, '/')).replace(/[\x00-\x1f\x7f]/g, '').trim();
    if (!base || base === '.' || base === '..') base = opts.fallback ?? 'download';
    target = path.resolve(base);
  }
  if (!opts.force && fs.existsSync(target)) {
    throw new CliError(`Refusing to overwrite existing file: ${target}`, { code: 'file_exists', fix: 'pass --force to overwrite' });
  }
  return target;
}

// ── Time zones & times ────────────────────────────────────────────

/** The user's time zone: explicit override → MS365_CLI_TIMEZONE → system zone → UTC. */
export function resolveTimeZone(override?: string): string {
  const tz = override || process.env.MS365_CLI_TIMEZONE || Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
  try {
    new Intl.DateTimeFormat('en-US', { timeZone: tz });
  } catch {
    throw new CliError(`Unknown time zone "${tz}". Use an IANA name like America/Los_Angeles.`, { code: 'usage', exitCode: ExitCode.USAGE });
  }
  return tz;
}

/** Header asking Outlook to return dateTime values in `tz` instead of UTC. */
export function preferTimeZone(tz: string): Record<string, string> {
  return { Prefer: `outlook.timezone="${tz}"` };
}

interface WallParts { year: number; month: number; day: number; hour: number; minute: number; second: number }

function wallParts(date: Date, tz: string): WallParts {
  const fmt = new Intl.DateTimeFormat('en-US', {
    timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
  });
  const p = Object.fromEntries(fmt.formatToParts(date).map(x => [x.type, x.value]));
  return { year: +p.year, month: +p.month, day: +p.day, hour: +p.hour, minute: +p.minute, second: +p.second };
}

/** Convert a wall-clock time in `tz` to the instant it denotes (DST-aware). */
export function wallTimeToDate(p: WallParts, tz: string): Date {
  const asUtc = Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second);
  // Two passes handle offsets that change between the guess and the answer (DST edges).
  let guess = asUtc;
  for (let i = 0; i < 2; i++) {
    const w = wallParts(new Date(guess), tz);
    const shown = Date.UTC(w.year, w.month - 1, w.day, w.hour, w.minute, w.second);
    guess += asUtc - shown;
  }
  return new Date(guess);
}

const pad = (n: number, w = 2) => String(n).padStart(w, '0');

/** Format an instant as a Graph dateTime (no offset) in `tz`: 2026-10-14T09:00:00. */
export function toGraphDateTime(date: Date, tz: string): string {
  const w = wallParts(date, tz);
  return `${w.year}-${pad(w.month)}-${pad(w.day)}T${pad(w.hour)}:${pad(w.minute)}:${pad(w.second)}`;
}

/** Format an instant as ISO 8601 with the numeric offset of `tz`: 2026-10-14T09:00:00-07:00. */
export function toIsoWithOffset(date: Date, tz: string): string {
  const w = wallParts(date, tz);
  const wallMs = Date.UTC(w.year, w.month - 1, w.day, w.hour, w.minute, w.second);
  const offMin = Math.round((wallMs - Math.floor(date.getTime() / 1000) * 1000) / 60000);
  const sign = offMin >= 0 ? '+' : '-';
  const a = Math.abs(offMin);
  return `${toGraphDateTime(date, tz)}${sign}${pad(Math.floor(a / 60))}:${pad(a % 60)}`;
}

/**
 * Parse a Graph dateTimeTimeZone value. Graph omits the offset; without a Prefer header
 * the zone is UTC, and `new Date()` would wrongly read it as local time.
 */
export function parseGraphDateTime(value: { dateTime: string; timeZone?: string } | string, fallbackTz = 'UTC'): Date {
  const dateTime = typeof value === 'string' ? value : value.dateTime;
  const tz = (typeof value === 'string' ? undefined : value.timeZone) || fallbackTz;
  if (/[zZ]|[+-]\d\d:\d\d$/.test(dateTime)) return new Date(dateTime);
  const m = dateTime.match(/^(\d{4})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d)(?::(\d\d))?)?/);
  if (!m) return new Date(NaN);
  const parts = { year: +m[1], month: +m[2], day: +m[3], hour: +(m[4] ?? 0), minute: +(m[5] ?? 0), second: +(m[6] ?? 0) };
  if (tz === 'UTC' || tz === 'Etc/UTC') {
    return new Date(Date.UTC(parts.year, parts.month - 1, parts.day, parts.hour, parts.minute, parts.second));
  }
  try {
    return wallTimeToDate(parts, tz);
  } catch {
    // Windows zone names (e.g. "Pacific Standard Time") aren't IANA; fall back to UTC.
    return new Date(Date.UTC(parts.year, parts.month - 1, parts.day, parts.hour, parts.minute, parts.second));
  }
}

/** Start and end (exclusive) of the calendar day containing `date` in `tz`, as UTC ISO strings. */
export function dayRange(date: Date, tz: string, days = 1): { start: string; end: string } {
  const w = wallParts(date, tz);
  const start = wallTimeToDate({ ...w, hour: 0, minute: 0, second: 0 }, tz);
  const endWall = new Date(Date.UTC(w.year, w.month - 1, w.day + days));
  const end = wallTimeToDate({ year: endWall.getUTCFullYear(), month: endWall.getUTCMonth() + 1, day: endWall.getUTCDate(), hour: 0, minute: 0, second: 0 }, tz);
  return { start: start.toISOString(), end: end.toISOString() };
}

const WEEKDAYS = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];

export interface ParsedTime {
  /** The instant. */
  date: Date;
  /** Wall-clock time in the resolved zone, Graph format (no offset). */
  dateTime: string;
  timeZone: string;
  /** True when the input named a day but no time ("tomorrow", "2026-10-20"). */
  dateOnly: boolean;
}

/**
 * Parse a user/agent time expression in `tz`:
 * ISO 8601 (with or without offset) · YYYY-MM-DD [HH:MM] · now · today/tomorrow/yesterday [HH:MM]
 * · [next] mon…sun [HH:MM] · HH:MM (today) · +N(m|h|d|w) / -N(…)
 */
export function parseTime(input: string, opts: { tz?: string; now?: Date } = {}): ParsedTime {
  const tz = resolveTimeZone(opts.tz);
  const now = opts.now ?? new Date();
  const s = input.trim().toLowerCase();
  const done = (date: Date, dateOnly = false): ParsedTime => ({ date, dateTime: toGraphDateTime(date, tz), timeZone: tz, dateOnly });
  const bad = () => new CliError(
    `Unrecognized time "${input}". Use ISO 8601 (2026-10-20T15:00), "tomorrow 15:00", "next mon 9:30", "+2h" or "+3d".`,
    { code: 'usage', exitCode: ExitCode.USAGE });

  // Absolute ISO with an explicit offset/Z
  if (/^\d{4}-\d\d-\d\dt\d\d:\d\d(:\d\d(\.\d+)?)?(z|[+-]\d\d:\d\d)$/.test(s)) {
    const d = new Date(input.trim());
    if (isNaN(d.getTime())) throw bad();
    return done(d);
  }

  const timeOfDay = (t?: string): { hour: number; minute: number } | null => {
    if (!t) return null;
    const m = t.match(/^(\d{1,2})(?::(\d\d))?\s*(am|pm)?$/);
    if (!m) throw bad();
    let hour = +m[1];
    const minute = +(m[2] ?? 0);
    if (m[3] === 'pm' && hour < 12) hour += 12;
    if (m[3] === 'am' && hour === 12) hour = 0;
    if (hour > 23 || minute > 59) throw bad();
    return { hour, minute };
  };
  const onDay = (base: Date, addDays: number, t?: string): ParsedTime => {
    const w = wallParts(base, tz);
    const d = new Date(Date.UTC(w.year, w.month - 1, w.day + addDays));
    const tod = timeOfDay(t);
    const date = wallTimeToDate({ year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate(),
      hour: tod?.hour ?? 0, minute: tod?.minute ?? 0, second: 0 }, tz);
    return done(date, !tod);
  };

  // Local ISO date or date-time (interpreted in tz)
  let m = s.match(/^(\d{4})-(\d\d)-(\d\d)(?:[t ](\d\d):(\d\d)(?::(\d\d))?)?$/);
  if (m) {
    const [year, month, day, hour, minute, second] = [+m[1], +m[2], +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0)];
    const daysInMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();
    if (month < 1 || month > 12 || day < 1 || day > daysInMonth || hour > 23 || minute > 59 || second > 59) throw bad();
    const date = wallTimeToDate({ year, month, day, hour, minute, second }, tz);
    return done(date, m[4] === undefined);
  }

  if (s === 'now') return done(now);

  m = s.match(/^([+-])(\d+)\s*(m|min|h|d|w)$/);
  if (m) {
    const n = (m[1] === '-' ? -1 : 1) * +m[2];
    const ms = { m: 60e3, min: 60e3, h: 3600e3, d: 86400e3, w: 7 * 86400e3 }[m[3] as 'm' | 'min' | 'h' | 'd' | 'w'];
    if (m[3] === 'd' || m[3] === 'w') {
      // Calendar days keep the wall-clock time across DST changes
      const w = wallParts(now, tz);
      const days = m[3] === 'd' ? n : n * 7;
      const d = new Date(Date.UTC(w.year, w.month - 1, w.day + days));
      return done(wallTimeToDate({ year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate(), hour: w.hour, minute: w.minute, second: w.second }, tz));
    }
    return done(new Date(now.getTime() + n * ms));
  }

  m = s.match(/^(today|tomorrow|yesterday)(?:\s+(?:at\s+)?(.+))?$/);
  if (m) return onDay(now, { today: 0, tomorrow: 1, yesterday: -1 }[m[1] as 'today' | 'tomorrow' | 'yesterday'], m[2]);

  m = s.match(/^(next\s+)?(sun|mon|tue|wed|thu|fri|sat)[a-z]*(?:\s+(?:at\s+)?(.+))?$/);
  if (m) {
    const target = WEEKDAYS.indexOf(m[2]);
    const w = wallParts(now, tz);
    const today = new Date(Date.UTC(w.year, w.month - 1, w.day)).getUTCDay();
    let delta = (target - today + 7) % 7;
    if (delta === 0 && m[1]) delta = 7; // "next mon" on a Monday means a week out
    return onDay(now, delta, m[3]);
  }

  if (/^\d{1,2}(:\d\d)?\s*(am|pm)?$/.test(s)) return onDay(now, 0, s);

  throw bad();
}
