import { Command, Argument, Option } from 'commander';
import fs from 'fs';
import path from 'path';
import { setTimeout as delay } from 'timers/promises';
import { GraphClient } from '../graph-client.js';
import { configDir, writeFileAtomic } from '../auth.js';
import { CliError, ExitCode, handleErrors, toCliError } from '../errors.js';
import { info, isDryRun, parseTime, resolveTimeZone, preferTimeZone } from '../utils.js';
import { resolve } from '../resolve.js';

const RESOURCES = ['mail', 'events', 'files', 'contacts'] as const;
type Resource = typeof RESOURCES[number];

const MAIL_SELECT = 'id,subject,from,receivedDateTime,createdDateTime,lastModifiedDateTime,changeKey,isRead,hasAttachments,bodyPreview';
/** Fields every `watch --select` keeps, so a line always has its dedupe key (AH-4(6)). */
const SELECT_BASE = ['id', 'changeKey', 'lastModifiedDateTime'];
const MIN_INTERVAL_S = 15;
/** Longest wait between --follow passes after failures (R-10). */
const MAX_BACKOFF_S = 15 * 60;
/** Exit code for a pass cut short by Ctrl-C (F22: shell convention, 128 + SIGINT). */
const EXIT_INTERRUPTED = 130;
const NAME_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._-]*$/;
/** Version of the --checkpoint blob (OQ-11). */
const CHECKPOINT_VERSION = 1;

/** An events calendarView window, ISO instants. */
interface Window {
  start: string;
  end: string;
}

/** Sync state: where to resume, when the pass that produced it started, and since when coverage is unbroken. */
interface Cursor {
  resource: Resource;
  endpoint: string;
  deltaLink: string;
  syncedAt: string;
  /** Start of the last baseline; absent in cursor files written before AH-4. */
  coverageStart?: string;
}

/** The decoded --checkpoint blob (AH-4(1)): a cursor plus the scope it covers. */
interface Checkpoint extends Cursor {
  v: number;
  coverageStart: string;
  folder?: string;
  window?: Window;
}

/** One delta feed: the initial endpoint, its first-request params and the headers every request needs. */
interface Feed {
  resource: Resource;
  endpoint: string;
  initialParams: Record<string, string>;
  headers: Record<string, string>;
  window?: Window;
  /** --select fields to keep on each item (SELECT_BASE included). */
  select?: string[];
}

type ChangeType = 'created' | 'updated' | 'deleted' | 'backfill';

/** Indirection so tests can drive --follow without real waits and feed `--checkpoint @-`. */
export const watchDeps = {
  sleep: (ms: number, signal?: AbortSignal): Promise<void> => delay(ms, undefined, { signal }).catch(() => {}),
  readStdin: (): string => fs.readFileSync(0, 'utf8'),
};

function usage(message: string): CliError {
  return new CliError(message, { code: 'usage', exitCode: ExitCode.USAGE });
}

function mismatch(message: string): CliError {
  return new CliError(message, { code: 'checkpoint_mismatch', exitCode: ExitCode.USAGE, fix: 'pass the checkpoint printed by this same watch command, or --checkpoint new' });
}

/** Thrown out of sync() when Ctrl-C arrives mid-pass; nothing from that pass is saved. */
const interrupted = () => new CliError('Interrupted; the cursor was not advanced.', { code: 'interrupted', exitCode: EXIT_INTERRUPTED as ExitCode });

/** The events window from --start/--end (defaults now … +30 d). */
function eventsWindow(opts: Record<string, string | boolean | undefined>): Window {
  const tz = resolveTimeZone(opts.timezone as string | undefined);
  const start = parseTime((opts.start as string | undefined) ?? 'now', { tz }).date;
  const end = parseTime((opts.end as string | undefined) ?? '+30d', { tz }).date;
  if (end <= start) throw usage('--end must be after --start.');
  return { start: start.toISOString(), end: end.toISOString() };
}

/** `folder` is the resolved mail folder ID or well-known name; `window` the events window. */
function buildFeed(resource: Resource, opts: Record<string, string | boolean | undefined>, folder: string, window: Window | undefined): Feed {
  const select = typeof opts.select === 'string' ? [...new Set([...opts.select.split(',').map(s => s.trim()).filter(Boolean), ...SELECT_BASE])] : undefined;
  switch (resource) {
    case 'mail':
      return { resource, endpoint: `/me/mailFolders/${encodeURIComponent(folder)}/messages/delta`, initialParams: { $select: select?.join(',') ?? MAIL_SELECT }, headers: {}, select };
    case 'events':
      // calendarView/delta may ignore $select †; the fields are projected locally instead (AH-4(6)).
      return {
        resource,
        endpoint: '/me/calendarView/delta',
        initialParams: { startDateTime: window!.start, endDateTime: window!.end },
        headers: preferTimeZone(resolveTimeZone(opts.timezone as string | undefined)),
        window,
        select,
      };
    case 'files':
      // token=latest returns a deltaLink for "now" without enumerating the whole drive.
      return { resource, endpoint: '/me/drive/root/delta', initialParams: opts.includeExisting ? {} : { token: 'latest' }, headers: {} };
    case 'contacts':
      return { resource, endpoint: '/me/contacts/delta', initialParams: {}, headers: {} };
  }
}

/**
 * --name, or a default from the resource and the resolved folder (Q-13), so `Inbox`, `inbox` and
 * no --folder share one cursor. Under MCP (MS365_CLI_MCP=1) the default gets an `mcp-` prefix so
 * agent calls never consume the user's own cursors (X-6).
 */
function cursorName(resource: Resource, opts: Record<string, string | boolean | undefined>, folder: string): string {
  if (opts.name !== undefined) {
    const name = String(opts.name);
    if (!NAME_PATTERN.test(name)) throw usage(`Invalid --name "${name}": use letters, digits, ".", "_" or "-".`);
    return name;
  }
  const base = resource === 'mail' && folder !== 'inbox' ? `mail-${folder.replace(/[^A-Za-z0-9._-]/g, '_')}` : resource;
  return process.env.MS365_CLI_MCP === '1' ? `mcp-${base}` : base;
}

const cursorPath = (name: string) => path.join(configDir(), 'watch', `${name}.json`);

function readCursor(file: string, feed: Feed): Cursor | undefined {
  let raw: string;
  try {
    raw = fs.readFileSync(file, 'utf8');
  } catch {
    return undefined;
  }
  let cursor: Cursor;
  try {
    cursor = JSON.parse(raw) as Cursor;
  } catch {
    throw usage(`Cursor file ${file} is corrupt. Run again with --reset.`);
  }
  if (cursor.resource !== feed.resource || cursor.endpoint !== feed.endpoint) {
    throw usage(`Cursor ${file} belongs to ${cursor.endpoint}, not ${feed.endpoint}. Use a different --name or --reset.`);
  }
  return cursor;
}

function saveCursor(file: string, cursor: Cursor): void {
  fs.mkdirSync(path.dirname(file), { recursive: true, mode: 0o700 });
  writeFileAtomic(file, JSON.stringify(cursor, null, 2));
}

function removeCursor(file: string): void {
  try { fs.unlinkSync(file); } catch { /* none saved */ }
}

/** The --checkpoint value's text: the blob itself, `@file`, or `@-` for stdin. */
function checkpointText(value: string): string {
  if (value === '@-') {
    if (process.env.MS365_CLI_MCP === '1') throw usage('--checkpoint @- cannot read stdin under MCP; pass the blob itself.');
    return watchDeps.readStdin().trim();
  }
  if (value.startsWith('@')) {
    try {
      return fs.readFileSync(value.slice(1), 'utf8').trim();
    } catch (err) {
      throw usage(`Cannot read --checkpoint file ${value.slice(1)}: ${err instanceof Error ? err.message : String(err)}`);
    }
  }
  return value.trim();
}

/** Decode a blob; undefined for an empty value or `new` (start a baseline). */
function decodeCheckpoint(text: string): Checkpoint | undefined {
  if (text === '' || text === 'new') return undefined;
  let parsed: unknown;
  try { parsed = JSON.parse(Buffer.from(text, 'base64url').toString('utf8')); } catch { /* reported below */ }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) throw usage('--checkpoint is not a checkpoint printed by ms365 watch.');
  const cp = parsed as Record<string, unknown>;
  if (cp.v !== CHECKPOINT_VERSION) throw usage(`--checkpoint version ${String(cp.v)} is not supported (expected ${CHECKPOINT_VERSION}).`);
  for (const key of ['resource', 'endpoint', 'deltaLink', 'syncedAt', 'coverageStart']) {
    if (typeof cp[key] !== 'string') throw usage(`--checkpoint lacks ${key}.`);
  }
  return parsed as Checkpoint;
}

const encodeCheckpoint = (cp: Checkpoint) => Buffer.from(JSON.stringify(cp)).toString('base64url');

/**
 * A Graph path reduced to what a deltaLink keeps of the endpoint: version, owner (`me`/`users`) and the
 * delta kind (last two segments). Graph may rewrite folder ids as `mailFolders('…')` †, so the folder is
 * compared through the blob's `folder` field instead (E10).
 */
function deltaShape(pathname: string): string {
  let p = pathname;
  try { p = decodeURIComponent(p); } catch { /* keep as written */ }
  const segments = p.toLowerCase().replace(/\('?([^')]*)'?\)/g, '/$1').split('/').filter(Boolean);
  return [segments[0], segments[1], ...segments.slice(-2)].join('/');
}

/** The scope checks of AH-4(1): before any request, a blob must belong to this command (OQ-19). */
function checkCheckpoint(cp: Checkpoint, client: GraphClient, resource: Resource, opts: Record<string, string | boolean | undefined>): void {
  if (cp.resource !== resource) throw mismatch(`--checkpoint belongs to watch ${cp.resource}, not watch ${resource}.`);
  let link: URL | undefined;
  try { link = new URL(cp.deltaLink); } catch { /* reported below */ }
  const expected = new URL(client.buildUrl(cp.endpoint));
  if (link?.hostname !== 'graph.microsoft.com' || deltaShape(link.pathname) !== deltaShape(expected.pathname)) {
    throw mismatch(`--checkpoint resumes ${link?.pathname ?? 'an invalid URL'}, not ${expected.pathname}.`);
  }
  if (resource === 'events' && cp.window && (opts.start !== undefined || opts.end !== undefined)) {
    const asked = eventsWindow({ ...opts, start: opts.start ?? cp.window.start, end: opts.end ?? cp.window.end });
    if (asked.start !== cp.window.start || asked.end !== cp.window.end) {
      throw mismatch(`--checkpoint covers ${cp.window.start}…${cp.window.end}, not ${asked.start}…${asked.end}.`);
    }
  }
}

/** Errors that end --follow at once instead of being retried next pass (R-10). */
function isFatal(err: unknown): boolean {
  const e = toCliError(err);
  return e.exitCode === ExitCode.AUTH || e.exitCode === ExitCode.USAGE || e.code === 'delta_expired' || e.code === 'interrupted';
}

/** Graph's "this delta token can no longer be used" responses (Outlook and OneDrive variants). */
function isExpiredToken(err: unknown): boolean {
  if (!(err instanceof CliError)) return false;
  return err.status === 410 || /syncState|resync/i.test(err.graphCode ?? '');
}

/** Outlook marks removals with "@removed"; OneDrive with a "deleted" facet. */
const isRemoved = (item: Record<string, unknown>) => '@removed' in item || 'deleted' in item;

function classify(item: Record<string, unknown>, since: string): ChangeType {
  if (isRemoved(item)) return 'deleted';
  const created = typeof item.createdDateTime === 'string' ? Date.parse(item.createdDateTime) : NaN;
  return created >= Date.parse(since) ? 'created' : 'updated';
}

/**
 * One NDJSON change line (AH-4(2)): `version` and `lastModified` are the dedupe key; mail is `untrusted`;
 * a delete says why it left (`@removed.reason`: `deleted`, or `changed` = moved out of scope †).
 */
function changeLine(feed: Feed, item: Record<string, unknown>, type: ChangeType): Record<string, unknown> {
  const { _ref, ...rest } = item;
  const removed = item['@removed'];
  const reason = removed !== null && typeof removed === 'object' && 'reason' in removed && typeof removed.reason === 'string' ? removed.reason : 'deleted';
  const removedReason = type === 'deleted' ? reason : undefined;
  const kept = feed.select ? Object.fromEntries(Object.entries(rest).filter(([k]) => feed.select!.includes(k))) : rest;
  return {
    type,
    resource: feed.resource,
    id: item.id,
    version: item.changeKey ?? item['@odata.etag'] ?? item.eTag ?? null,
    lastModified: item.lastModifiedDateTime ?? null,
    ...(feed.resource === 'mail' && { untrusted: true }),
    ...(removedReason !== undefined && { removedReason }),
    ...(_ref !== undefined && { _ref }),
    item: kept,
  };
}

/**
 * Run one delta sync to its deltaLink, printing each change as an NDJSON line.
 * Without a cursor this is a baseline: items are printed (as `backfill`) only with `includeExisting`.
 * Returns the new cursor, or undefined under --dry-run. An aborted `signal` stops it after the
 * current page with an `interrupted` error.
 */
async function sync(client: GraphClient, feed: Feed, cursor: Cursor | undefined, includeExisting: boolean, signal?: AbortSignal): Promise<Cursor | undefined> {
  const startedAt = new Date().toISOString();
  let url = cursor?.deltaLink ?? feed.endpoint;
  let queryParams: Record<string, string> | undefined = cursor ? undefined : feed.initialParams;
  for (;;) {
    const response = await client.request(url, { queryParams, headers: feed.headers });
    if (response.status === 0) return undefined;
    const items = ((response.data as { value?: Record<string, unknown>[] }).value) ?? [];
    for (const item of items) {
      if (!cursor && (!includeExisting || isRemoved(item))) continue;
      console.log(JSON.stringify(changeLine(feed, item, cursor ? classify(item, cursor.syncedAt) : 'backfill')));
    }
    if (response.deltaLink) {
      return { resource: feed.resource, endpoint: feed.endpoint, deltaLink: response.deltaLink, syncedAt: startedAt, coverageStart: cursor?.coverageStart ?? startedAt };
    }
    if (!response.nextLink) throw new CliError('Graph returned neither a nextLink nor a deltaLink.', { code: 'bad_response' });
    if (signal?.aborted) throw interrupted();
    url = response.nextLink;
    queryParams = undefined;
  }
}

export function registerWatchCommands(program: Command, getClient: () => GraphClient): void {
  program
    .command('watch')
    .description('Print changes since the last run as NDJSON (delta queries; cursor saved per --name, or held by the caller with --checkpoint)')
    .addArgument(new Argument('<resource>', 'What to watch').choices(RESOURCES))
    .option('--folder <folder>', 'mail: folder id or well-known name (default: inbox)')
    .option('--start <time>', 'events: window start (default: now)')
    .option('--end <time>', 'events: window end (default: +30d)')
    .option('--timezone <tz>', 'events: IANA zone for --start/--end and returned times')
    .option('--name <name>', 'Cursor name (default: the resource, or mail-<folder>)')
    .option('--reset', 'Discard the saved cursor and start a new baseline')
    .option('--checkpoint <blob|@file|@->', 'Resume from a checkpoint the caller holds (the blob, a file, or stdin); "new" or "" starts a baseline. No cursor file is read or written')
    .option('--select <fields>', 'mail, events: keep only these item fields (id, changeKey and lastModifiedDateTime are always kept)')
    .option('--include-existing', 'On a baseline, print existing items (type "backfill") instead of only saving the cursor')
    .option('--strict', 'Exit 1 when the cursor has expired instead of emitting a gap and a backfill')
    .option('--follow', 'Keep watching, one pass per --interval, until Ctrl-C')
    .option('--interval <seconds>', `Seconds between passes with --follow (min ${MIN_INTERVAL_S})`, '60')
    .option('--max-failures <n>', 'With --follow, exit after this many consecutive failed passes', '10')
    .addOption(new Option('--max-passes <n>').hideHelp())
    .addHelpText('after', `
Output is always NDJSON (no -o). Line types: created, updated, deleted (a tombstone with removedReason),
backfill (an existing item from a baseline or a re-sync), gap and, with --checkpoint, one checkpoint line
at the end of every pass. Each change line has id, version and lastModified; mail lines have untrusted: true.
created/updated is a heuristic: dedupe on id + version, never on type.

Delivery is at-least-once: the cursor is saved (or the checkpoint printed) only after a pass completes,
so changes printed by a pass that fails or is interrupted are printed again by the next run. With
--checkpoint, persist the records and the checkpoint line together, then pass that checkpoint next time.
An expired cursor prints {"type":"gap","reason":"sync_state_expired",...}, then backfill lines, then the
new state (--strict exits instead). With --checkpoint, an events window with less than half left rolls
forward: a gap with reason window_rolled, backfill lines for the new window, then the checkpoint.
With --follow, a failed pass logs a warning and is retried after interval x 2^n (at most 15 min);
auth and usage errors, an expired cursor under --strict, or --max-failures consecutive failures end
the watch. Ctrl-C during a pass exits 130 without saving it; Ctrl-C between passes exits 0.`)
    .action(handleErrors(async (resource: Resource, opts) => {
      const interval = Number(opts.interval);
      if (!Number.isFinite(interval) || interval < MIN_INTERVAL_S) throw usage(`--interval must be at least ${MIN_INTERVAL_S} seconds.`);
      const maxFailures = Number(opts.maxFailures);
      if (!Number.isInteger(maxFailures) || maxFailures < 1) throw usage('--max-failures must be a positive integer.');
      const maxPasses = opts.maxPasses ? Number(opts.maxPasses) : Infinity;
      const held = opts.checkpoint !== undefined;
      if (held && (opts.name !== undefined || opts.reset)) throw usage('--checkpoint cannot be combined with --name or --reset.');
      if (opts.select !== undefined && resource !== 'mail' && resource !== 'events') throw usage('--select applies to watch mail and watch events.');
      if (opts.name !== undefined) cursorName(resource, opts, ''); // reject a bad --name before any request
      const client = getClient();
      const checkpoint = held ? decodeCheckpoint(checkpointText(String(opts.checkpoint))) : undefined;
      if (checkpoint) checkCheckpoint(checkpoint, client, resource, opts);
      const folder = resource !== 'mail' ? ''
        : opts.folder === undefined && checkpoint?.folder !== undefined ? checkpoint.folder
          : await resolve(client, 'mail-folder', (opts.folder as string | undefined) ?? 'inbox');
      if (checkpoint && resource === 'mail' && checkpoint.folder !== folder) throw mismatch(`--checkpoint follows folder ${String(checkpoint.folder)}, not ${folder}.`);
      let window = resource === 'events' ? checkpoint?.window ?? eventsWindow(opts) : undefined;
      let feed = buildFeed(resource, opts, folder, window);
      if (checkpoint && checkpoint.endpoint !== feed.endpoint) throw mismatch(`--checkpoint belongs to ${checkpoint.endpoint}, not ${feed.endpoint}.`);
      const name = held ? '' : cursorName(resource, opts, folder);
      const file = held ? '' : cursorPath(name);
      // A preview never deletes or writes the cursor (DR-18); --reset only changes what it requests.
      const dryRun = isDryRun();

      let cursor: Cursor | undefined = held ? checkpoint : opts.reset ? undefined : readCursor(file, feed);
      if (opts.reset && !dryRun) removeCursor(file);

      const stop = new AbortController();
      const gap = (reason: string, from: string, to: string, extra: Record<string, unknown> = {}) =>
        console.log(JSON.stringify({ type: 'gap', resource, reason, from, to, ...extra }));

      /** One pass; false when nothing was saved (dry-run). */
      const pass = async (): Promise<boolean> => {
        let next: Cursor | undefined;
        // AH-4(5): roll an events window that has less than half left, so coverage never silently ends.
        if (held && cursor && window && !dryRun) {
          const start = Date.parse(window.start), end = Date.parse(window.end), now = Date.now();
          if (end - now < (end - start) / 2) {
            const rolled = { start: new Date(now).toISOString(), end: new Date(now + end - start).toISOString() };
            gap('window_rolled', window.end, rolled.end, { window: rolled });
            window = rolled;
            feed = buildFeed(resource, opts, folder, window);
            cursor = undefined;
            next = await sync(client, feed, undefined, true, stop.signal);
          }
        }
        if (!next) {
          try {
            next = await sync(client, feed, cursor, !!opts.includeExisting, stop.signal);
          } catch (err) {
            if (!cursor || !isExpiredToken(err)) throw err;
            const what = held ? 'checkpoint' : `saved delta cursor "${name}"`;
            if (opts.strict) {
              throw new CliError(`The ${what} has expired. Run again with ${held ? '--checkpoint new' : '--reset'}.`, { code: 'delta_expired', status: (err as CliError).status, graphCode: (err as CliError).graphCode });
            }
            // AH-4(3): never re-baseline silently — say what was lost, then re-send everything that exists.
            info(`Warning: the ${what} has expired; emitting a gap and re-syncing existing items as backfill.`);
            gap('sync_state_expired', cursor.syncedAt, new Date().toISOString());
            if (!dryRun && !held) removeCursor(file);
            next = await sync(client, feed, undefined, true, stop.signal);
          }
        }
        if (!next) return false;
        cursor = next;
        if (!held) {
          saveCursor(file, next);
          return true;
        }
        const blob: Checkpoint = {
          v: CHECKPOINT_VERSION, ...next, coverageStart: next.coverageStart ?? next.syncedAt,
          ...(resource === 'mail' && { folder }), ...(window && { window }),
        };
        console.log(JSON.stringify({
          type: 'checkpoint', resource, checkpoint: encodeCheckpoint(blob), coverageStart: blob.coverageStart, coverageEnd: next.syncedAt,
          syncedAt: next.syncedAt, ...(blob.folder !== undefined && { folder }), ...(window && { window }),
        }));
        return true;
      };

      if (!opts.follow) {
        await pass();
        return;
      }
      // First Ctrl-C: finish the current page (or sleep) and stop; second: exit at once.
      const onSigint = () => {
        if (stop.signal.aborted) process.exit(EXIT_INTERRUPTED);
        stop.abort();
      };
      process.on('SIGINT', onSigint);
      try {
        let failures = 0;
        for (let n = 1; ; n++) {
          let waitS = interval;
          try {
            if (!await pass()) break;
            failures = 0;
          } catch (err) {
            if (isFatal(err) || ++failures >= maxFailures) throw err;
            waitS = Math.min(interval * 2 ** failures, MAX_BACKOFF_S);
            info(`Warning: watch pass failed (${failures}/${maxFailures}): ${toCliError(err).message}; retrying in ${waitS}s.`);
          }
          if (n >= maxPasses || stop.signal.aborted) break;
          await watchDeps.sleep(waitS * 1000, stop.signal);
          if (stop.signal.aborted) break;
        }
      } finally {
        process.off('SIGINT', onSigint);
      }
    }));
}
