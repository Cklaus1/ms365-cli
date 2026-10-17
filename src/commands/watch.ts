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

const MAIL_SELECT = 'id,subject,from,receivedDateTime,createdDateTime,isRead,hasAttachments,bodyPreview';
const MIN_INTERVAL_S = 15;
/** Longest wait between --follow passes after failures (R-10). */
const MAX_BACKOFF_S = 15 * 60;
/** Exit code for a pass cut short by Ctrl-C (F22: shell convention, 128 + SIGINT). */
const EXIT_INTERRUPTED = 130;
const NAME_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._-]*$/;

/** Saved sync state: where to resume, and when the pass that produced it started. */
interface Cursor {
  resource: Resource;
  endpoint: string;
  deltaLink: string;
  syncedAt: string;
}

/** One delta feed: the initial endpoint, its first-request params and the headers every request needs. */
interface Feed {
  resource: Resource;
  endpoint: string;
  initialParams: Record<string, string>;
  headers: Record<string, string>;
}

type ChangeType = 'created' | 'updated' | 'deleted';

/** Indirection so tests can drive --follow without real waits. */
export const watchDeps = {
  sleep: (ms: number, signal?: AbortSignal): Promise<void> => delay(ms, undefined, { signal }).catch(() => {}),
};

function usage(message: string): CliError {
  return new CliError(message, { code: 'usage', exitCode: ExitCode.USAGE });
}

/** Thrown out of sync() when Ctrl-C arrives mid-pass; nothing from that pass is saved. */
const interrupted = () => new CliError('Interrupted; the cursor was not advanced.', { code: 'interrupted', exitCode: EXIT_INTERRUPTED as ExitCode });

/** `folder` is the resolved mail folder ID or well-known name. */
function buildFeed(resource: Resource, opts: Record<string, string | boolean | undefined>, folder: string): Feed {
  switch (resource) {
    case 'mail':
      return { resource, endpoint: `/me/mailFolders/${encodeURIComponent(folder)}/messages/delta`, initialParams: { $select: MAIL_SELECT }, headers: {} };
    case 'events': {
      const tz = resolveTimeZone(opts.timezone as string | undefined);
      const start = parseTime((opts.start as string | undefined) ?? 'now', { tz }).date;
      const end = parseTime((opts.end as string | undefined) ?? '+30d', { tz }).date;
      if (end <= start) throw usage('--end must be after --start.');
      return {
        resource,
        endpoint: '/me/calendarView/delta',
        initialParams: { startDateTime: start.toISOString(), endDateTime: end.toISOString() },
        headers: preferTimeZone(tz),
      };
    }
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
 * Run one delta sync to its deltaLink, printing each change as an NDJSON line.
 * Without a cursor this is the initial sync: items are printed only with --include-existing.
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
      const type = cursor ? classify(item, cursor.syncedAt) : 'created';
      // AH-3: with --refs the item's _ref is the line's (one place for a host to read it).
      const { _ref, ...rest } = item;
      console.log(JSON.stringify({ type, resource: feed.resource, id: item.id, ...(_ref !== undefined && { _ref }), item: rest }));
    }
    if (response.deltaLink) return { resource: feed.resource, endpoint: feed.endpoint, deltaLink: response.deltaLink, syncedAt: startedAt };
    if (!response.nextLink) throw new CliError('Graph returned neither a nextLink nor a deltaLink.', { code: 'bad_response' });
    if (signal?.aborted) throw interrupted();
    url = response.nextLink;
    queryParams = undefined;
  }
}

export function registerWatchCommands(program: Command, getClient: () => GraphClient): void {
  program
    .command('watch')
    .description('Print changes since the last run as NDJSON (delta queries; cursor saved per --name)')
    .addArgument(new Argument('<resource>', 'What to watch').choices(RESOURCES))
    .option('--folder <folder>', 'mail: folder id or well-known name (default: inbox)')
    .option('--start <time>', 'events: window start (default: now)')
    .option('--end <time>', 'events: window end (default: +30d)')
    .option('--timezone <tz>', 'events: IANA zone for --start/--end and returned times')
    .option('--name <name>', 'Cursor name (default: the resource, or mail-<folder>)')
    .option('--reset', 'Discard the saved cursor and start a new baseline')
    .option('--include-existing', 'On the first run, print existing items instead of only saving the cursor')
    .option('--strict', 'Exit 1 when the saved cursor has expired instead of starting a new baseline')
    .option('--follow', 'Keep watching, one pass per --interval, until Ctrl-C')
    .option('--interval <seconds>', `Seconds between passes with --follow (min ${MIN_INTERVAL_S})`, '60')
    .option('--max-failures <n>', 'With --follow, exit after this many consecutive failed passes', '10')
    .addOption(new Option('--max-passes <n>').hideHelp())
    .addHelpText('after', `
Delivery is at-least-once: the cursor is saved only after a pass completes, so changes printed by a
pass that fails or is interrupted are printed again by the next run. Consumers should dedupe on id
plus lastModifiedDateTime.
With --follow, a failed pass logs a warning and is retried after interval x 2^n (at most 15 min);
auth and usage errors, an expired cursor under --strict, or --max-failures consecutive failures end
the watch. Ctrl-C during a pass exits 130 without saving it; Ctrl-C between passes exits 0.`)
    .action(handleErrors(async (resource: Resource, opts) => {
      const interval = Number(opts.interval);
      if (!Number.isFinite(interval) || interval < MIN_INTERVAL_S) throw usage(`--interval must be at least ${MIN_INTERVAL_S} seconds.`);
      const maxFailures = Number(opts.maxFailures);
      if (!Number.isInteger(maxFailures) || maxFailures < 1) throw usage('--max-failures must be a positive integer.');
      const maxPasses = opts.maxPasses ? Number(opts.maxPasses) : Infinity;
      if (opts.name !== undefined) cursorName(resource, opts, ''); // reject a bad --name before any request
      const client = getClient();
      const folder = resource === 'mail' ? await resolve(client, 'mail-folder', (opts.folder as string | undefined) ?? 'inbox') : '';
      const name = cursorName(resource, opts, folder);
      const feed = buildFeed(resource, opts, folder);
      const file = cursorPath(name);
      // A preview never deletes or writes the cursor (DR-18); --reset only changes what it requests.
      const dryRun = isDryRun();

      let cursor = opts.reset ? undefined : readCursor(file, feed);
      if (opts.reset && !dryRun) removeCursor(file);

      const stop = new AbortController();

      /** One pass; false when nothing was saved (dry-run). */
      const pass = async (): Promise<boolean> => {
        let next: Cursor | undefined;
        try {
          next = await sync(client, feed, cursor, !!opts.includeExisting, stop.signal);
        } catch (err) {
          if (!cursor || !isExpiredToken(err)) throw err;
          if (opts.strict) {
            throw new CliError(`The saved delta cursor "${name}" has expired. Run again with --reset.`, { code: 'delta_expired', status: (err as CliError).status, graphCode: (err as CliError).graphCode });
          }
          info(`Warning: the saved delta cursor "${name}" has expired; starting a new baseline.`);
          if (!dryRun) removeCursor(file);
          next = await sync(client, feed, undefined, !!opts.includeExisting, stop.signal);
        }
        if (!next) return false;
        saveCursor(file, next);
        cursor = next;
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
