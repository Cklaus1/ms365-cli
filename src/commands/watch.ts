import { Command, Argument, Option } from 'commander';
import fs from 'fs';
import path from 'path';
import { setTimeout as delay } from 'timers/promises';
import { GraphClient } from '../graph-client.js';
import { configDir, writeFileAtomic } from '../auth.js';
import { CliError, ExitCode, handleErrors } from '../errors.js';
import { info, parseTime, resolveTimeZone, preferTimeZone } from '../utils.js';

const RESOURCES = ['mail', 'events', 'files', 'contacts'] as const;
type Resource = typeof RESOURCES[number];

const MAIL_SELECT = 'id,subject,from,receivedDateTime,createdDateTime,isRead,hasAttachments,bodyPreview';
const MIN_INTERVAL_S = 15;
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

function buildFeed(resource: Resource, opts: Record<string, string | boolean | undefined>): Feed {
  switch (resource) {
    case 'mail': {
      const folder = encodeURIComponent((opts.folder as string | undefined) ?? 'inbox');
      return { resource, endpoint: `/me/mailFolders/${folder}/messages/delta`, initialParams: { $select: MAIL_SELECT }, headers: {} };
    }
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

function cursorName(resource: Resource, opts: Record<string, string | boolean | undefined>): string {
  if (opts.name !== undefined) {
    const name = String(opts.name);
    if (!NAME_PATTERN.test(name)) throw usage(`Invalid --name "${name}": use letters, digits, ".", "_" or "-".`);
    return name;
  }
  if (resource === 'mail' && opts.folder) return `mail-${String(opts.folder).replace(/[^A-Za-z0-9._-]/g, '_')}`;
  return resource;
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
 * Returns the new cursor, or undefined under --dry-run.
 */
async function sync(client: GraphClient, feed: Feed, cursor: Cursor | undefined, includeExisting: boolean): Promise<Cursor | undefined> {
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
      console.log(JSON.stringify({ type, resource: feed.resource, id: item.id, item }));
    }
    if (response.deltaLink) return { resource: feed.resource, endpoint: feed.endpoint, deltaLink: response.deltaLink, syncedAt: startedAt };
    if (!response.nextLink) throw new CliError('Graph returned neither a nextLink nor a deltaLink.', { code: 'bad_response' });
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
    .addOption(new Option('--max-passes <n>').hideHelp())
    .action(handleErrors(async (resource: Resource, opts) => {
      const interval = Number(opts.interval);
      if (!Number.isFinite(interval) || interval < MIN_INTERVAL_S) throw usage(`--interval must be at least ${MIN_INTERVAL_S} seconds.`);
      const maxPasses = opts.maxPasses ? Number(opts.maxPasses) : Infinity;
      const name = cursorName(resource, opts);
      const feed = buildFeed(resource, opts);
      const file = cursorPath(name);
      const client = getClient();

      let cursor = opts.reset ? undefined : readCursor(file, feed);
      if (opts.reset) removeCursor(file);

      /** One pass; false when nothing was saved (dry-run). */
      const pass = async (): Promise<boolean> => {
        let next: Cursor | undefined;
        try {
          next = await sync(client, feed, cursor, !!opts.includeExisting);
        } catch (err) {
          if (!cursor || !isExpiredToken(err)) throw err;
          if (opts.strict) {
            throw new CliError(`The saved delta cursor "${name}" has expired. Run again with --reset.`, { code: 'delta_expired', status: (err as CliError).status, graphCode: (err as CliError).graphCode });
          }
          info(`Warning: the saved delta cursor "${name}" has expired; starting a new baseline.`);
          removeCursor(file);
          next = await sync(client, feed, undefined, !!opts.includeExisting);
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
      const stop = new AbortController();
      const onSigint = () => stop.abort();
      process.on('SIGINT', onSigint);
      try {
        for (let n = 1; await pass() && n < maxPasses && !stop.signal.aborted; n++) {
          await watchDeps.sleep(interval * 1000, stop.signal);
          if (stop.signal.aborted) break;
        }
      } finally {
        process.off('SIGINT', onSigint);
      }
    }));
}
