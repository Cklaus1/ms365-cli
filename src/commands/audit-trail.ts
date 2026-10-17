import { Command, InvalidArgumentError } from 'commander';
import { GraphClient } from '../graph-client.js';
import { printList } from '../formatter.js';
import { outputOption, addListOptions, info, odataQuote, parseTime, resolveTimeZone, truncate, validateEmail } from '../utils.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';

/**
 * Admin oversight of access itself (AO-9, AO-11..AO-16). Reads Microsoft's audit log and directory and prints
 * *access events and grants*, never the content that was accessed.
 *
 * Privacy contract (AO-9, AC 7): a record's `auditData` is reduced by `projectRecord()` / `accessFacts()` to an
 * allow-list of access metadata before anything is printed. `Item`, `AffectedItems` (subjects),
 * `Folders[].FolderItems` (InternetMessageId), recipients and raw `auditData` never leave this module; item counts
 * are counted, never copied.
 *
 * The query engine (`createAuditQuery` → `waitForAuditQuery` → `printAuditRecords`) is shared by every
 * audit-search command; the projection is pure so the offline CSV import can reuse it.
 */

// --- query engine: POST /security/auditLog/queries → poll → paged records (AO-13) ---

export const QUERIES = '/security/auditLog/queries';
/** Exchange mailbox-access record types (AO-9). */
export const MAILBOX_RECORD_TYPES = ['exchangeItem', 'exchangeItemGroup', 'exchangeItemAggregated'] as const;
/** Operations a mailbox-access query asks for unless --operations overrides them. */
export const DEFAULT_OPERATIONS = ['MailItemsAccessed', 'FolderBind', 'MessageBind', 'SendAs', 'SendOnBehalf'] as const;
const POLL_START_MS = 2_000;
const POLL_CAP_MS = 15_000;
export const DEFAULT_TIMEOUT_S = 300;
/** Audit (Standard) retention; Audit (Premium) keeps a year (†, AO-12). */
export const STANDARD_RETENTION_DAYS = 180;

/** The poll clock. Tests replace it; production waits for real. */
export const timing = {
  now: (): number => Date.now(),
  sleep: (ms: number): Promise<void> => new Promise(resolve => setTimeout(resolve, ms)),
};

/** POST body of a security auditLogQuery (Graph v1.0). `serviceFilter(s)` is never sent (†, record types suffice). */
export interface AuditQueryBody {
  displayName: string;
  filterStartDateTime: string;
  filterEndDateTime: string;
  recordTypeFilters?: string[];
  operationFilters?: string[];
  keywordFilter?: string;
  userPrincipalNameFilters?: string[];
}

export type AuditQueryStatus = 'notStarted' | 'running' | 'succeeded' | 'failed' | 'cancelled';

export interface AuditQuery {
  id: string;
  status: AuditQueryStatus | string;
  isRecordCountLimitExceeded?: boolean;
  recordCountLimit?: number;
  approximateReturnedRecordCount?: number;
}

/**
 * Create an audit query. Returns null under --dry-run (the POST was traced; there is nothing to poll).
 * `--read-only` blocks the POST (exit 4) with a message that says why a "read" command is a write.
 */
export async function createAuditQuery(client: GraphClient, body: AuditQueryBody): Promise<AuditQuery | null> {
  let res;
  try {
    res = await client.request(QUERIES, { method: 'POST', body: JSON.stringify(body) });
  } catch (err) {
    if (err instanceof CliError && err.code === 'read_only') {
      throw new CliError('Read-only mode: this command creates an audit search query (a POST; no mailbox or directory data is changed), which --read-only blocks.',
        { code: 'read_only', exitCode: ExitCode.PERMISSION, fix: 'run it without --read-only' });
    }
    throw err;
  }
  if (client.isPreview(res)) return null;
  const query = res.data as AuditQuery;
  if (!query?.id) throw new CliError('Graph did not return an audit query id.', { code: 'bad_response' });
  return query;
}

/**
 * Poll until the query leaves notStarted/running: backoff 2 s doubling to a 15 s cap, up to `timeoutMs`.
 * `succeeded` returns the final query; `failed`/`cancelled` and a timeout exit 1.
 */
export async function waitForAuditQuery(client: GraphClient, query: AuditQuery, timeoutMs = DEFAULT_TIMEOUT_S * 1000): Promise<AuditQuery> {
  const deadline = timing.now() + timeoutMs;
  let backoff = POLL_START_MS;
  let current = query;
  while (current.status === 'notStarted' || current.status === 'running') {
    const left = deadline - timing.now();
    if (left <= 0) {
      throw new CliError(`Audit query ${query.id} did not finish within ${Math.round(timeoutMs / 1000)} s (status: ${current.status}).`, {
        code: 'timeout', exitCode: ExitCode.GENERIC,
        fix: 'narrow the window with --since/--until or --operations, or raise --timeout',
      });
    }
    await timing.sleep(Math.min(backoff, left));
    backoff = Math.min(backoff * 2, POLL_CAP_MS);
    current = (await client.request(`${QUERIES}/${encodeURIComponent(query.id)}`)).data as AuditQuery;
  }
  if (current.status !== 'succeeded') {
    throw new CliError(`Audit query ${query.id} ${current.status}.`, {
      code: `audit_query_${current.status}`, exitCode: ExitCode.GENERIC,
      fix: 'run the command again; if it keeps failing, narrow --since/--until',
    });
  }
  if (current.isRecordCountLimitExceeded === true) { // authoritative; never inferred from the counts (AO-13)
    const n = current.recordCountLimit ?? current.approximateReturnedRecordCount;
    info(`Warning: the audit query was truncated at ${n ?? 'the record limit of'} records; narrow --since/--until or --operations.`);
  }
  return current;
}

/** The records collection of a finished query (paged by @odata.nextLink). */
export function auditRecordsPath(queryId: string): string {
  return `${QUERIES}/${encodeURIComponent(queryId)}/records`;
}

// --- allow-list projection (AO-9) ---

type Obj = Record<string, unknown>;
const obj = (v: unknown): Obj => (v && typeof v === 'object' && !Array.isArray(v) ? v as Obj : {});
const str = (v: unknown): string | undefined => (typeof v === 'string' && v !== '' ? v : undefined);
const numOr = (v: unknown, d: number | undefined): number | undefined => (typeof v === 'number' && Number.isFinite(v) ? v : d);

/** auditData.LogonType (Exchange). */
const LOGON_TYPES: Record<number, string> = {
  0: 'owner', 1: 'admin', 2: 'delegated', 3: 'transport', 4: 'systemService', 5: 'bestAccess', 6: 'delegatedAdmin',
};
/** auditData.UserType (unified audit log), used when no Graph envelope is present (CSV import). */
const USER_TYPES: Record<number, string> = {
  0: 'regular', 1: 'reserved', 2: 'admin', 3: 'dcAdmin', 4: 'system', 5: 'application', 6: 'servicePrincipal',
  7: 'customPolicy', 8: 'systemPolicy', 9: 'partnerTechnician', 10: 'guest',
};

/** The printed columns of a mailbox-access row, in order. Nothing else from a record is ever printed. */
export const MAILBOX_FIELDS = [
  'createdDateTime', 'operation', 'actor', 'userType', 'logonType', 'clientIp', 'client', 'appId',
  'operationCount', 'externalAccess', 'folderCount',
] as const;

export interface MailboxAccessRow {
  createdDateTime: string | null;
  operation: string | null;
  actor: string | null;
  userType: string | null;
  logonType: string | null;
  clientIp: string | null;
  client: string | null;
  appId: string | null;
  operationCount: number;
  externalAccess: boolean | null;
  folderCount: number;
}

/** The Graph auditLogRecord properties outside auditData that the projection may read. */
export interface AuditEnvelope {
  createdDateTime?: unknown;
  operation?: unknown;
  userPrincipalName?: unknown;
  userType?: unknown;
  clientIp?: unknown;
}

/**
 * Reduce one audit record to the printed allow-list (pure). `auditData` is the record's auditData (or a Purview
 * CSV `AuditData` cell, parsed); `envelope` is the Graph record when there is one — its fields win, auditData's
 * CreationTime/Operation/UserId/UserType/ClientIP* fill in otherwise.
 */
export function projectRecord(auditData: unknown, envelope: AuditEnvelope = {}): MailboxAccessRow {
  const a = obj(auditData);
  const logon = numOr(a.LogonType, undefined);
  const userTypeCode = numOr(a.UserType, undefined);
  const client = str(a.ClientInfoString);
  const folders = Array.isArray(a.Folders) ? a.Folders.length : 0;
  const external = typeof a.ExternalAccess === 'boolean' ? a.ExternalAccess : null;
  return {
    createdDateTime: str(envelope.createdDateTime) ?? str(a.CreationTime) ?? null,
    operation: str(envelope.operation) ?? str(a.Operation) ?? null,
    actor: str(envelope.userPrincipalName) ?? str(a.UserId) ?? null,
    userType: str(envelope.userType) ?? (userTypeCode !== undefined ? USER_TYPES[userTypeCode] ?? String(userTypeCode) : null),
    logonType: logon !== undefined ? LOGON_TYPES[logon] ?? String(logon) : null,
    clientIp: str(envelope.clientIp) ?? str(a.ClientIPAddress) ?? str(a.ClientIP) ?? null,
    client: client ? truncate(client, 120) : null,
    appId: appIdOf(a) ?? null,
    operationCount: numOr(a.OperationCount, 1)!,
    externalAccess: external,
    folderCount: folders,
  };
}

function appIdOf(a: Obj): string | undefined {
  return str(a.AppId) ?? str(a.ClientAppId) ?? str(obj(a.AppAccessContext).ClientAppId);
}

/** projectRecord() plus the facts aggregation needs (target mailbox, raw logon code, item count). Pure. */
export interface AccessFacts extends MailboxAccessRow {
  /** auditData.MailboxOwnerUPN, lower-cased; null when absent. */
  mailboxOwner: string | null;
  /** auditData.LogonType as a number; null when absent. */
  logonTypeCode: number | null;
  /** Items touched: Σ len(Folders[].FolderItems), else len(AffectedItems), else 1 for a single Item. Counted, never copied. */
  itemCount: number;
}

export function accessFacts(auditData: unknown, envelope: AuditEnvelope = {}): AccessFacts {
  const a = obj(auditData);
  const folders = Array.isArray(a.Folders) ? a.Folders : [];
  let itemCount = folders.reduce((n: number, f) => n + (Array.isArray(obj(f).FolderItems) ? (obj(f).FolderItems as unknown[]).length : 0), 0);
  if (itemCount === 0 && Array.isArray(a.AffectedItems)) itemCount = a.AffectedItems.length;
  if (itemCount === 0 && a.Item && typeof a.Item === 'object') itemCount = 1;
  return {
    ...projectRecord(a, envelope),
    mailboxOwner: str(a.MailboxOwnerUPN)?.toLowerCase() ?? null,
    logonTypeCode: numOr(a.LogonType, undefined) ?? null,
    itemCount,
  };
}

/** True when the record's mailbox owner is `upn` (case-insensitive; the authoritative mailbox filter, AO-9). */
export function isMailboxOf(auditData: unknown, upn: string): boolean {
  const owner = str(obj(auditData).MailboxOwnerUPN);
  return !!owner && owner.toLowerCase() === upn.toLowerCase();
}

// --- coverage (AO-12) ---

export interface Window { since: Date; until: Date }

/** --since/--until (default: the last `days` days up to now); since must precede until (exit 2). */
export function parseWindow(opts: { since?: string; until?: string }, days = 7): Window {
  const tz = resolveTimeZone();
  const until = opts.until ? parseTime(opts.until, { tz }).date : new Date();
  const since = opts.since ? parseTime(opts.since, { tz }).date : new Date(until.getTime() - days * 86_400_000);
  if (since.getTime() >= until.getTime()) {
    throw new CliError('--since must be before --until.', { code: 'usage', exitCode: ExitCode.USAGE });
  }
  return { since, until };
}

/**
 * The stderr coverage line every audit search prints (AO-12): window, operations, retention floor, and the fixed
 * caveat that an empty result is not proof of no access. `notes` are appended (e.g. own-mailbox events omitted).
 */
export function printCoverage(window: Window, operations: readonly string[], notes: string[] = []): void {
  const day = (d: Date) => d.toISOString().replace(/\.\d{3}Z$/, 'Z');
  info(`Coverage: ${day(window.since)} → ${day(window.until)}; operations: ${operations.join(', ')}; `
    + `retention: Audit (Standard) ${STANDARD_RETENTION_DAYS} days, Audit (Premium) 1 year. `
    + 'Message reads (MailItemsAccessed) appear only where mailbox auditing records them; an empty result is not proof of no access.'
    + (notes.length ? ' ' + notes.join(' ') : ''));
  const ageDays = (Date.now() - window.since.getTime()) / 86_400_000;
  if (ageDays > STANDARD_RETENTION_DAYS) {
    info(`Warning: --since is more than ${STANDARD_RETENTION_DAYS} days ago, beyond Audit (Standard) retention; older events exist only for Audit (Premium) mailboxes (1 year).`);
  }
}

// --- 403 hints (AO-16) ---

const PURVIEW_DENIED = /do(?:n'?t|esn'?t|nt| not) have any permissions/i;
const CONSENT_MISSING = /scope|consent|AuditLogsQuery|AuditLog\.Read\.All/i;
const GRANT = 'admin consent via `ms365 --org-mode auth grant-scopes`';

export const HINTS = {
  purview: 'the signed-in person needs the Purview Audit Reader role group (View-Only Audit Logs: Purview → Settings → Roles and scopes → Role groups); role changes can take an hour or more to apply',
  queryConsent: `the ms365 app needs the AuditLogsQuery-Exchange.Read.All permission: ${GRANT}`,
  query: 'two prerequisites: (a) the Purview Audit Reader role group for the signed-in person, and (b) AuditLogsQuery-Exchange.Read.All consent for the ms365 app (`ms365 --org-mode auth grant-scopes`)',
  entraRole: 'the signed-in person needs an Entra role that reads directory audits: Reports Reader, Security Reader or Security Administrator',
  directoryConsent: `the ms365 app needs the AuditLog.Read.All permission: ${GRANT}`,
} as const;

/** Rethrow a 403 with the hint for this command's prerequisite (exit 4). */
function withHint(err: unknown, kind: 'query' | 'directory'): never {
  if (err instanceof CliError && err.code === 'permission_denied') {
    const m = err.message;
    const fix = kind === 'query'
      ? PURVIEW_DENIED.test(m) ? HINTS.purview : CONSENT_MISSING.test(m) ? HINTS.queryConsent : HINTS.query
      : CONSENT_MISSING.test(m) ? HINTS.directoryConsent : HINTS.entraRole;
    throw new CliError(err.message.replace(/ Try running with --org-mode\.$/, ''),
      { code: err.code, exitCode: ExitCode.PERMISSION, status: err.status, graphCode: err.graphCode, requestId: err.requestId, fix });
  }
  throw err;
}

/** Wrap a command action so a 403 carries the audit-specific hint. */
export function auditHints(kind: 'query' | 'directory', fn: (...args: any[]) => Promise<void>): (...args: any[]) => Promise<void> { // eslint-disable-line @typescript-eslint/no-explicit-any
  return async (...args) => {
    try { await fn(...args); } catch (err) { withHint(err, kind); }
  };
}

// --- shared option parsers ---

function timeoutSeconds(v: string): number {
  const n = Number(v);
  if (!/^\d+$/.test(v.trim()) || n < 1 || n > 3600) throw new InvalidArgumentError('must be an integer from 1 to 3600 (seconds).');
  return n;
}

function operationList(v: string): string[] {
  const ops = v.split(',').map(s => s.trim()).filter(Boolean);
  if (ops.length === 0 || ops.some(o => !/^[A-Za-z][A-Za-z0-9-]*$/.test(o))) {
    throw new InvalidArgumentError('a comma-separated list of operation names, e.g. MailItemsAccessed,SendAs.');
  }
  return ops;
}

/**
 * The shared list flags for a query's records. A query's records are always read in full (a later --cursor run
 * would create a new query), so --all is implied; --page-limit and --page-delay apply.
 */
function pageOptions(cmd: Command): Command {
  addListOptions(cmd);
  const all = cmd.options.find(o => o.long === '--all');
  if (all) all.description = 'Implied: every record page of the query is read';
  return cmd;
}

// --- directory audits (AO-11) ---

/** Role and consent changes kept client-side (an `or` list of display names in $filter is unverified †). */
export const CHANGE_ACTIVITIES = new Set([
  'Add member to role', 'Add eligible member to role', 'Consent to application', 'Add app role assignment to service principal',
  'Add app role assignment grant to user', 'Add delegated permission grant', 'Add OAuth2PermissionGrant',
]);

export interface ChangeRow {
  activityDateTime: string | null;
  activityDisplayName: string | null;
  result: string | null;
  initiatedBy: string | null;
  targets: string;
}

/** Reduce one directoryAudit to the printed allow-list (pure): no modifiedProperties, ids or additionalDetails. */
export function projectChange(audit: unknown): ChangeRow {
  const a = obj(audit);
  const by = obj(a.initiatedBy);
  const user = obj(by.user);
  const app = obj(by.app);
  const targets = (Array.isArray(a.targetResources) ? a.targetResources : []).map((t) => {
    const r = obj(t);
    const name = str(r.userPrincipalName) ?? str(r.displayName) ?? '?';
    return str(r.type) ? `${name} (${r.type})` : name;
  });
  return {
    activityDateTime: str(a.activityDateTime) ?? null,
    activityDisplayName: str(a.activityDisplayName) ?? null,
    result: str(a.result) ?? null,
    initiatedBy: str(user.userPrincipalName) ?? str(user.displayName) ?? str(app.displayName) ?? null,
    targets: targets.join('; '),
  };
}

// --- commands ---

export function registerAuditTrailCommands(program: Command, getClient: () => GraphClient): void {
  const trail = program.command('audit-trail')
    .description('Admin oversight of access: who accessed which mailbox and how, and role/consent changes (access events only, never content)');

  pageOptions(trail
    .command('mailbox')
    .description('Who accessed a mailbox and how (owner, delegate, admin, app); creates a saved audit query (no data is changed). Prints access metadata only, never subjects or recipients.')
    .argument('<upn>', 'Mailbox owner UPN')
    .option('--since <time>', 'Window start (ISO 8601, "2026-10-01", "-30d"; default 7 days ago)')
    .option('--until <time>', 'Window end (default now)')
    .option('--operations <list>', `Comma-separated operations (default ${DEFAULT_OPERATIONS.join(',')})`, operationList)
    .option('--by <upn>', 'Only events by this actor (exact UPN)')
    .option('--timeout <s>', `Seconds to wait for the query (default ${DEFAULT_TIMEOUT_S})`, timeoutSeconds, DEFAULT_TIMEOUT_S))
    .addOption(outputOption())
    .action(handleErrors(auditHints('query', async (upnArg: string, opts) => {
      const upn = validateEmail(upnArg);
      const by = opts.by !== undefined ? validateEmail(opts.by) : undefined;
      const window = parseWindow(opts);
      const operations: string[] = opts.operations ?? [...DEFAULT_OPERATIONS];
      const client = getClient();

      printCoverage(window, operations);
      const query = await createAuditQuery(client, {
        displayName: `ms365 audit-trail mailbox ${new Date().toISOString()}`,
        filterStartDateTime: window.since.toISOString(),
        filterEndDateTime: window.until.toISOString(),
        recordTypeFilters: [...MAILBOX_RECORD_TYPES],
        operationFilters: operations,
        keywordFilter: upn, // server-side narrowing (†); the owner filter below is authoritative
        ...(by ? { userPrincipalNameFilters: [by] } : {}),
      });
      if (!query) return; // --dry-run: the POST was traced, nothing to poll
      const done = await waitForAuditQuery(client, query, opts.timeout * 1000);

      await printList(client, auditRecordsPath(done.id), {}, { all: true, pageLimit: opts.pageLimit, pageDelay: opts.pageDelay, output: opts.output }, {
        transform: async (items) => items
          .map(i => obj(i))
          .filter(r => isMailboxOf(r.auditData, upn))
          .map(r => projectRecord(r.auditData, r)),
      });
    })));

  addListOptions(trail
    .command('changes')
    .description('Admin role and app consent changes from the directory audit log (who changed what, when)')
    .option('--since <time>', 'Window start (ISO 8601, "2026-10-01", "-30d"; default 7 days ago)')
    .option('--until <time>', 'Window end (default: no upper bound)'), { description: 'Audit entries per page (before the activity filter)' })
    .addOption(outputOption())
    .action(handleErrors(auditHints('directory', async (opts) => {
      const tz = resolveTimeZone();
      const since = opts.since ? parseTime(opts.since, { tz }).date : new Date(Date.now() - 7 * 86_400_000);
      const until = opts.until ? parseTime(opts.until, { tz }).date : undefined;
      if (until && since.getTime() >= until.getTime()) {
        throw new CliError('--since must be before --until.', { code: 'usage', exitCode: ExitCode.USAGE });
      }
      const time = [`activityDateTime ge ${odataQuote(since.toISOString())}`, ...(until ? [`activityDateTime le ${odataQuote(until.toISOString())}`] : [])];
      const filter = `${time.join(' and ')} and (category eq 'RoleManagement' or category eq 'ApplicationManagement')`;
      const queryParams: Record<string, string> = { $filter: filter, $orderby: 'activityDateTime desc' };
      if (opts.top) queryParams.$top = opts.top;
      await printList(getClient(), '/auditLogs/directoryAudits', { queryParams }, opts, {
        transform: async (items) => items
          .filter(i => CHANGE_ACTIVITIES.has(String(obj(i).activityDisplayName)))
          .map(projectChange),
      });
    })));
}
