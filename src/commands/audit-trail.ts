import fs from 'fs';
import { Command, InvalidArgumentError } from 'commander';
import { GraphClient } from '../graph-client.js';
import { printList, formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption, addListOptions, info, odataQuote, parseTime, resolveTimeZone, truncate, validateEmail } from '../utils.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import { tokenInfo } from '../auth.js';
import { csvRecords } from '../csv.js';

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

/** The calling app. In Exchange records `AppId` is often the resource (Microsoft Graph), so the client comes first. */
function appIdOf(a: Obj): string | undefined {
  return str(a.ClientAppId) ?? str(obj(a.AppAccessContext).ClientAppId) ?? str(a.AppId);
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
  appsRole: 'the signed-in person needs a directory-reading Entra role: Directory Readers, Global Reader, Application Administrator or Cloud Application Administrator',
  appsConsent: `the ms365 app needs the Directory.Read.All and Application.Read.All permissions: ${GRANT}`,
} as const;

const APPS_CONSENT_MISSING = /scope|consent|Directory\.Read\.All|Application\.Read\.All/i;

type HintKind = 'query' | 'directory' | 'apps';

/** Rethrow a 403 with the hint for this command's prerequisite (exit 4). */
function withHint(err: unknown, kind: HintKind): never {
  if (err instanceof CliError && err.code === 'permission_denied') {
    const m = err.message;
    const fix = kind === 'query'
      ? PURVIEW_DENIED.test(m) ? HINTS.purview : CONSENT_MISSING.test(m) ? HINTS.queryConsent : HINTS.query
      : kind === 'apps'
        ? APPS_CONSENT_MISSING.test(m) ? HINTS.appsConsent : HINTS.appsRole
        : CONSENT_MISSING.test(m) ? HINTS.directoryConsent : HINTS.entraRole;
    throw new CliError(err.message.replace(/ Try running with --org-mode\.$/, ''),
      { code: err.code, exitCode: ExitCode.PERMISSION, status: err.status, graphCode: err.graphCode, requestId: err.requestId, fix });
  }
  throw err;
}

/** Wrap a command action so a 403 carries the audit-specific hint. */
export function auditHints(kind: HintKind, fn: (...args: any[]) => Promise<void>): (...args: any[]) => Promise<void> { // eslint-disable-line @typescript-eslint/no-explicit-any
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

// --- target annotation (AO-30) ---

/** The fixed gloss for a flagged target: Graph does not expose the recipient type (†), so the flag must not accuse. */
export const DISABLED_OR_SHARED_GLOSS = 'disabled, unlicensed account: a shared mailbox or a retired account';

/**
 * Per-run cache of target-mailbox annotations, keyed by lower-cased UPN: one
 * `GET /users/{upn}?$select=accountEnabled,assignedLicenses` per distinct mailbox (User.Read.All).
 * `disabled-or-shared` = accountEnabled false and no licence; `not-found` = 404; null = nothing to flag.
 * A 403 leaves every target unflagged with one stderr warning and is never fatal.
 */
export class TargetAnnotator {
  private cache = new Map<string, string | null>();
  private denied = false;
  constructor(private client: GraphClient) {}

  async target(upn: string): Promise<string | null> {
    const key = upn.toLowerCase();
    if (this.cache.has(key)) return this.cache.get(key)!;
    let flag: string | null = null;
    if (!this.denied) {
      try {
        const res = await this.client.request(`/users/${encodeURIComponent(key)}`, { queryParams: { $select: 'accountEnabled,assignedLicenses' } });
        const u = obj(res.data);
        const licences = Array.isArray(u.assignedLicenses) ? u.assignedLicenses.length : 0;
        if (u.accountEnabled === false && licences === 0) flag = 'disabled-or-shared';
      } catch (err) {
        if (!(err instanceof CliError)) throw err;
        if (err.status === 404 || err.code === 'not_found') flag = 'not-found';
        else if (err.status === 403 || err.code === 'permission_denied') {
          this.denied = true;
          info('Warning: target annotation skipped (403 reading the mailbox accounts; it needs User.Read.All). Rows are unflagged.');
        } else throw err;
      }
    }
    this.cache.set(key, flag);
    return flag;
  }
}

// --- audit-trail apps (AO-10, AO-28, AO-33) ---

export const GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000';
export const EXCHANGE_APP_ID = '00000002-0000-0ff1-ce00-000000000000';
/** Application permissions (app roles of Microsoft Graph) that read mail, chats, calendars or files tenant-wide. Resolved by value at runtime. */
export const BROAD_APP_ROLES = ['Mail.Read', 'Mail.ReadBasic.All', 'Mail.ReadWrite', 'Chat.Read.All', 'ChannelMessage.Read.All', 'Calendars.Read', 'Files.Read.All'] as const;
/** Delegated Graph scopes that, consented for all users, reach other people's mail/chats/files (exact token match). */
export const BROAD_DELEGATED_SCOPES = ['Mail.Read', 'Mail.Read.Shared', 'Mail.ReadWrite', 'Mail.ReadWrite.Shared', 'Chat.Read', 'ChannelMessage.Read.All', 'Files.Read.All', 'Sites.Read.All'] as const;
/** Delegated Exchange Online scopes that are full mailbox access through EWS/IMAP/POP (AO-33; values †). */
export const BROAD_EXCHANGE_SCOPES = ['EWS.AccessAsUser.All', 'IMAP.AccessAsUser.All', 'POP.AccessAsUser.All'] as const;

export const APP_FIELDS = [
  'displayName', 'appId', 'servicePrincipalId', 'publisher', 'origin', 'kind', 'scopes',
  'appRoleAssignmentRequired', 'disabledByMicrosoftStatus', 'assignedPrincipals', 'granted',
] as const;

export interface AppRow {
  displayName: string | null;
  appId: string | null;
  servicePrincipalId: string;
  /** verifiedPublisher.displayName, or "unverified". */
  publisher: string;
  /** "this tenant" | "external" (appOwnerOrganizationId vs the token's tid) | "unknown". */
  origin: string;
  kind: 'application' | 'delegated-all-users' | 'delegated-all-users (Exchange Online)';
  scopes: string;
  appRoleAssignmentRequired: boolean | null;
  disabledByMicrosoftStatus: string | null;
  assignedPrincipals: number | null;
  /** Earliest app role assignment (application kind); grants carry no date on v1.0. */
  granted: string | null;
}

const SP_SELECT = 'id,appId,displayName,verifiedPublisher,appOwnerOrganizationId,appRoleAssignmentRequired,signInAudience,servicePrincipalType,disabledByMicrosoftStatus';

/**
 * The signed-in tenant: the token's `tid` claim (no /organization call). A seam so tests can set it; undefined
 * when the token is not a readable JWT (origin then prints "unknown").
 */
export const tenant = {
  id: async (client: GraphClient): Promise<string | undefined> => {
    try {
      // GraphClient keeps its AuthManager private; the token is only decoded here, never printed or sent.
      const auth = (client as unknown as { auth: { getToken(): Promise<string> } }).auth;
      return tokenInfo(await auth.getToken()).tenantId;
    } catch { return undefined; }
  },
};

/** Exact, case-sensitive token match of a space-separated grant scope against a list. */
function scopeTokens(scope: unknown, wanted: readonly string[]): string[] {
  const have = new Set((typeof scope === 'string' ? scope : '').split(/\s+/).filter(Boolean));
  return wanted.filter(w => have.has(w));
}

/** Every item of a paged collection. */
async function allItems(client: GraphClient, path: string, queryParams?: Record<string, string>): Promise<Obj[]> {
  const data = obj((await client.requestAllPages(path, queryParams ? { queryParams } : {})).data);
  return (Array.isArray(data.value) ? data.value : []).map(obj);
}

async function getOrNull(client: GraphClient, path: string, queryParams?: Record<string, string>): Promise<Obj | null> {
  try {
    return obj((await client.request(path, queryParams ? { queryParams } : {})).data);
  } catch (err) {
    if (err instanceof CliError && (err.status === 404 || err.code === 'not_found')) return null;
    throw err;
  }
}

/** AllPrincipals delegated grants: server-side `consentType eq` filter, falling back to the full list on a 400 (†). */
async function allPrincipalsGrants(client: GraphClient): Promise<Obj[]> {
  let items: Obj[];
  try {
    items = await allItems(client, '/oauth2PermissionGrants', { $filter: "consentType eq 'AllPrincipals'" });
  } catch (err) {
    if (!(err instanceof CliError) || err.status !== 400) throw err;
    info('Note: Graph rejected the consentType filter on oauth2PermissionGrants; reading every grant and filtering locally.');
    items = await allItems(client, '/oauth2PermissionGrants');
  }
  return items.filter(g => g.consentType === 'AllPrincipals' && !str(g.principalId));
}

/**
 * Apps with broad access (pure facts, no score — E3): application permissions on Microsoft Graph and tenant-wide
 * delegated grants on Graph and Exchange Online. Returns null under --dry-run (the first GET was traced).
 */
export async function listBroadApps(client: GraphClient): Promise<AppRow[] | null> {
  const graphRes = await client.request(`/servicePrincipals(appId='${odataQuote(GRAPH_APP_ID)}')`, { queryParams: { $select: 'id,appRoles' } });
  if (client.isPreview(graphRes)) return null;
  const graph = obj(graphRes.data);
  const graphSpId = str(graph.id);
  if (!graphSpId) throw new CliError('The Microsoft Graph service principal was not found in this tenant.', { code: 'bad_response' });

  const roleValue = new Map<string, string>();
  for (const r of Array.isArray(graph.appRoles) ? graph.appRoles : []) {
    const role = obj(r);
    if (str(role.id) && (BROAD_APP_ROLES as readonly string[]).includes(String(role.value))) roleValue.set(String(role.id), String(role.value));
  }

  const exoSpId = str((await getOrNull(client, `/servicePrincipals(appId='${odataQuote(EXCHANGE_APP_ID)}')`, { $select: 'id' }))?.id);

  interface Acc { spId: string; kind: AppRow['kind']; scopes: Set<string>; granted: string | null; fallbackName: string | null }
  const acc = new Map<string, Acc>();
  const add = (spId: string, kind: AppRow['kind'], scopes: string[], granted: string | null, fallbackName: string | null) => {
    const key = `${spId}|${kind}`;
    const a = acc.get(key) ?? { spId, kind, scopes: new Set<string>(), granted: null, fallbackName };
    scopes.forEach(s => a.scopes.add(s));
    if (granted && (!a.granted || granted < a.granted)) a.granted = granted;
    acc.set(key, a);
  };

  for (const a of await allItems(client, `/servicePrincipals/${encodeURIComponent(graphSpId)}/appRoleAssignedTo`)) {
    const value = roleValue.get(String(a.appRoleId));
    const spId = str(a.principalId);
    if (value && spId) add(spId, 'application', [value], str(a.createdDateTime) ?? null, str(a.principalDisplayName) ?? null);
  }

  for (const g of await allPrincipalsGrants(client)) {
    const spId = str(g.clientId);
    if (!spId) continue;
    if (g.resourceId === graphSpId) {
      const hit = scopeTokens(g.scope, BROAD_DELEGATED_SCOPES);
      if (hit.length) add(spId, 'delegated-all-users', hit, null, null);
    } else if (exoSpId && g.resourceId === exoSpId) {
      const hit = scopeTokens(g.scope, BROAD_EXCHANGE_SCOPES);
      if (hit.length) add(spId, 'delegated-all-users (Exchange Online)', hit, null, null);
    }
  }

  const home = await tenant.id(client);
  const sps = new Map<string, { sp: Obj | null; assignedPrincipals: number | null }>();
  const rows: AppRow[] = [];
  for (const a of acc.values()) {
    let facts = sps.get(a.spId);
    if (!facts) {
      const sp = await getOrNull(client, `/servicePrincipals/${encodeURIComponent(a.spId)}`, { $select: SP_SELECT });
      const assignedPrincipals = sp
        ? (await allItems(client, `/servicePrincipals/${encodeURIComponent(a.spId)}/appRoleAssignedTo`, { $select: 'id', $top: '999' })).length
        : null;
      facts = { sp, assignedPrincipals };
      sps.set(a.spId, facts);
    }
    const sp = facts.sp ?? {};
    const owner = str(sp.appOwnerOrganizationId);
    rows.push({
      displayName: str(sp.displayName) ?? a.fallbackName,
      appId: str(sp.appId) ?? null,
      servicePrincipalId: a.spId,
      publisher: str(obj(sp.verifiedPublisher).displayName) ?? 'unverified',
      origin: !owner || !home ? 'unknown' : owner.toLowerCase() === home.toLowerCase() ? 'this tenant' : 'external',
      kind: a.kind,
      scopes: [...a.scopes].join(' '),
      appRoleAssignmentRequired: typeof sp.appRoleAssignmentRequired === 'boolean' ? sp.appRoleAssignmentRequired : null,
      disabledByMicrosoftStatus: str(sp.disabledByMicrosoftStatus) ?? null,
      assignedPrincipals: facts.assignedPrincipals,
      granted: a.granted,
    });
  }
  return rows.sort((x, y) => String(x.displayName ?? '').localeCompare(String(y.displayName ?? '')) || x.kind.localeCompare(y.kind));
}

// --- audit-trail access (AO-29) ---

export const ACCESS_FIELDS = ['actor', 'mailbox', 'logonType', 'events', 'items', 'first', 'last', 'client', 'appId', 'target'] as const;

export interface AccessRow {
  actor: string | null;
  mailbox: string | null;
  logonType: string | null;
  /** Audit records in the group. */
  events: number;
  /** Items touched (counted from Folders[].FolderItems / AffectedItems / Item, never copied). */
  items: number;
  first: string | null;
  last: string | null;
  /** Distinct ClientInfoString values (each ≤ 120 chars), at most three. */
  client: string | null;
  /** Distinct app IDs, at most three. */
  appId: string | null;
  /** AO-30 annotation: disabled-or-shared | not-found | null. */
  target: string | null;
}

export interface AccessFilter {
  app?: string;
  by?: string;
  /** Only access to this mailbox (owner UPN, case-insensitive); used by `import --mailbox`. */
  mailbox?: string;
}

/**
 * Aggregates access records per (actor, mailbox, logonType). Own-mailbox events (LogonType 0, or the mailbox owner
 * is the actor) are dropped before aggregation and only counted (AO-29 guard c). Feed it page by page; nothing
 * but the aggregate is kept.
 */
export class AccessAggregator {
  ownOmitted = 0;
  private groups = new Map<string, { row: AccessRow; clients: Set<string>; apps: Set<string> }>();
  constructor(private filter: AccessFilter) {}

  add(auditData: unknown, envelope: AuditEnvelope = {}): void {
    const f = accessFacts(auditData, envelope);
    if (this.filter.app && f.appId?.toLowerCase() !== this.filter.app.toLowerCase()) return;
    if (this.filter.by && f.actor?.toLowerCase() !== this.filter.by.toLowerCase()) return;
    if (this.filter.mailbox && f.mailboxOwner !== this.filter.mailbox.toLowerCase()) return;
    const actor = f.actor?.toLowerCase() ?? null;
    if (f.logonTypeCode === 0 || (actor !== null && f.mailboxOwner === actor)) { this.ownOmitted++; return; }
    const key = JSON.stringify([actor, f.mailboxOwner, f.logonType]);
    let g = this.groups.get(key);
    if (!g) {
      g = { row: { actor, mailbox: f.mailboxOwner, logonType: f.logonType, events: 0, items: 0, first: null, last: null, client: null, appId: null, target: null },
        clients: new Set(), apps: new Set() };
      this.groups.set(key, g);
    }
    g.row.events++;
    g.row.items += f.itemCount;
    const t = f.createdDateTime;
    if (t && (!g.row.first || t < g.row.first)) g.row.first = t;
    if (t && (!g.row.last || t > g.row.last)) g.row.last = t;
    if (f.client) g.clients.add(f.client);
    if (f.appId) g.apps.add(f.appId);
  }

  rows(): AccessRow[] {
    const list = (s: Set<string>) => s.size === 0 ? null : [...s].slice(0, 3).join('; ') + (s.size > 3 ? ` (+${s.size - 3} more)` : '');
    return [...this.groups.values()]
      .map(g => ({ ...g.row, client: list(g.clients), appId: list(g.apps) }))
      .sort((a, b) => String(a.actor).localeCompare(String(b.actor)) || String(a.mailbox).localeCompare(String(b.mailbox))
        || String(a.logonType).localeCompare(String(b.logonType)));
  }
}

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Exactly one value per flag (commander passes the previous value on a repeat). */
function single<T>(parse: (v: string) => T, flag: string) {
  return (v: string, prev: T | undefined): T => {
    if (prev !== undefined) throw new InvalidArgumentError(`give ${flag} once.`);
    return parse(v);
  };
}

function appIdArg(v: string): string {
  const t = v.trim();
  if (!GUID.test(t)) throw new InvalidArgumentError('an application (client) ID GUID, e.g. 00000000-0000-0000-0000-000000000000.');
  return t.toLowerCase();
}

/** One exact UPN: userPrincipalNameFilters is exact-match, so lists and wildcards are refused. */
function singleUpn(v: string): string {
  const t = v.trim();
  if (/[\s,;*?%]/.test(t) || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(t)) throw new InvalidArgumentError('exactly one UPN (no lists or wildcards).');
  return t;
}

// --- audit-trail import (AO-31) ---

/** Purview export caps (Learn): an export with exactly this many rows was probably cut off. */
export const EXPORT_CAPS: ReadonlyArray<{ rows: number; tier: string }> = [
  { rows: 50_000, tier: 'Audit (Standard)' },
  { rows: 1_000_000, tier: 'Audit (Premium)' },
];
/** Read size for the export file; the CSV parser only ever holds one chunk plus one partial record. */
export const IMPORT_CHUNK_BYTES = 64 * 1024;

/** `import` prints the access aggregate without the target annotation (offline: no directory lookup). */
export const IMPORT_FIELDS = ACCESS_FIELDS.filter(f => f !== 'target');

/** Header names (trimmed, lower-cased) across the export variants; unknown columns are ignored. */
const IMPORT_COLUMNS = {
  auditData: ['auditdata'],
  recordId: ['recordid'],
  creationDate: ['creationdate'],
  userId: ['userid', 'userids'],
  operation: ['operation', 'operations'],
} as const;
type ColumnIndex = Record<keyof typeof IMPORT_COLUMNS, number>; // -1 when the column is absent

function columnIndex(header: string[]): ColumnIndex {
  const names = header.map(h => h.trim().toLowerCase());
  const out = {} as ColumnIndex;
  for (const [key, aliases] of Object.entries(IMPORT_COLUMNS) as [keyof ColumnIndex, readonly string[]][]) {
    out[key] = names.findIndex(n => aliases.includes(n));
  }
  return out;
}

const ISO_TIME = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?$/i;

/** An ISO 8601 time as UTC ISO (a zone-less time is UTC: Purview writes CreationTime without one); else undefined. */
function isoTime(v: unknown): string | undefined {
  if (typeof v !== 'string') return undefined;
  const t = v.trim();
  const m = ISO_TIME.exec(t);
  if (!m) return undefined;
  const d = new Date(m[1] ? t : `${t}Z`);
  return Number.isNaN(d.getTime()) ? undefined : d.toISOString();
}

const badCsv = (message: string) => new CliError(message, { code: 'bad_csv', exitCode: ExitCode.GENERIC });

export interface ImportResult {
  /** Data rows in the file (every record after the header). */
  rows: number;
  /** Rows whose operation is not a mailbox-access operation (DEFAULT_OPERATIONS). */
  otherOperations: number;
  first: string | null;
  last: string | null;
  agg: AccessAggregator;
}

/**
 * Stream a Purview audit export through the access aggregation (AO-31). Columns are found by header name; only
 * `AuditData` is required, and `CreationDate`, `UserId|UserIds`, `Operation|Operations` win over the JSON's
 * CreationTime/UserId/Operation when present. Each AuditData is parsed, aggregated and dropped. Errors name the
 * row number and a GUID-shaped RecordId only, never field text.
 */
export async function importAuditCsv(source: AsyncIterable<Uint8Array | string>, filter: AccessFilter): Promise<ImportResult> {
  const agg = new AccessAggregator(filter);
  const operations = new Set<string>(DEFAULT_OPERATIONS);
  let cols: ColumnIndex | undefined;
  let rows = 0;
  let otherOperations = 0;
  let first: string | null = null;
  let last: string | null = null;
  for await (const rec of csvRecords(source)) {
    if (!cols) {
      cols = columnIndex(rec);
      if (cols.auditData < 0) {
        throw new CliError('The file has no AuditData column; is it a Purview audit-log export?', {
          code: 'bad_csv', exitCode: ExitCode.USAGE,
          fix: 'export from Purview → Audit → Export, or Search-UnifiedAuditLog | Export-Csv',
        });
      }
      continue;
    }
    rows++;
    const c = cols;
    const cell = (i: number): string => (i >= 0 ? rec[i] ?? '' : '').trim();
    const where = (): string => {
      const id = cell(c.recordId);
      return `row ${rows}${GUID.test(id) ? ` (RecordId ${id})` : ''}`;
    };
    const raw = cell(c.auditData);
    if (!raw) throw badCsv(`${where()}: AuditData is empty.`);
    let parsed: unknown;
    try { parsed = JSON.parse(raw); } catch { throw badCsv(`${where()}: AuditData is not JSON.`); } // never the parser's message: it quotes the text
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) throw badCsv(`${where()}: AuditData is not a JSON object.`);
    const a = parsed as Obj;

    const createdDateTime = isoTime(cell(c.creationDate)) ?? isoTime(a.CreationTime);
    if (createdDateTime) {
      if (!first || createdDateTime < first) first = createdDateTime;
      if (!last || createdDateTime > last) last = createdDateTime;
    }
    const envelope: AuditEnvelope = { createdDateTime, operation: str(cell(c.operation)), userPrincipalName: str(cell(c.userId)) };
    const operation = str(envelope.operation) ?? str(a.Operation);
    if (!operation || !operations.has(operation)) { otherOperations++; continue; }
    agg.add(a, envelope);
  }
  if (!cols) throw new CliError('The file is empty.', { code: 'bad_csv', exitCode: ExitCode.USAGE });
  return { rows, otherOperations, first, last, agg };
}

// --- commands ---

export function registerAuditTrailCommands(program: Command, getClient: () => GraphClient): void {
  const trail = program.command('audit-trail')
    .description('Admin oversight of access: who accessed which mailbox and how, which apps hold broad access, and role/consent changes (access events and grants only, never content)');

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
      const target = await new TargetAnnotator(client).target(upn); // AO-30: annotation only
      if (target === 'disabled-or-shared') info(`Target mailbox ${upn}: disabled-or-shared (${DISABLED_OR_SHARED_GLOSS}).`);
      else if (target === 'not-found') info(`Target mailbox ${upn}: not-found (no directory account with this UPN).`);
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

  trail
    .command('apps')
    .description('Apps with broad access to mail, chats or files: Microsoft Graph application permissions and tenant-wide (all users) delegated grants on Graph and Exchange Online. Facts only, no risk score.')
    .addOption(outputOption())
    .action(handleErrors(auditHints('apps', async (opts) => {
      const rows = await listBroadApps(getClient());
      if (!rows) return; // --dry-run: the first GET was traced
      console.log(formatOutput({ value: rows }, opts.output as OutputFormat));
    })));

  trail
    .command('access')
    .description('Which mailboxes an app (--app) or a person (--by) opened: aggregated per actor, mailbox and logon type; own-mailbox events are omitted. Creates a saved audit query (no data is changed).')
    .option('--app <appId>', 'Application (client) ID of the app (GUID)', single(appIdArg, '--app'))
    .option('--by <upn>', 'The acting person (one exact UPN)', single(singleUpn, '--by'))
    .option('--since <time>', 'Window start (ISO 8601, "2026-10-01", "-30d"; default 30 days ago)')
    .option('--until <time>', 'Window end (default now)')
    .option('--timeout <s>', `Seconds to wait for the query (default ${DEFAULT_TIMEOUT_S})`, timeoutSeconds, DEFAULT_TIMEOUT_S)
    .option('--page-limit <n>', 'Max record pages to read', (v: string) => {
      if (!/^\d+$/.test(v.trim()) || Number(v) < 1) throw new InvalidArgumentError('a positive integer.');
      return Number(v);
    })
    .addOption(outputOption())
    .action(handleErrors(auditHints('query', async (opts) => {
      const app: string | undefined = opts.app;
      const by: string | undefined = opts.by;
      if (!app && !by) {
        throw new CliError('audit-trail access needs --app <appId> or --by <upn>.', {
          code: 'usage', exitCode: ExitCode.USAGE, fix: 'for every access to one mailbox use `ms365 audit-trail mailbox <upn>`',
        });
      }
      const window = parseWindow(opts, 30);
      const operations = [...DEFAULT_OPERATIONS];
      const spanDays = (window.until.getTime() - window.since.getTime()) / 86_400_000;
      const client = getClient();

      const query = await createAuditQuery(client, {
        displayName: `ms365 audit-trail access ${new Date().toISOString()}`,
        filterStartDateTime: window.since.toISOString(),
        filterEndDateTime: window.until.toISOString(),
        recordTypeFilters: [...MAILBOX_RECORD_TYPES],
        operationFilters: operations,
        ...(app ? { keywordFilter: app } : {}), // server-side narrowing (†); the appId filter below is authoritative
        ...(by ? { userPrincipalNameFilters: [by] } : {}),
      });
      const spanNote = spanDays > STANDARD_RETENTION_DAYS ? [`The window spans ${Math.round(spanDays)} days, beyond the ${STANDARD_RETENTION_DAYS}-day Audit (Standard) retention.`] : [];
      if (!query) { printCoverage(window, operations, spanNote); return; } // --dry-run: the POST was traced, nothing to poll
      const done = await waitForAuditQuery(client, query, opts.timeout * 1000);

      const agg = new AccessAggregator({ app, by });
      for await (const page of client.requestPagesStreaming(auditRecordsPath(done.id), {}, opts.pageLimit)) {
        for (const r of page.map(obj)) agg.add(r.auditData, r);
      }
      printCoverage(window, operations, [...spanNote, `${agg.ownOmitted} own-mailbox events omitted.`]);

      const rows = agg.rows();
      const annotate = new TargetAnnotator(client);
      for (const row of rows) if (row.mailbox) row.target = await annotate.target(row.mailbox);
      if (rows.some(r => r.target === 'disabled-or-shared')) info(`target disabled-or-shared: ${DISABLED_OR_SHARED_GLOSS}.`);
      console.log(formatOutput({ value: rows }, opts.output as OutputFormat));
    })));

  trail
    .command('import')
    .description('Offline analysis of a Purview audit-log export (CSV): which mailboxes an app or person opened, who opened a mailbox, or all non-owner mailbox access; aggregated per actor, mailbox and logon type, own-mailbox events omitted. No sign-in, no network.')
    .argument('<file.csv>', 'Purview audit export (portal export or Search-UnifiedAuditLog | Export-Csv)')
    .option('--app <appId>', 'Only access by this application (client) ID (GUID)', single(appIdArg, '--app'))
    .option('--by <upn>', 'Only access by this person (one exact UPN)', single(singleUpn, '--by'))
    .option('--mailbox <upn>', 'Only access to this mailbox (owner UPN), by anyone but its owner', single(singleUpn, '--mailbox'))
    .option('--cross-mailbox', 'All non-owner mailbox access in the file')
    .addOption(outputOption())
    .action(handleErrors(async (file: string, opts) => {
      const filter: AccessFilter = { app: opts.app, by: opts.by, mailbox: opts.mailbox };
      if (!filter.app && !filter.by && !filter.mailbox && !opts.crossMailbox) {
        throw new CliError('audit-trail import needs a view: --app <appId>, --by <upn>, --mailbox <upn> or --cross-mailbox.', { code: 'usage', exitCode: ExitCode.USAGE });
      }
      let stat: fs.Stats;
      try { stat = fs.statSync(file); } catch {
        throw new CliError(`Cannot read ${file}: no such file.`, { code: 'not_found', exitCode: ExitCode.NOT_FOUND });
      }
      if (!stat.isFile()) throw new CliError(`${file} is not a file.`, { code: 'usage', exitCode: ExitCode.USAGE });

      const res = await importAuditCsv(fs.createReadStream(file, { highWaterMark: IMPORT_CHUNK_BYTES }), filter);
      const cap = EXPORT_CAPS.find(x => x.rows === res.rows);
      if (cap) {
        info(`Warning: export probably truncated: ${cap.rows.toLocaleString('en-US')} rows is the ${cap.tier} export cap, and Purview cuts off silently. Export shorter windows and import each file.`);
      }
      info(`Coverage (file): ${res.rows} rows, ${res.first ?? '?'} → ${res.last ?? '?'}; operations: ${DEFAULT_OPERATIONS.join(', ')} `
        + `(${res.otherOperations} rows of other operations ignored); ${res.agg.ownOmitted} own-mailbox events omitted. `
        + 'Message reads (MailItemsAccessed) appear only where mailbox auditing records them; an empty result is not proof of no access.');
      const rows = res.agg.rows().map(r => Object.fromEntries(IMPORT_FIELDS.map(f => [f, r[f]])));
      console.log(formatOutput({ value: rows }, opts.output as OutputFormat));
    }));
}
