import { Command, InvalidArgumentError, Option } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption, info, odataQuote } from '../utils.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import { parseCsvObjects } from '../csv.js';

/**
 * Aggregate activity reports (Microsoft 365 usage reports, Reports.Read.All).
 *
 * These are counts Microsoft already computes — emails sent/received, Teams messages, meetings.
 * They never contain message content, and this module never prints a per-person row: tenant-level
 * counts come straight from Graph, and department roll-ups pool any group smaller than the
 * minimum group size (default 5) so small teams cannot be singled out.
 */

const PERIODS = ['D7', 'D30', 'D90', 'D180'] as const;
export const MIN_GROUP_SIZE = 5;

/** Lower-case, strip spaces/underscores: "Send Count" and "sendCount" both become "sendcount". */
const norm = (k: string) => k.toLowerCase().replace(/[\s_]/g, '');

type Row = Record<string, unknown>;

function normalize(row: Row): Row {
  return Object.fromEntries(Object.entries(row).map(([k, v]) => [norm(k), v]));
}

function num(v: unknown): number {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

/** Tenant-level count reports: one row per day, no user data. */
const COUNT_REPORTS: Record<string, { fn: string; description: string }> = {
  email: { fn: 'getEmailActivityCounts', description: 'Daily email activity: sent, received, read, meetings created/interacted (tenant totals)' },
  teams: { fn: 'getTeamsUserActivityCounts', description: 'Daily Teams activity: team and private chat messages, calls, meetings (tenant totals)' },
  'active-users': { fn: 'getOffice365ActiveUserCounts', description: 'Daily active users per Microsoft 365 service (tenant totals)' },
};

/** Per-user detail reports, used only for the department roll-up and never printed. */
const DETAIL = {
  email: { fn: 'getEmailActivityUserDetail', fields: ['sendcount', 'receivecount', 'readcount', 'meetingcreatedcount', 'meetinginteractedcount'] },
  teams: { fn: 'getTeamsUserActivityUserDetail', fields: ['teamchatmessagecount', 'privatechatmessagecount', 'callcount', 'meetingcount', 'meetingsorganizedcount', 'meetingsattendedcount'] },
} as const;

export interface RollUpRow { department: string; people: number; [metric: string]: string | number }

/**
 * Sum metrics by department without ever describing fewer than `floor` people, even across runs.
 *
 * - Departments of `min` or more people get their own row.
 * - Departments between `floor` and `min` are pooled into "Other (groups under N people)".
 * - Departments under the `floor` join that pool only when all of them together reach the floor;
 *   otherwise they are dropped at **every** minimum (AO-22). They therefore only ever appear as one
 *   block of at least `floor` people, so subtracting rows from runs with different `--min-group-size`
 *   values cannot isolate a group smaller than the floor.
 * - The pool is printed only if it reaches `min`. Per-person values are never returned.
 */
export function rollUp(
  details: Row[],
  departmentOf: Map<string, string>,
  fields: readonly string[],
  min = MIN_GROUP_SIZE,
  floor = MIN_GROUP_SIZE,
): { rows: RollUpRow[]; unmatched: number } {
  const zero = () => Object.fromEntries(fields.map(f => [f, 0])) as Record<string, number>;
  const groups = new Map<string, { people: number; sums: Record<string, number> }>();
  let unmatched = 0;
  for (const raw of details) {
    const r = normalize(raw);
    const upn = String(r.userprincipalname ?? '').toLowerCase();
    if (!upn || /^[0-9a-f]{16,}$/i.test(upn)) { unmatched++; continue; } // concealed names can't be joined
    const dept = departmentOf.get(upn);
    if (!dept) { unmatched++; continue; }
    const g = groups.get(dept) ?? { people: 0, sums: zero() };
    g.people++;
    for (const f of fields) g.sums[f] += num(r[f]);
    groups.set(dept, g);
  }
  const add = (into: { people: number; sums: Record<string, number> }, g: { people: number; sums: Record<string, number> }) => {
    into.people += g.people;
    for (const f of fields) into.sums[f] += g.sums[f];
  };
  // The sub-floor block is decided at the floor, independent of `min`.
  const smallPool = { people: 0, sums: zero() };
  for (const g of groups.values()) if (g.people < floor) add(smallPool, g);

  const rows: RollUpRow[] = [];
  const pool = { people: 0, sums: zero() };
  if (smallPool.people >= floor) add(pool, smallPool);
  for (const [department, g] of groups) {
    if (g.people < floor) continue;
    if (g.people >= min) rows.push({ department, people: g.people, ...g.sums });
    else add(pool, g);
  }
  if (pool.people >= min) rows.push({ department: `Other (groups under ${min} people)`, people: pool.people, ...pool.sums });
  rows.sort((a, b) => b.people - a.people);
  return { rows, unmatched };
}

/** CSV cells arrive as text; turn plain numbers back into numbers so CSV and JSON output agree. */
function csvValue(v: string): string | number {
  return /^-?\d+(\.\d+)?$/.test(v) ? Number(v) : v;
}

/**
 * Read a usage report whatever format Graph answers in (AO-4). v1.0 rejects `$format=application/json`
 * ("JSON format is not supported", verified live 2026-10-17), so no format is requested: Graph answers with the
 * 302 → CSV download from reports.office.com (followed without the bearer token); a JSON answer is still
 * accepted and paged (AO-23). Rows keep their source keys; callers normalize.
 */
async function readReport(client: GraphClient, fn: string, period: string): Promise<Row[]> {
  const res = await client.requestAllPages(`/reports/${fn}(period='${odataQuote(period)}')`);
  const value = ((res.data as { value?: unknown[] }).value ?? []);
  // A non-JSON body comes back as one item: text types as { rawResponse }, octet-stream as base64 { contentBytes }.
  if (value.length === 1) {
    const body = value[0] as { rawResponse?: unknown; contentBytes?: unknown };
    const text = typeof body.rawResponse === 'string' ? body.rawResponse
      : typeof body.contentBytes === 'string' ? Buffer.from(body.contentBytes, 'base64').toString('utf8')
        : undefined;
    if (text !== undefined) {
      return parseCsvObjects(text).map(o => Object.fromEntries(Object.entries(o).map(([k, v]) => [k, csvValue(v)])));
    }
  }
  return value as Row[];
}

const REPORTS_403_FIX = 'the signed-in person needs Reports Reader (or an Exchange/Teams admin role) for per-user detail; '
  + 'Global Reader and Usage Summary Reports Reader can only run `reports email|teams|active-users`; '
  + 'admin consent via `ms365 --org-mode auth grant-scopes`';

/** Rethrow a 403 with the reports-specific roles hint (AO-27). */
function withReportsHint(fn: (...args: any[]) => Promise<void>): (...args: any[]) => Promise<void> { // eslint-disable-line @typescript-eslint/no-explicit-any
  return async (...args) => {
    try {
      await fn(...args);
    } catch (err) {
      if (err instanceof CliError && err.code === 'permission_denied') {
        throw new CliError(err.message, { code: err.code, exitCode: err.exitCode, status: err.status, graphCode: err.graphCode, requestId: err.requestId, fix: REPORTS_403_FIX });
      }
      throw err;
    }
  };
}

export function registerReportsCommands(program: Command, getClient: () => GraphClient): void {
  const reports = program.command('reports').description('Aggregate activity reports (counts only: no message content, no per-person rows)');

  for (const [name, def] of Object.entries(COUNT_REPORTS)) {
    reports
      .command(name)
      .description(def.description)
      .addOption(new Option('--period <period>', 'Reporting window').choices([...PERIODS]).default('D30'))
      .addOption(outputOption())
      .action(handleErrors(withReportsHint(async (opts) => {
        const rows = (await readReport(getClient(), def.fn, opts.period)).map(normalize);
        console.log(formatOutput({ period: opts.period, report: name, value: rows }, opts.output as OutputFormat));
      })));
  }

  reports
    .command('by-department')
    .description(`Activity totals per department, from Microsoft's usage reports. Groups under ${MIN_GROUP_SIZE} people are pooled so nobody is singled out.`)
    .addOption(new Option('--workload <workload>', 'Which activity').choices(['email', 'teams']).default('email'))
    .addOption(new Option('--period <period>', 'Reporting window').choices([...PERIODS]).default('D30'))
    .option('--min-group-size <n>', `Smallest group shown (not below ${MIN_GROUP_SIZE})`, (v) => {
      const n = Number(v);
      if (!Number.isInteger(n) || n < MIN_GROUP_SIZE) {
        throw new InvalidArgumentError(`must be an integer of ${MIN_GROUP_SIZE} or more.`);
      }
      return n;
    }, MIN_GROUP_SIZE)
    .addOption(outputOption())
    .action(handleErrors(withReportsHint(async (opts) => {
      const client = getClient();
      const detail = DETAIL[opts.workload as 'email' | 'teams'];
      const details = await readReport(client, detail.fn, opts.period);

      const users = await client.requestAllPages('/users', { queryParams: { $select: 'userPrincipalName,department,accountEnabled', $top: '999' } });
      const departmentOf = new Map<string, string>();
      const enabledWithDept: string[] = [];
      for (const u of ((users.data as { value?: Array<{ userPrincipalName?: string; department?: string | null; accountEnabled?: boolean }> }).value ?? [])) {
        if (!u.userPrincipalName || !u.department) continue;
        const upn = u.userPrincipalName.toLowerCase();
        departmentOf.set(upn, u.department);
        if (u.accountEnabled !== false) enabledWithDept.push(upn);
      }

      if (client.isDryRun()) return; // the requests were traced; nothing to compute
      const { rows, unmatched } = rollUp(details, departmentOf, detail.fields, opts.minGroupSize);
      if (details.length > 0 && details.length === unmatched) {
        throw new CliError(
          'The usage report hides user names, so people cannot be matched to departments.',
          { code: 'concealed_names', exitCode: ExitCode.GENERIC, fix: 'an admin turns off "Display concealed user, group, and site names in all reports" (Microsoft 365 admin center → Settings → Org settings → Reports), or use the tenant-level reports' },
        );
      }
      if (unmatched > 0) info(`${unmatched} account(s) had no department and are not included.`);
      // Cross-check coverage (AO-23): a truncated or stale report silently undercounts departments.
      const reported = new Set(details.map(d => String(normalize(d).userprincipalname ?? '').toLowerCase()));
      const missing = enabledWithDept.filter(u => !reported.has(u)).length;
      if (missing > 0) info(`${missing} enabled account(s) with a department are not in the usage report; their activity is not included.`);
      console.log(formatOutput({ period: opts.period, workload: opts.workload, minGroupSize: opts.minGroupSize, value: rows }, opts.output as OutputFormat));
    })));
}
