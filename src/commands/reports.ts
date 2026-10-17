import { Command, InvalidArgumentError, Option } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption, info, odataQuote } from '../utils.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';

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
 * Sum metrics by department. Departments with fewer than `min` people are pooled into
 * "Other (groups under N people)"; if that pool is itself under `min`, it is dropped, so no
 * output row ever describes fewer than `min` people. Per-person values are never returned.
 */
export function rollUp(
  details: Row[],
  departmentOf: Map<string, string>,
  fields: readonly string[],
  min = MIN_GROUP_SIZE,
  excluded: ReadonlySet<string> = new Set(),
): { rows: RollUpRow[]; unmatched: number; excluded: number } {
  const groups = new Map<string, { people: number; sums: Record<string, number> }>();
  let unmatched = 0;
  let skipped = 0;
  for (const raw of details) {
    const r = normalize(raw);
    const upn = String(r.userprincipalname ?? '').toLowerCase();
    if (!upn || /^[0-9a-f]{16,}$/i.test(upn)) { unmatched++; continue; } // concealed names can't be joined
    if (excluded.has(upn)) { skipped++; continue; }
    const dept = departmentOf.get(upn);
    if (!dept) { unmatched++; continue; }
    const g = groups.get(dept) ?? { people: 0, sums: Object.fromEntries(fields.map(f => [f, 0])) };
    g.people++;
    for (const f of fields) g.sums[f] += num(r[f]);
    groups.set(dept, g);
  }
  const rows: RollUpRow[] = [];
  const pool = { people: 0, sums: Object.fromEntries(fields.map(f => [f, 0])) as Record<string, number> };
  for (const [department, g] of groups) {
    if (g.people >= min) rows.push({ department, people: g.people, ...g.sums });
    else { pool.people += g.people; for (const f of fields) pool.sums[f] += g.sums[f]; }
  }
  if (pool.people >= min) rows.push({ department: `Other (groups under ${min} people)`, people: pool.people, ...pool.sums });
  rows.sort((a, b) => b.people - a.people);
  return { rows, unmatched, excluded: skipped };
}

async function jsonReport(client: GraphClient, fn: string, period: string): Promise<Row[]> {
  const res = await client.request(`/reports/${fn}(period='${odataQuote(period)}')`, { queryParams: { $format: 'application/json' }, responseType: 'json' });
  if (client.isPreview(res)) return [];
  return ((res.data as { value?: Row[] }).value ?? []);
}

export function registerReportsCommands(program: Command, getClient: () => GraphClient): void {
  const reports = program.command('reports').description('Aggregate activity reports (counts only: no message content, no per-person rows)');

  for (const [name, def] of Object.entries(COUNT_REPORTS)) {
    reports
      .command(name)
      .description(def.description)
      .addOption(new Option('--period <period>', 'Reporting window').choices([...PERIODS]).default('D30'))
      .addOption(outputOption())
      .action(handleErrors(async (opts) => {
        const rows = (await jsonReport(getClient(), def.fn, opts.period)).map(normalize);
        console.log(formatOutput({ period: opts.period, report: name, value: rows }, opts.output as OutputFormat));
      }));
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
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const detail = DETAIL[opts.workload as 'email' | 'teams'];
      const details = await jsonReport(client, detail.fn, opts.period);

      const users = await client.requestAllPages('/users', { queryParams: { $select: 'userPrincipalName,department', $top: '999' } });
      const departmentOf = new Map<string, string>();
      for (const u of ((users.data as { value?: Array<{ userPrincipalName?: string; department?: string | null }> }).value ?? [])) {
        if (u.userPrincipalName && u.department) departmentOf.set(u.userPrincipalName.toLowerCase(), u.department);
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
      console.log(formatOutput({ period: opts.period, workload: opts.workload, minGroupSize: opts.minGroupSize, value: rows }, opts.output as OutputFormat));
    }));
}
