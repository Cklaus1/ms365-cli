/**
 * `ms365 health` (AH-8(2)-(3), E7, E12): one bounded probe per resource, reported in the host's `connector-health`
 * vocabulary. The status of every resource is in-band (exit 0); only a bad invocation exits non-zero.
 */
import { Option, type Command } from 'commander';
import { activeCeiling, unmetClause } from '../ceiling.js';
import { CliError, ExitCode, handleErrors, toCliError } from '../errors.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { requestBudget, type GraphClient } from '../graph-client.js';
import { isDryRun, outputOption } from '../utils.js';
import { checkpointText, decodeCheckpoint, defaultCursor, type Cursor } from './watch.js';

export type HealthStatus = 'HEALTHY' | 'PARTIAL' | 'STALE' | 'AUTH_REQUIRED' | 'RATE_LIMITED' | 'UNSUPPORTED' | 'DISABLED';

const PROBES = {
  mail: { scope: 'Mail.Read', endpoint: '/me/messages' },
  events: { scope: 'Calendars.Read', endpoint: '/me/events' },
};
type HealthResource = keyof typeof PROBES;

/** A probe never waits longer than this, and never retries (R-16). */
const PROBE_TIMEOUT_MS = 10_000;

/**
 * The fixed ms365 → `connector-health` mapping (AH-8(3), OQ-13, OQ-19) of a run's outcome: an error, or a read
 * result / watch line that may report a coverage hole. Hosts apply the same table to any ms365 run; `DISABLED` is
 * host-only and never produced here.
 */
export function healthOf(signal: unknown): { status: HealthStatus; reason: string | null } {
  if (signal && typeof signal === 'object' && !(signal instanceof Error)) {
    if ('type' in signal && signal.type === 'gap') {
      return { status: 'PARTIAL', reason: 'reason' in signal && typeof signal.reason === 'string' ? signal.reason : 'gap' };
    }
    if ('complete' in signal && signal.complete === false) {
      return { status: 'PARTIAL', reason: 'incompleteReason' in signal && typeof signal.incompleteReason === 'string' ? signal.incompleteReason : 'incomplete' };
    }
    return { status: 'HEALTHY', reason: null };
  }
  const e = toCliError(signal);
  if (e.code === 'scope_ceiling') return { status: 'UNSUPPORTED', reason: 'scope_ceiling' };
  if (e.code === 'missing_scope') return { status: 'AUTH_REQUIRED', reason: 'missing_scope' };
  if (e.code === 'permission_denied') return { status: 'AUTH_REQUIRED', reason: 'access_denied' };
  if (e.code === 'delta_expired') return { status: 'PARTIAL', reason: 'delta_expired' };
  if (e.code === 'cursor_mismatch' || e.code === 'checkpoint_mismatch' || e.exitCode === ExitCode.USAGE) return { status: 'UNSUPPORTED', reason: 'usage' };
  if (e.exitCode === ExitCode.AUTH) return { status: 'AUTH_REQUIRED', reason: e.code };
  if (e.exitCode === ExitCode.THROTTLED) return { status: 'RATE_LIMITED', reason: e.code };
  // graph_error (5xx and anything else unexpected), timeout, network_error, bad_response, error.
  return { status: 'STALE', reason: e.code };
}

interface HealthRecord {
  resource: HealthResource;
  status: HealthStatus;
  reason: string | null;
  checkedAt: string;
  lastSuccessAt: string | null;
  coverageStart: string | null;
  coverageEnd: string | null;
  checkpoint: string | null;
}

export function registerHealthCommand(program: Command, getClient: () => GraphClient): void {
  program
    .command('health')
    .description('Report each resource\'s health for an agent host (HEALTHY, PARTIAL, STALE, AUTH_REQUIRED, RATE_LIMITED, UNSUPPORTED)')
    .argument('[resources...]', 'mail, events (default: both)')
    .addOption(new Option('--checkpoint <blob|@file|@->', 'A checkpoint printed by watch; its coverage replaces the saved cursor\'s for its resource'))
    .addOption(outputOption('json'))
    .addHelpText('after', `
Probes: GET /me, then one $top=1 read per resource, each with no retries and at most 10 s (or
MS365_CLI_REQUEST_TIMEOUT_MS if lower), so health finishes even when Graph does not answer. A resource
outside the profile's scope ceiling is UNSUPPORTED (scope_ceiling) and is not probed. Without
--checkpoint, coverage comes from the cursor of a plain "ms365 watch <resource>".`)
    .action(handleErrors(async (names: string[], opts) => {
      const resources: HealthResource[] = [];
      for (const name of names.length ? names : ['mail', 'events']) {
        if (name !== 'mail' && name !== 'events') {
          throw new CliError(`Unknown health resource "${name}": use mail or events.`, { code: 'usage', exitCode: ExitCode.USAGE });
        }
        if (!resources.includes(name)) resources.push(name);
      }
      const blob = typeof opts.checkpoint === 'string' ? checkpointText(opts.checkpoint) : undefined;
      const held = blob === undefined ? undefined : decodeCheckpoint(blob);
      if (held && !resources.some(r => r === held.resource)) {
        throw new CliError(`--checkpoint is for ${held.resource}, which is not among the resources checked.`, { code: 'checkpoint_mismatch', exitCode: ExitCode.USAGE });
      }

      const client = getClient();
      const probe = { retries: 0, timeoutMs: Math.min(PROBE_TIMEOUT_MS, requestBudget().timeoutMs) };
      const active = activeCeiling(program.opts().profile);
      const records: HealthRecord[] = [];
      let me: unknown;
      for (const resource of resources) {
        const checkedAt = new Date().toISOString();
        const { scope, endpoint } = PROBES[resource];
        let outcome: { status: HealthStatus; reason: string | null };
        if (active && unmetClause([{ scopes: [scope] }], active.ceiling)) {
          outcome = { status: 'UNSUPPORTED', reason: 'scope_ceiling' };
        } else {
          // GET /me once: a failure there (auth, outage) is every resource's state.
          if (me === undefined) me = await client.request('/me', { ...probe, queryParams: { $select: 'id' } }).then(r => r.data, err => toCliError(err));
          let signal: unknown = me;
          if (!(me instanceof Error)) {
            signal = await client.request(endpoint, { ...probe, queryParams: { $top: '1', $select: 'id' } }).then(r => r.data, err => toCliError(err));
          }
          outcome = healthOf(signal);
        }
        const cursor: Cursor | undefined = held?.resource === resource ? held : defaultCursor(resource);
        records.push({
          resource, ...outcome, checkedAt,
          lastSuccessAt: cursor?.syncedAt ?? null,
          coverageStart: cursor?.coverageStart ?? null,
          coverageEnd: cursor?.syncedAt ?? null,
          checkpoint: held?.resource === resource ? blob ?? null : null,
        });
      }
      if (isDryRun()) return; // the probes were traced; there is no state to report
      console.log(formatOutput(records, opts.output as OutputFormat));
    }));
}
