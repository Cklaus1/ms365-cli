import fs from 'fs';
import { Command } from 'commander';
import { usedScopes } from '../catalog.js';
import { activeCeiling, assertWithinCeiling } from '../ceiling.js';
import { CliError, ConfirmationRefusedError, ExitCode, handleErrors } from '../errors.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import type { GraphClient } from '../graph-client.js';
import { withIdempotency } from '../idempotency.js';
import { PLAN_COMMANDS, PLAN_CONFIRM, planDigest, readPlan, requestCommand, resolveRefs, sha256Hex, usage, type Plan } from '../plan.js';
import { profileType, readProfile } from '../profiles.js';
import { confirm, info, isDryRun, outputOption, success } from '../utils.js';

/**
 * Every check that needs no token: digest, command, ceiling (AH-5(7)), expiry, attachment bytes. Nothing is sent.
 * Returns the commands the plan's requests amount to (its label plus what each request is), for the confirmation.
 */
function checkPlan(plan: Plan, expectDigest: string, profileFlag: string | undefined): string[] {
  const digest = planDigest(plan);
  if (digest !== plan.digest || digest !== expectDigest) {
    throw new CliError(`The plan's digest is ${digest}, which does not match ${digest !== plan.digest ? 'its digest field' : '--expect-digest'}; nothing was sent.`,
      { code: 'plan_digest_mismatch', exitCode: ExitCode.USAGE, fix: 'apply the plan that was approved, unmodified' });
  }
  if (!PLAN_COMMANDS[plan.command]) throw usage(`"${plan.command}" cannot be applied from a plan.`, 'plan_invalid');
  // The label is the plan's own claim; the ceiling and the confirmation also cover what each request really does.
  const commands = new Set([plan.command]);
  for (const r of plan.requests) {
    const command = requestCommand(r);
    if (!command) throw usage(`The plan's request ${r.method} ${r.path} is not one a plan command sends; nothing was sent.`, 'plan_invalid');
    commands.add(command);
  }
  const ceiling = activeCeiling(profileFlag);
  if (ceiling) for (const c of commands) assertWithinCeiling(c, usedScopes(c), ceiling.ceiling, ceiling.profile);
  if (!(Date.parse(plan.expiresAt) > Date.now())) {
    throw new CliError(`The plan expired at ${plan.expiresAt}; nothing was sent.`, { code: 'plan_expired', exitCode: ExitCode.USAGE, fix: 'make a new plan' });
  }
  for (const a of plan.attachments) {
    let hash: string | undefined;
    try { hash = sha256Hex(fs.readFileSync(a.path)); } catch { /* missing counts as changed */ }
    if (hash !== a.sha256) {
      throw new CliError(`Attachment ${a.name} (${a.path}) is ${hash ? 'not the file that was planned' : 'missing'}; nothing was sent.`,
        { code: 'plan_attachment_changed', exitCode: ExitCode.USAGE, fix: 'make a new plan with the current file' });
    }
  }
  return [...commands];
}

/**
 * OQ-8 (b): without a terminal, a delete/cancel plan is applied only when the approval happened in the host. By the
 * time this runs, checkPlan has matched `--expect-digest`, the expiry and the ceiling for every request; this adds
 * that the active profile is `type: host`, that every request is the delete/cancel its label names (a mislabelled
 * plan, E27, was not what the host showed), and that the plan names the tenant and account the identity check
 * compares (done before sending). Anything else keeps exit 7 with nothing sent.
 */
function assertHostApproval(plan: Plan, commands: string[], profileFlag: string | undefined): void {
  const name = profileFlag || process.env.MS365_CLI_PROFILE;
  const refuse = (why: string) => new ConfirmationRefusedError(`Applying ${plan.command} without a terminal ${why}; nothing was sent.`);
  if (!name || profileType(readProfile(name)) !== 'host') {
    throw refuse('needs a type: host profile (the approval happens in the agent host and --expect-digest binds it); run it in a terminal otherwise');
  }
  if (commands.length !== 1 || commands[0] !== plan.command) {
    throw refuse(`is refused: its requests amount to ${commands.join(' and ')}, not only what its label says`);
  }
  if (!plan.identity.tenantId || !plan.identity.accountId) {
    throw refuse('is refused: the plan does not name the tenant and account it was made for');
  }
}

/** Send the plan's requests in order; `{{n.id}}` takes the id request n returned. */
async function sendRequests(client: GraphClient, plan: Plan): Promise<Record<string, unknown>> {
  const ids: (string | undefined)[] = [];
  const responses: Record<string, unknown>[] = [];
  for (const [i, r] of plan.requests.entries()) {
    const path = resolveRefs(r.path, ids, true);
    let data: unknown;
    let status: number;
    try {
      ({ data, status } = await client.request(path, {
        method: r.method, headers: r.headers, ...(r.body !== null && { body: resolveRefs(r.body, ids, false) }),
      }));
    } catch (err) {
      // An earlier request already took effect, so a retry is not safe: the ledger keeps the plan as unknown.
      if (i > 0 && err instanceof CliError) err.maybeProcessed = true;
      throw err;
    }
    const id: unknown = data && typeof data === 'object' ? Reflect.get(data, 'id') : undefined;
    ids.push(typeof id === 'string' ? id : undefined);
    responses.push({ method: r.method, path, status, ...(typeof id === 'string' && { id }) });
  }
  return { outcome: 'APPLIED', command: plan.command, digest: plan.digest, appliedAt: new Date().toISOString(), responses };
}

export function registerApplyCommand(program: Command, getClient: () => GraphClient): void {
  program
    .command('apply <planFile>')
    .description('Send exactly the requests of a plan written by --plan, after checking its digest, the profile ceiling, expiry, attachment bytes and token identity; If-Match is part of the plan (412 → exit 4 version_changed). A calendar delete/cancel plan asks on a terminal; without one it is applied only under a type: host profile (the host approved the digest), otherwise exit 7')
    .requiredOption('--expect-digest <sha256>', 'The plan digest that was approved; any other plan is refused')
    .addOption(outputOption('json'))
    .action(handleErrors(async (file: string, opts) => {
      if (!/^[0-9a-f]{64}$/.test(opts.expectDigest)) throw usage('--expect-digest must be a sha256 in lowercase hex.');
      const plan = readPlan(file);
      const commands = checkPlan(plan, opts.expectDigest, program.opts().profile);
      const destructive = commands.filter(c => PLAN_CONFIRM[c]);
      // A delete or cancel asks on a terminal (and previews under --dry-run); without one, only a host approval counts.
      const hostApproval = destructive.length > 0 && !process.stdin.isTTY && !isDryRun();
      if (hostApproval) assertHostApproval(plan, commands, program.opts().profile);
      else if (destructive.length) {
        const field = (key: string): unknown => Reflect.get(Object(plan.payload), key);
        const notifies = field('notifies');
        await confirm(PLAN_CONFIRM[plan.command]
          ? `Apply ${plan.command} ${String(field('eventId'))} (${String(field('scope'))}; notifies ${Array.isArray(notifies) ? notifies.length : '?'} attendee(s))?`
          : `Apply ${plan.command}, whose requests include ${destructive.join(' and ')}?`);
      }
      const client = getClient();
      if (client.isDryRun()) {
        await sendRequests(client, plan);
        return;
      }
      const who = await client.identity();
      const want = plan.identity;
      if ((want.tenantId && who.tenantId !== want.tenantId) || (want.accountId && who.accountId !== want.accountId)) {
        throw new CliError('The token belongs to another tenant or account than the plan was made for; nothing was sent.',
          { code: 'token_identity_mismatch', exitCode: ExitCode.AUTH, fix: 'apply the plan with the token of the account that made it' });
      }
      // The digest is the idempotency key (D23): a plan is applied at most once per ledger TTL.
      const outcome = await withIdempotency(`plan:${plan.digest}`, 'apply', plan.digest, () => sendRequests(client, plan));
      if (outcome.status === 'replay') info(`Plan ${plan.digest.slice(0, 12)}… was already applied at ${outcome.ts}; not sent again.`);
      if (opts.output === 'text') {
        if (outcome.status === 'done') success(`Applied ${plan.command}: ${plan.requests.length} request(s) sent.`);
        return;
      }
      console.log(formatOutput(outcome.result, opts.output as OutputFormat));
    }));
}
