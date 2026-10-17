import { Command, Option } from 'commander';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import { commandMeta, type ScopeSelector } from '../catalog.js';
import { activeCeiling, ceilingSurface } from '../ceiling.js';
import { PLAN_COMMANDS } from '../plan.js';
import type { CatalogCommand, CatalogOption } from '../docs-gen.js';

/** Commands Graph itself de-duplicates (AH-6(5)), so a host may retry them after a lost response. */
const PROVIDER_IDEMPOTENCY: Record<string, string> = {
  'calendar create': 'With --idempotency-key the key is sent as the event transactionId; Graph returns the first event for a repeat, so a retry after a lost response creates no second event.',
};

export function describeOption(o: Option): CatalogOption {
  return {
    flags: o.flags,
    long: o.long ?? null,
    short: o.short ?? null,
    description: o.description,
    required: o.mandatory,
    takesValue: o.required || o.optional,
    choices: o.argChoices ?? null,
    default: o.defaultValue ?? null,
  };
}

/** A selector with its positional argument named, as `describe` prints it. */
export function describeSelector(cmd: Command, sel: ScopeSelector | null | undefined) {
  if (!sel) return null;
  if ('option' in sel) return { option: sel.option, map: sel.map };
  return { argument: cmd.registeredArguments[sel.arg]?.name() ?? String(sel.arg), map: sel.map, ...(sel.withOption ? { withOption: sel.withOption } : {}) };
}

export function describeCommand(cmd: Command, path: string): CatalogCommand {
  const meta = commandMeta(path);
  return {
    path,
    description: cmd.description(),
    arguments: cmd.registeredArguments.map(a => ({ name: a.name(), required: a.required, variadic: a.variadic })),
    options: cmd.options.map(describeOption),
    scopes: meta?.scopes ?? null,
    scopesBy: describeSelector(cmd, meta?.scopesBy),
    destructive: meta?.destructive ?? null,
    idempotent: meta?.idempotent ?? null,
    offline: meta?.offline ?? null,
    plan: !!PLAN_COMMANDS[path],
    providerIdempotency: PROVIDER_IDEMPOTENCY[path] ?? null,
  };
}

/** Every leaf command under `program`, walked live so the catalogue never drifts from the registered tree. */
function leafCommands(program: Command): Array<{ cmd: Command; path: string }> {
  const out: Array<{ cmd: Command; path: string }> = [];
  const walk = (c: Command, chain: string[]) => {
    if (c.commands.length === 0) out.push({ cmd: c, path: chain.join(' ') });
    for (const sub of c.commands) walk(sub, [...chain, sub.name()]);
  };
  for (const top of program.commands) walk(top, [top.name()]);
  return out;
}

function summaryLine(c: CatalogCommand & { usage: string }, width: number): string {
  const by = c.scopesBy && ('option' in c.scopesBy ? c.scopesBy.option : `<${c.scopesBy.argument}>`);
  const scopes = c.scopes?.join(' ') + (by ? ` (by ${by})` : '');
  const tags = [c.destructive && 'destructive', c.offline ? 'offline' : scopes].filter(Boolean).join(', ');
  return `${c.usage.padEnd(width)}  ${c.description}${tags ? `  (${tags})` : ''}`;
}

export function registerDescribeCommand(program: Command): void {
  program
    .command('describe [command...]')
    .description('Describe every command (arguments, options, Graph scopes, safety) as a machine-readable catalogue')
    .option('--json', 'Print the catalogue as JSON')
    .option('--all', 'Under a profile scope ceiling, also list the commands it does not allow (marked allowed: false)')
    .action(handleErrors(async (words: string[], opts: { json?: boolean; all?: boolean }) => {
      const prefix = words.join(' ');
      const leaves = leafCommands(program).filter(({ path }) => !prefix || path === prefix || path.startsWith(`${prefix} `));
      if (leaves.length === 0) {
        throw new CliError(`Unknown command: ${prefix}`, { code: 'unknown_command', exitCode: ExitCode.USAGE, fix: 'ms365 describe' });
      }
      // AH-1(5): under a ceiling profile the catalogue is the effective surface; --all keeps the rest, marked.
      const ceiling = activeCeiling(program.opts().profile)?.ceiling;
      const entries = leaves
        .map(leaf => ({ leaf, c: describeCommand(leaf.cmd, leaf.path) }))
        .map(e => (ceiling ? { ...e, c: { ...e.c, allowed: ceilingSurface(e.c.path, ceiling).allowed } } : e))
        .filter(e => opts.all || !('allowed' in e.c) || e.c.allowed);
      const commands = entries.map(e => e.c);

      if (opts.json) {
        console.log(JSON.stringify({ version: program.version() ?? null, globalOptions: program.options.map(describeOption), commands }, null, 2));
        return;
      }
      const rows = entries.map(({ leaf, c }) => ({ ...c, usage: [c.path, leaf.cmd.usage().replace(/^\[options\]\s*/, '')].join(' ').trim() }));
      const width = Math.max(0, ...rows.map(r => r.usage.length));
      for (const r of rows) console.log(summaryLine(r, width));
    }));
}
