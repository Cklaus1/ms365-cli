import { Command, Option } from 'commander';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import { commandMeta } from '../catalog.js';

function describeOption(o: Option) {
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

function describeCommand(cmd: Command, path: string) {
  const meta = commandMeta(path);
  return {
    path,
    description: cmd.description(),
    arguments: cmd.registeredArguments.map(a => ({ name: a.name(), required: a.required, variadic: a.variadic })),
    options: cmd.options.map(describeOption),
    scopes: meta?.scopes ?? null,
    destructive: meta?.destructive ?? null,
    idempotent: meta?.idempotent ?? null,
    offline: meta?.offline ?? null,
  };
}

type Described = ReturnType<typeof describeCommand>;

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

function summaryLine(c: Described & { usage: string }, width: number): string {
  const tags = [c.destructive && 'destructive', c.offline ? 'offline' : c.scopes?.join(' ')].filter(Boolean).join(', ');
  return `${c.usage.padEnd(width)}  ${c.description}${tags ? `  (${tags})` : ''}`;
}

export function registerDescribeCommand(program: Command): void {
  program
    .command('describe [command...]')
    .description('Describe every command (arguments, options, Graph scopes, safety) as a machine-readable catalogue')
    .option('--json', 'Print the catalogue as JSON')
    .action(handleErrors(async (words: string[], opts: { json?: boolean }) => {
      const prefix = words.join(' ');
      const leaves = leafCommands(program).filter(({ path }) => !prefix || path === prefix || path.startsWith(`${prefix} `));
      if (leaves.length === 0) {
        throw new CliError(`Unknown command: ${prefix}`, { code: 'unknown_command', exitCode: ExitCode.USAGE, fix: 'ms365 describe' });
      }
      const commands = leaves.map(({ cmd, path }) => describeCommand(cmd, path));

      if (opts.json) {
        console.log(JSON.stringify({ version: program.version() ?? null, globalOptions: program.options.map(describeOption), commands }, null, 2));
        return;
      }
      const rows = commands.map((c, i) => ({ ...c, usage: [c.path, leaves[i].cmd.usage().replace(/^\[options\]\s*/, '')].join(' ').trim() }));
      const width = Math.max(...rows.map(r => r.usage.length));
      for (const r of rows) console.log(summaryLine(r, width));
    }));
}
