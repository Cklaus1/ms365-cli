import { Command } from 'commander';
import { CliError, ExitCode, handleErrors } from '../errors.js';
import { info } from '../utils.js';
import { buildTools, helperTokenEnv, serve, spawnRunner } from '../mcp-server.js';
import { activeCeiling } from '../ceiling.js';

/** Same floor as `watch --interval`. */
const MIN_WATCH_INTERVAL_S = 15;

/** Global flags given to `ms365 ... mcp serve` that every tool call inherits. */
function forwardedGlobals(program: Command): string[] {
  const o = program.opts();
  const out: string[] = [];
  if (o.orgMode) out.push('--org-mode');
  if (o.dryRun) out.push('--dry-run');
  for (const [key, flag] of [['apiVersion', '--api-version'], ['actor', '--actor'], ['account', '--account'], ['profile', '--profile']] as const) {
    if (o[key]) out.push(`${flag}=${o[key]}`);
  }
  return out;
}

export function registerMcpCommands(program: Command): void {
  const mcp = program.command('mcp').description('Model Context Protocol server');

  mcp
    .command('serve')
    .description('Serve every command as an MCP tool over stdio (JSON-RPC; each call runs the CLI)')
    .option('--allow <patterns>', 'Expose only tools matching these comma-separated patterns (e.g. "mail_*,calendar_list")')
    .option('--read-only', 'Expose only read tools and run every call with --read-only')
    .option('--watch-interval <seconds>', `Seconds between passes for subscribed watch resources (min ${MIN_WATCH_INTERVAL_S})`, '60')
    .action(handleErrors(async (_opts, cmd: Command) => {
      const opts = cmd.optsWithGlobals();
      const watchInterval = Number(opts.watchInterval);
      if (!Number.isFinite(watchInterval) || watchInterval < MIN_WATCH_INTERVAL_S) {
        throw new CliError(`--watch-interval must be at least ${MIN_WATCH_INTERVAL_S} seconds.`, { code: 'usage', exitCode: ExitCode.USAGE });
      }
      const readOnly = !!opts.readOnly || process.env.MS365_CLI_READ_ONLY === 'true' || process.env.MS365_CLI_READ_ONLY === '1';
      const allow = opts.allow ? String(opts.allow).split(',').map((p: string) => p.trim()).filter(Boolean) : undefined;
      // AH-1(4): under a ceiling profile only the commands it allows are tools; each call re-checks it (--profile is forwarded).
      const ceiling = activeCeiling(program.opts().profile);
      const tools = buildTools(program, { allow, readOnly, ceiling: ceiling?.ceiling });
      info(`ms365 MCP server on stdio: ${tools.length} tools${readOnly ? ' (read-only)' : ''}${ceiling ? ` (ceiling of profile "${ceiling.profile}")` : ''}`);
      await serve(process.stdin, (line) => process.stdout.write(`${line}\n`), {
        tools,
        version: program.version() ?? '0.0.0',
        // AH-2: with a token helper the server runs it once per token lifetime and hands the token to each call.
        runner: spawnRunner(undefined, undefined, process.env.MS365_CLI_TOKEN_COMMAND ? helperTokenEnv(process.env.MS365_CLI_TOKEN_COMMAND) : undefined),
        readOnly,
        globals: forwardedGlobals(program),
        watchIntervalMs: watchInterval * 1000,
        log: info,
      });
    }));
}
