import { Command } from 'commander';
import { handleErrors } from '../errors.js';
import { info } from '../utils.js';
import { buildTools, serve, spawnRunner } from '../mcp-server.js';

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
    .action(handleErrors(async (_opts, cmd: Command) => {
      const opts = cmd.optsWithGlobals();
      const readOnly = !!opts.readOnly || process.env.MS365_CLI_READ_ONLY === 'true' || process.env.MS365_CLI_READ_ONLY === '1';
      const allow = opts.allow ? String(opts.allow).split(',').map((p: string) => p.trim()).filter(Boolean) : undefined;
      const tools = buildTools(program, { allow, readOnly });
      info(`ms365 MCP server on stdio: ${tools.length} tools${readOnly ? ' (read-only)' : ''}`);
      await serve(process.stdin, (line) => process.stdout.write(`${line}\n`), {
        tools,
        version: program.version() ?? '0.0.0',
        runner: spawnRunner(),
        readOnly,
        globals: forwardedGlobals(program),
      });
    }));
}
