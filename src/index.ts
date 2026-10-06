import dotenv from 'dotenv';
// quiet: dotenv 17 otherwise prints a banner to stdout, breaking `-o json` pipelines.
dotenv.config({ quiet: true });

import { Command, CommanderError } from 'commander';
import { AuthManager, buildScopes, envToken } from './auth.js';
import { GraphClient } from './graph-client.js';
import { exitProcess, reportError } from './errors.js';
import { registerAllCommands } from './commands/registry.js';
import logger from './logger.js';
import { setQuiet } from './utils.js';

// Exit immediately on broken pipe - prevents zombie processes when piped to head/tail/etc.
for (const stream of [process.stdout, process.stderr]) {
  stream.on('error', (err: NodeJS.ErrnoException) => {
    if (err.code === 'EPIPE' || err.code === 'ERR_STREAM_DESTROYED') process.exit(0);
  });
}

const envFlag = (...names: string[]) => names.some(n => process.env[n] === 'true' || process.env[n] === '1');

interface GlobalOptions {
  orgMode?: boolean;
  readOnly?: boolean;
  dryRun?: boolean;
  apiVersion: string;
  quiet?: boolean;
}

/** Commands that run without a client ID or token cache. */
const NO_AUTH_COMMANDS: Record<string, true> = { auth: true, schema: true, 'generate-skills': true, serve: true };

const program = new Command();

program
  .name('ms365')
  .description('Microsoft 365 CLI - Interact with Microsoft 365 via the Graph API')
  .version('0.1.0')
  .option('--org-mode', 'Enable organization mode (Teams, SharePoint, shared mailboxes)')
  .option('--read-only', 'Restrict to read-only operations')
  .option('--dry-run', 'Preview the API request without executing it')
  .option('--api-version <version>', 'Graph API version', process.env.MS365_CLI_API_VERSION || 'v1.0')
  .option('-q, --quiet', 'Suppress non-data output (confirmations, progress)')
  // Global flags must precede the subcommand, so an argument value like `--body --dry-run` can never flip a mode.
  .enablePositionalOptions()
  .exitOverride();

/** Global flags after Commander parsing; env vars fill in what the command line leaves unset. */
function globals(): GlobalOptions {
  const opts = program.opts<GlobalOptions>();
  return {
    ...opts,
    orgMode: opts.orgMode || envFlag('MS365_CLI_ORG_MODE', 'MS365_MCP_ORG_MODE'),
    readOnly: opts.readOnly || envFlag('MS365_CLI_READ_ONLY'),
  };
}

// Lazy-initialized singletons, built from parsed global options on first use.
let authManager: AuthManager | null = null;
let graphClient: GraphClient | null = null;

function getAuth(): AuthManager {
  if (!authManager) {
    authManager = new AuthManager(buildScopes(Boolean(globals().orgMode)));
  }
  return authManager;
}

function getClient(): GraphClient {
  if (!graphClient) {
    const g = globals();
    graphClient = new GraphClient(getAuth(), Boolean(g.dryRun), g.apiVersion, Boolean(g.readOnly));
  }
  return graphClient;
}

// Register all command modules
registerAllCommands(program, getAuth, getClient);

// Apply global flags and load the token cache before any command that talks to Graph (an env token needs no cache).
program.hook('preAction', async (_program, actionCommand) => {
  const g = globals();
  if (g.quiet) setQuiet(true);

  const commandChain: string[] = [];
  for (let cmd: Command | null = actionCommand; cmd && cmd !== program; cmd = cmd.parent) {
    commandChain.unshift(cmd.name());
  }
  logger.info(`Running: ms365 ${commandChain.join(' ')}`);

  if (NO_AUTH_COMMANDS[commandChain[0]] || g.dryRun || envToken()) return;
  await getAuth().initialize();
});

async function main(): Promise<number> {
  try {
    await program.parseAsync(process.argv);
    return typeof process.exitCode === 'number' ? process.exitCode : 0;
  } catch (err) {
    // --help / --version throw with exitCode 0; usage errors were already printed by Commander.
    if (err instanceof CommanderError) return err.exitCode;
    reportError(err);
    return 1;
  }
}

main().then(exitProcess);
