import dotenv from 'dotenv';
// quiet: dotenv 17 otherwise prints a banner to stdout, breaking `-o json` pipelines.
dotenv.config({ quiet: true });

import { Command, CommanderError } from 'commander';
import { AuthManager, buildScopes } from './auth.js';
import { GraphClient } from './graph-client.js';
import { exitProcess, reportError } from './errors.js';
import { registerAuthCommands } from './commands/auth.js';
import { registerMailCommands } from './commands/mail.js';
import { registerCalendarCommands } from './commands/calendar.js';
import { registerDriveCommands } from './commands/drive.js';
import { registerTodoCommands } from './commands/todo.js';
import { registerContactsCommands } from './commands/contacts.js';
import { registerTeamsCommands } from './commands/teams.js';
import { registerSharePointCommands } from './commands/sharepoint.js';
import { registerPlannerCommands } from './commands/planner.js';
import { registerOneNoteCommands } from './commands/onenote.js';
import { registerExcelCommands } from './commands/excel.js';
import { registerSearchCommands } from './commands/search.js';
import { registerUserCommands } from './commands/user.js';
import { registerWorkflowCommands } from './commands/workflow.js';
import { registerGroupsCommands } from './commands/groups.js';
import { registerSchemaCommands } from './commands/schema.js';
import { registerGenerateSkillsCommand } from './commands/generate-skills.js';
import { registerSetupCommand } from './commands/setup.js';
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
const NO_AUTH_COMMANDS: Record<string, true> = { auth: true, schema: true, 'generate-skills': true };

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
registerAuthCommands(program, getAuth);
registerMailCommands(program, getClient);
registerCalendarCommands(program, getClient);
registerDriveCommands(program, getClient);
registerTodoCommands(program, getClient);
registerContactsCommands(program, getClient);
registerTeamsCommands(program, getClient);
registerSharePointCommands(program, getClient);
registerPlannerCommands(program, getClient);
registerOneNoteCommands(program, getClient);
registerExcelCommands(program, getClient);
registerSearchCommands(program, getClient);
registerUserCommands(program, getClient);
registerWorkflowCommands(program, getClient);
registerGroupsCommands(program, getClient);
registerSchemaCommands(program);
registerGenerateSkillsCommand(program);
registerSetupCommand(program, getAuth);

// Apply global flags and load the token cache before any command that talks to Graph.
program.hook('preAction', async (_program, actionCommand) => {
  const g = globals();
  if (g.quiet) setQuiet(true);

  const commandChain: string[] = [];
  for (let cmd: Command | null = actionCommand; cmd && cmd !== program; cmd = cmd.parent) {
    commandChain.unshift(cmd.name());
  }
  logger.info(`Running: ms365 ${commandChain.join(' ')}`);

  if (NO_AUTH_COMMANDS[commandChain[0]] || g.dryRun) return;
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
