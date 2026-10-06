import 'dotenv/config';

// Exit immediately on broken pipe — prevents zombie processes when piped to head/tail/etc.
process.stdout.on('error', (err) => {
  if (err.code === 'EPIPE' || err.code === 'ERR_STREAM_DESTROYED') process.exit(0);
});
process.stderr.on('error', (err) => {
  if (err.code === 'EPIPE' || err.code === 'ERR_STREAM_DESTROYED') process.exit(0);
});

// Force stdout to flush after every write — prevents buffering when not a TTY
// (e.g., when run from scripts, Claude Code, or piped to another process)
const origWrite = process.stdout.write.bind(process.stdout);
process.stdout.write = function(chunk: any, ...args: any[]): boolean {
  const result = origWrite(chunk, ...args);
  try { (process.stdout as any)._handle?.flush?.(); } catch {}
  return result;
};

import { Command } from 'commander';
import { AuthManager, buildScopes } from './auth.js';
import { GraphClient } from './graph-client.js';
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

const program = new Command();

program
  .name('ms365')
  .description('Microsoft 365 CLI - Interact with Microsoft 365 via the Graph API')
  .version('0.1.0')
  .option('--org-mode', 'Enable organization mode (Teams, SharePoint, shared mailboxes)')
  .option('--read-only', 'Restrict to read-only operations')
  .option('--dry-run', 'Preview the API request without executing it')
  .option('--api-version <version>', 'Graph API version (default: v1.0)', 'v1.0')
  .option('-q, --quiet', 'Suppress non-data output (confirmations, progress)');

// Determine flags from env or argv
const orgMode =
  process.argv.includes('--org-mode') ||
  process.env.MS365_CLI_ORG_MODE === 'true' ||
  process.env.MS365_CLI_ORG_MODE === '1' ||
  process.env.MS365_MCP_ORG_MODE === 'true' ||
  process.env.MS365_MCP_ORG_MODE === '1';

const dryRun = process.argv.includes('--dry-run');
const readOnly =
  process.argv.includes('--read-only') ||
  process.env.MS365_CLI_READ_ONLY === 'true' ||
  process.env.MS365_CLI_READ_ONLY === '1';

if (process.argv.includes('--quiet') || process.argv.includes('-q')) {
  setQuiet(true);
}

// Extract --api-version value (CLI flag takes precedence over env var)
let apiVersion = process.env.MS365_CLI_API_VERSION || 'v1.0';
const apiIdx = process.argv.indexOf('--api-version');
if (apiIdx !== -1 && process.argv[apiIdx + 1]) {
  apiVersion = process.argv[apiIdx + 1];
}

const scopes = buildScopes(orgMode);

// Lazy-initialized singletons
let authManager: AuthManager | null = null;
let graphClient: GraphClient | null = null;

function getAuth(): AuthManager {
  if (!authManager) {
    authManager = new AuthManager(scopes);
  }
  return authManager;
}

function getClient(): GraphClient {
  if (!graphClient) {
    const auth = getAuth();
    graphClient = new GraphClient(auth, dryRun, apiVersion, readOnly);
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

// Ensure auth is initialized before non-auth commands
program.hook('preAction', async (thisCommand) => {
  const commandChain: string[] = [];
  let cmd: Command | null = thisCommand;
  while (cmd) {
    commandChain.unshift(cmd.name());
    cmd = cmd.parent;
  }

  // Skip auth initialization for auth commands and dry-run
  if (commandChain.includes('auth')) return;
  if (dryRun) return;

  logger.info(`Running: ms365 ${commandChain.slice(1).join(' ')}`);

  const auth = getAuth();
  await auth.initialize();
});

program.parseAsync(process.argv).catch((err) => {
  logger.error(err.message);
  console.error('Error:', err.message);
  process.exit(1);
});
