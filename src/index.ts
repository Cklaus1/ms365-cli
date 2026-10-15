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

import { Command, Option } from 'commander';
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
import { registerMeetingsCommands } from './commands/meetings.js';
import { registerReportsCommands } from './commands/reports.js';
import { registerChiefCommands } from './commands/chief.js';
import { registerWatchCommands } from './commands/watch.js';
import { registerAuditCommands } from './commands/audit.js';
import { registerIdempotencyCommands } from './commands/idempotency.js';
import { registerDescribeCommand } from './commands/describe.js';
import { registerMcpCommands } from './commands/mcp.js';
import { registerCompletionCommand } from './commands/completion.js';
import { registerResolveCommands } from './commands/resolve.js';
import { registerNotesCommands } from './commands/notes.js';
import { setActor, setCommandPath } from './audit.js';
import { requiredScopes } from './catalog.js';
import { assertConsented } from './preflight.js';
import { profileType, readProfile, type Profile } from './profiles.js';
import logger from './logger.js';
import { setDryRun, setQuiet } from './utils.js';
import { setFields, setCursor } from './formatter.js';
import { reportError, setErrorFormat, ExitCode, CliError } from './errors.js';

const program = new Command();

program
  .name('ms365')
  .description('Microsoft 365 CLI - Interact with Microsoft 365 via the Graph API')
  .version('0.1.0')
  .option('--org-mode', 'Enable organization mode (Teams, SharePoint, shared mailboxes)')
  .option('--read-only', 'Restrict to read-only operations')
  .option('--dry-run', 'Preview the API request without executing it')
  .addOption(new Option('--api-version <version>', 'Graph API version').choices(['v1.0', 'beta']))
  .option('-q, --quiet', 'Suppress non-data output (confirmations, progress)')
  .addOption(new Option('--error-format <format>', 'Error output on stderr (default: text)').choices(['text', 'json']))
  .option('--fields <list>', 'Only output these comma-separated fields (dot paths allowed, e.g. id,from.emailAddress.address)')
  .option('--actor <name>', 'Name recorded as the actor in the audit log (default: MS365_CLI_ACTOR or "cli")')
  .option('--cursor <token>', 'Fetch the page after a previous list call (the nextCursor it printed on stderr)')
  .option('--account <upn|homeAccountId>', 'Use this cached account for this call only (env MS365_CLI_ACCOUNT)')
  .option('--profile <name>', 'Apply a saved profile: client ID, tenant, scopes, org mode, account (env MS365_CLI_PROFILE)');

// Usage errors (unknown option, missing argument) exit 2; help/version exit 0.
// Must be set before subcommands are registered so they inherit it.
program.exitOverride((err) => {
  const ok = err.code === 'commander.helpDisplayed' || err.code === 'commander.version' || err.code === 'commander.help' || err.exitCode === 0;
  process.exit(ok ? ExitCode.OK : ExitCode.USAGE);
});

const envTrue = (...names: string[]) => names.some(n => process.env[n] === 'true' || process.env[n] === '1');

const profileName = (): string | undefined => program.opts().profile || process.env.MS365_CLI_PROFILE || undefined;

/** The --profile (or MS365_CLI_PROFILE) settings; empty when none is chosen. */
function activeProfile(): Profile {
  const name = profileName();
  return name ? readProfile(name) : {};
}

/** The name of the active profile when it is an admin profile (ID-7/ID-8), else undefined. */
function adminProfileName(): string | undefined {
  const name = profileName();
  return name && profileType(readProfile(name)) === 'admin' ? name : undefined;
}

/** Global flags, resolved from parsed options (after parse) with env-var fallbacks, then the profile. */
function globalFlags() {
  const o = program.opts();
  const profile = activeProfile();
  const apiVersion = o.apiVersion || process.env.MS365_CLI_API_VERSION || 'v1.0';
  if (apiVersion !== 'v1.0' && apiVersion !== 'beta') {
    throw new CliError(`Invalid API version "${apiVersion}". Use v1.0 or beta.`, { code: 'usage', exitCode: ExitCode.USAGE });
  }
  return {
    orgMode: !!o.orgMode || envTrue('MS365_CLI_ORG_MODE') || !!profile.orgMode,
    account: o.account || process.env.MS365_CLI_ACCOUNT || profile.account,
    profile,
    dryRun: !!o.dryRun,
    readOnly: !!o.readOnly || envTrue('MS365_CLI_READ_ONLY'),
    apiVersion,
  };
}

/** Commands that never call Graph and must work without a client ID or login. */
const OFFLINE_COMMANDS = new Set(['auth', 'schema', 'generate-skills', 'describe', 'audit', 'idempotency', 'mcp', 'completion']);

// Lazy-initialized singletons
let authManager: AuthManager | null = null;
let graphClient: GraphClient | null = null;

function getAuth(): AuthManager {
  if (!authManager) {
    const { orgMode, account, profile } = globalFlags();
    const admin = adminProfileName();
    // Admin profiles (ID-8): the token is borrowed from the Azure CLI per call; the profile's clientId/orgMode are ignored.
    authManager = admin
      ? new AuthManager(buildScopes(!!program.opts().orgMode || envTrue('MS365_CLI_ORG_MODE')), { source: 'az', tenantId: profile.tenantId, profileName: admin })
      : new AuthManager(profile.scopes ?? buildScopes(orgMode), { clientId: profile.clientId, tenantId: profile.tenantId, account, profileName: profileName() });
  }
  return authManager;
}

function getClient(): GraphClient {
  if (!graphClient) {
    const { dryRun, apiVersion, readOnly, profile } = globalFlags();
    // A thunk: under --dry-run no AuthManager is built, so previews need no client ID (DR-16, F23).
    const clientId = adminProfileName() ? undefined : profile.clientId;
    graphClient = new GraphClient(getAuth, dryRun, apiVersion, readOnly,
      { clientId: clientId || process.env.MS365_CLI_CLIENT_ID || undefined });
  }
  return graphClient;
}

// Register all command modules
registerAuthCommands(program, getAuth, adminProfileName);
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
registerMeetingsCommands(program, getClient);
registerReportsCommands(program, getClient);
registerChiefCommands(program, getClient);
registerWatchCommands(program, getClient);
registerAuditCommands(program);
registerIdempotencyCommands(program);
registerNotesCommands(program, getClient);
registerResolveCommands(program, getClient);
registerDescribeCommand(program);
registerMcpCommands(program);
registerCompletionCommand(program);

// Ensure auth is initialized before non-auth commands
// The hook receives (hookedCommand = program, actionCommand); walk up from the action command.
program.hook('preAction', async (_program, actionCommand) => {
  const commandChain: string[] = [];
  let cmd: Command | null = actionCommand;
  while (cmd) {
    commandChain.unshift(cmd.name());
    cmd = cmd.parent;
  }

  const o = program.opts();
  if (o.errorFormat) setErrorFormat(o.errorFormat);
  if (o.quiet) setQuiet(true);
  // Before the offline-command return: local commands (auth logout, profile delete) honour --dry-run too (DR-22).
  if (o.dryRun) setDryRun(true);
  if (o.fields) setFields(String(o.fields).split(',').map((f: string) => f.trim()).filter(Boolean));
  if (o.cursor) setCursor(o.cursor);
  setActor(o.actor || process.env.MS365_CLI_ACTOR);
  const commandPath = commandChain.slice(1).join(' ');
  setCommandPath(commandPath);

  // Skip auth initialization for offline commands and dry-run
  if (OFFLINE_COMMANDS.has(commandChain[1])) return;
  if (globalFlags().dryRun) return;

  logger.info(`Running: ms365 ${commandChain.slice(1).join(' ')}`);

  const auth = getAuth();
  await auth.initialize();

  // Pre-flight consent check: fail fast (exit 4) instead of a mid-command 403.
  // Personal-account tokens are opaque (not JWTs) — then the check is skipped.
  // The clauses follow the parsed options and arguments (search --entity, watch <resource>, resolve <type>).
  const required = requiredScopes(commandPath, actionCommand.opts(), actionCommand.processedArgs);
  if (required.length > 0 && process.env.MS365_CLI_PREFLIGHT !== '0') {
    let granted: string[] | null = null;
    try {
      granted = await auth.grantedScopes();
    } catch (err) {
      if ((err as { code?: string }).code !== 'opaque_token') throw err;
    }
    if (granted) assertConsented(commandPath, required, granted, auth.tokenSource() === 'az' ? auth.getProfileName() ?? '' : undefined);
  }
});

program.parseAsync(process.argv).catch((err) => {
  logger.error(err instanceof Error ? err.message : String(err));
  process.exit(reportError(err));
});
