import { Command } from 'commander';
import { exec } from 'child_process';
import { AuthManager, explainLoginError, tokenInfo } from '../auth.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { confirm, dryRunTrace, info, isDryRun, outputOption } from '../utils.js';
import { deleteProfile, listProfiles, profileExists, readProfile, saveProfile, validateProfileName, type Profile } from '../profiles.js';
import { openBrowserCommand } from './setup.js';

/** Open the system browser at the sign-in URL; print it too, for machines where no browser opens. */
async function openSystemBrowser(url: string): Promise<void> {
  info(`Opening the browser to sign in. If it does not open, visit:\n${url}`);
  exec(openBrowserCommand(url), () => { /* the URL was printed above */ });
}

export function registerAuthCommands(program: Command, getAuth: () => AuthManager): void {
  const auth = program.command('auth').description('Authentication management');

  auth
    .command('login')
    .description('Login (device code flow by default)')
    .option('--browser', 'Sign in with the system browser instead (needs an http://localhost redirect URI on the app registration)')
    .action(handleErrors(async (opts) => {
      if (isDryRun()) return dryRunTrace('would sign in');
      const authManager = getAuth();
      await authManager.initialize();
      try {
        await authManager.login(opts.browser ? { openBrowser: openSystemBrowser } : {});
      } catch (err) {
        throw explainLoginError(err, authManager.requestsOrgScopes());
      }
      const result = await authManager.verifyLogin();
      if (result.success && result.user) {
        info(`Logged in as ${result.user.displayName} (${result.user.email})`);
      } else {
        info('Login completed. Run "ms365 auth status" to verify.');
      }
    }));

  auth
    .command('logout')
    .description('Log out and clear saved credentials')
    .action(handleErrors(async () => {
      if (isDryRun()) return dryRunTrace('would log out and clear the token cache and selected account');
      const authManager = getAuth();
      await authManager.initialize();
      await authManager.logout();
      info('Logged out successfully.');
    }));

  auth
    .command('status')
    .description('Verify current login status')
    .action(handleErrors(async () => {
      // Offline under --dry-run: no AuthManager, so no token-cache read or refresh (DR-22 class).
      if (isDryRun()) return dryRunTrace('would check the login status (token cache not read)');
      const authManager = getAuth();
      await authManager.initialize();
      const result = await authManager.verifyLogin();
      if (result.success && result.user) {
        console.log(`Logged in as ${result.user.displayName} (${result.user.email})`);
      } else {
        throw new CliError('Not logged in.', { code: 'not_logged_in', exitCode: ExitCode.AUTH, fix: 'ms365 auth login' });
      }
    }));

  auth
    .command('scopes')
    .description('Show the scopes granted to the current access token (plus audience, tenant, user, expiry)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      if (isDryRun()) return dryRunTrace('would show the scopes of the current access token (token cache not read)');
      const authManager = getAuth();
      await authManager.initialize();
      console.log(formatOutput(tokenInfo(await authManager.getToken()), opts.output as OutputFormat));
    }));

  auth
    .command('refresh')
    .description('Fetch a fresh access token now (e.g. after an admin grants new permissions) and show which scopes changed')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      if (isDryRun()) return dryRunTrace('would refresh the access token (token cache not read)');
      if (process.env.MS365_CLI_TOKEN) {
        throw new CliError('MS365_CLI_TOKEN is set, so there is no cached login to refresh.',
          { code: 'usage', exitCode: ExitCode.USAGE, fix: 'unset MS365_CLI_TOKEN, or supply a new token' });
      }
      const authManager = getAuth();
      await authManager.initialize();
      let before: string[] = [];
      try { before = tokenInfo(await authManager.getToken()).scopes; } catch { /* no usable cached token */ }
      const after = tokenInfo(await authManager.getToken(true));
      const added = after.scopes.filter(s => !before.includes(s));
      const removed = before.filter(s => !after.scopes.includes(s));
      console.log(formatOutput({
        user: after.user, expiresOn: after.expiresOn, scopes: after.scopes.length, added, removed,
      }, opts.output as OutputFormat));
    }));

  auth
    .command('accounts')
    .description('List all cached accounts')
    .action(handleErrors(async () => {
      const authManager = getAuth();
      await authManager.initialize();
      const accounts = await authManager.listAccounts();
      if (accounts.length === 0) {
        info('No accounts found. Run: ms365 auth login');
        return;
      }

      const activeId = await authManager.activeAccountId();
      const override = authManager.getAccountOverride();
      if (override && !activeId) info(`Warning: account "${override}" (from --account, MS365_CLI_ACCOUNT or the profile) matches no cached account.`);
      for (const acc of accounts) {
        const marker = acc.homeAccountId === activeId ? ' (active)' : '';
        console.log(`  ${acc.username}${marker}`);
        console.log(`    ID: ${acc.homeAccountId}`);
      }
    }));

  auth
    .command('select <accountId>')
    .description('Select an account by ID')
    .action(handleErrors(async (accountId: string) => {
      if (isDryRun()) return dryRunTrace(`would select account ${accountId}`);
      const authManager = getAuth();
      await authManager.initialize();
      const ok = await authManager.selectAccount(accountId);
      if (ok) {
        info('Account selected.');
      } else {
        throw new CliError('Account not found.', { code: 'not_found', exitCode: ExitCode.NOT_FOUND });
      }
    }));

  auth
    .command('remove <accountId>')
    .description('Remove a cached account')
    .action(handleErrors(async (accountId: string) => {
      if (isDryRun()) return dryRunTrace(`would remove account ${accountId}`);
      const authManager = getAuth();
      await authManager.initialize();
      const ok = await authManager.removeAccount(accountId);
      if (ok) {
        info('Account removed.');
      } else {
        throw new CliError('Account not found.', { code: 'not_found', exitCode: ExitCode.NOT_FOUND });
      }
    }));

  const profile = auth.command('profile').description('Named sign-in settings (client ID, tenant, scopes, org mode, account) applied with --profile');

  profile
    .command('list')
    .description('List saved profiles')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const profiles = listProfiles().map(name => ({ name, ...readProfile(name) }));
      console.log(formatOutput(profiles, opts.output as OutputFormat));
    }));

  profile
    .command('show <name>')
    .description('Show a saved profile')
    .addOption(outputOption())
    .action(handleErrors(async (name: string, opts) => {
      console.log(formatOutput(readProfile(name), opts.output as OutputFormat));
    }));

  profile
    .command('set <name>')
    .description('Create a profile, or update the given fields of an existing one')
    .option('--client-id <id>', 'Application (client) ID')
    .option('--tenant-id <id>', 'Tenant ID or domain')
    .option('--scopes <list>', 'Comma-separated scopes requested at login')
    .option('--org-mode', 'Enable organization mode')
    .option('--no-org-mode', 'Disable organization mode')
    .option('--account <upn|homeAccountId>', 'Cached account to use')
    .action(handleErrors(async (name: string, _opts, cmd: Command) => {
      validateProfileName(name);
      // --account and --org-mode are also global flags, which commander may claim first.
      const opts = cmd.optsWithGlobals();
      const next: Profile = profileExists(name) ? readProfile(name) : {};
      if (opts.clientId !== undefined) next.clientId = opts.clientId;
      if (opts.tenantId !== undefined) next.tenantId = opts.tenantId;
      if (opts.scopes !== undefined) next.scopes = String(opts.scopes).split(',').map(s => s.trim()).filter(Boolean);
      if (opts.orgMode !== undefined) next.orgMode = opts.orgMode;
      if (opts.account !== undefined) next.account = opts.account;
      if (isDryRun()) return dryRunTrace(`would save profile ${name}: ${JSON.stringify(next)}`);
      saveProfile(name, next);
      info(`Profile "${name}" saved.`);
    }));

  profile
    .command('delete <name>')
    .description('Delete a saved profile')
    .action(handleErrors(async (name: string) => {
      if (!profileExists(name)) {
        throw new CliError(`Profile "${name}" not found.`, { code: 'not_found', exitCode: ExitCode.NOT_FOUND, fix: 'ms365 auth profile list' });
      }
      // Local-only: confirm() does not ask under --dry-run, so this guard is what keeps the file (DR-22).
      if (isDryRun()) return dryRunTrace(`would delete profile ${name}`);
      if (!await confirm(`Delete profile ${name}?`)) return;
      deleteProfile(name);
      info(`Profile "${name}" deleted.`);
    }));
}
