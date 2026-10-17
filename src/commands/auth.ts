import { Argument, Command, Option } from 'commander';
import { exec } from 'child_process';
import { childEnv } from '../child-env.js';
import { externalTokenMode, externalTokenModeError } from '../external-token.js';
import { adminProfileUsageError, AuthManager, explainLoginError, planStorageSwitch, storageStatus, switchStorage, tokenInfo, type TokenInfo } from '../auth.js';
import type { StorageSetting } from '../keychain.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { confirm, dryRunTrace, info, isDryRun, outputOption } from '../utils.js';
import { deleteProfile, ignoredAdminFields, listProfiles, profileExists, profileType, readProfile, saveProfile, validateProfile, validateProfileName, type Profile } from '../profiles.js';
import { profileCeiling } from '../ceiling.js';
import { openBrowserCommand } from './setup.js';
import { GraphClient } from '../graph-client.js';
import { adminConsentUrl, azGraphToken, CONSENT_REDIRECT_URI, grantScopes } from '../grant-scopes.js';

/** Open the system browser at the sign-in URL; print it too, for machines where no browser opens. */
async function openSystemBrowser(url: string): Promise<void> {
  info(`Opening the browser to sign in. If it does not open, visit:\n${url}`);
  exec(openBrowserCommand(url), { env: childEnv() }, () => { /* the URL was printed above */ });
}

/**
 * @param adminProfile the active profile's name when it is an admin profile (ID-8). Checked before --dry-run and
 * before any AuthManager is built, so `auth login` etc. under an admin profile always exit 2.
 */
export function registerAuthCommands(program: Command, getAuth: () => AuthManager, adminProfile: () => string | undefined = () => undefined): void {
  const auth = program.command('auth').description('Authentication management');
  const rejectAdmin = () => {
    const name = adminProfile();
    if (name !== undefined) throw adminProfileUsageError(name);
  };
  // AH-2: a supplied token (MS365_CLI_TOKEN or a host profile) leaves no ms365 sign-in or cache to manage: exit 2,
  // checked before --dry-run like rejectAdmin.
  const rejectExternal = (subcommand: string) => {
    const name = program.opts().profile || process.env.MS365_CLI_PROFILE;
    if (externalTokenMode(name ? readProfile(name) : {})) throw externalTokenModeError(subcommand);
  };

  auth
    .command('login')
    .description('Login (device code flow by default)')
    .option('--browser', 'Sign in with the system browser instead (needs an http://localhost redirect URI on the app registration)')
    .action(handleErrors(async (opts) => {
      rejectExternal('login');
      rejectAdmin();
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
      rejectExternal('logout');
      rejectAdmin();
      if (isDryRun()) return dryRunTrace('would log out and clear the token cache and selected account');
      const authManager = getAuth();
      await authManager.initialize();
      await authManager.logout();
      info('Logged out successfully.');
    }));

  auth
    .command('status')
    .description('Verify current login status (-o json: token source, account, expiry, scopes and profile ceiling)')
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts) => {
      // Offline under --dry-run: no AuthManager, so no token-cache read or refresh (DR-22 class).
      if (isDryRun()) return dryRunTrace('would check the login status (token cache not read)');
      const authManager = getAuth();
      await authManager.initialize();
      const source = authManager.tokenSource();
      const admin = adminProfile();
      if (opts.output !== 'text') {
        // AH-8(1): what a host needs to decide whether the token is usable, from the token itself (no Graph call).
        // A missing or unusable token still exits 3 (token_*, not_logged_in, az_required).
        const t = tokenInfo(await authManager.getToken());
        const profileName: string | undefined = program.opts().profile || process.env.MS365_CLI_PROFILE || undefined;
        const profile = profileName ? readProfile(profileName) : undefined;
        const ceiling = profile ? profileCeiling(profile) : null;
        console.log(formatOutput({
          source, profile: profileName ?? null, type: profile ? profileType(profile) : null,
          tenantId: t.tenantId ?? null, accountId: t.accountId ?? null, upn: t.upn ?? null, expiresAt: t.expiresAt ?? null,
          scopes: t.scopes, ceiling: ceiling ? [...ceiling] : null, ok: t.expired !== true,
        }, opts.output as OutputFormat));
        return;
      }
      if (source === 'external') {
        // AH-2: the supplied token is checked (exit 3 token_*) and described; nothing is read from a cache.
        const token = await authManager.getToken();
        let t: TokenInfo | undefined;
        try { t = tokenInfo(token); } catch { /* opaque: nothing to report */ }
        console.log(`Token supplied by the host (MS365_CLI_TOKEN): ${t?.upn ?? 'unknown user'}${t?.tenantId ? ` (tenant ${t.tenantId})` : ''}${t?.expiresAt ? `, expires ${t.expiresAt}` : ''}`);
        if (t) console.log(`Scopes: ${t.scopes.join(' ') || '(none)'}`);
        return;
      }
      if (admin !== undefined) {
        // ID-8: report the borrowed identity and its scp from the token itself (az errors exit 3 az_required).
        const t = tokenInfo(await authManager.getToken());
        console.log(`Admin profile "${admin}": ${t.upn ?? 'unknown user'} (tenant ${t.tenantId ?? authManager.getTenantId()}) via the Azure CLI`);
        console.log(`Scopes: ${t.scopes.join(' ') || '(none)'}`);
        return;
      }
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
      const t = tokenInfo(await authManager.getToken());
      // Admin profiles: say where the token came from (ID-8); user output is unchanged.
      const admin = adminProfile();
      const out = admin !== undefined ? { ...t, source: 'azure-cli', profile: admin } : t;
      console.log(formatOutput(out, opts.output as OutputFormat));
    }));

  auth
    .command('refresh')
    .description('Fetch a fresh access token now (e.g. after an admin grants new permissions) and show which scopes changed')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      rejectExternal('refresh');
      if (isDryRun()) return dryRunTrace('would refresh the access token (token cache not read)');
      const authManager = getAuth();
      await authManager.initialize();
      let before: string[] = [];
      try { before = tokenInfo(await authManager.getToken()).scopes; } catch { /* no usable cached token */ }
      const after = tokenInfo(await authManager.getToken(true));
      const added = after.scopes.filter(s => !before.includes(s));
      const removed = before.filter(s => !after.scopes.includes(s));
      console.log(formatOutput({
        upn: after.upn, expiresAt: after.expiresAt, scopes: after.scopes.length, added, removed,
      }, opts.output as OutputFormat));
    }));

  auth
    .command('grant-scopes')
    .description('Add the scopes ms365 requests to its app registration and grant tenant-wide admin consent (uses the Azure CLI\'s admin sign-in; falls back to an admin-consent link)')
    .option('--link', 'Only print the one-click admin-consent link (no Azure CLI needed)')
    .addOption(outputOption())
    .addHelpText('after', `
Needs an admin: either the Azure CLI signed in as one (az login --tenant <tenant>), or an admin
opening the printed link. Add --org-mode to include the org scopes (Teams, SharePoint, meetings…).
Nothing is ever removed from the app registration or the existing consent grant.`)
    .action(handleErrors(async (opts, cmd) => {
      rejectExternal('grant-scopes');
      const g = cmd.optsWithGlobals();
      const authManager = getAuth();
      const appId = authManager.getClientId();
      const scopes = authManager.requestedScopes();
      const tenant = authManager.getTenantId();
      const printLink = (why?: string) => {
        if (why) info(why);
        const linkTenant = tenant === 'common' ? 'organizations' : tenant;
        console.log(adminConsentUrl(appId, linkTenant, scopes));
        info(`An admin opens this link and clicks Accept. It needs the redirect URI ${CONSENT_REDIRECT_URI} under the app's Authentication → Mobile and desktop applications. Then run: ms365 auth refresh`);
      };
      if (opts.link) return printLink();

      let token = 'dry-run';
      let azTenant: string | undefined;
      if (!isDryRun()) {
        const az = azGraphToken(tenant);
        if (!az.ok) {
          return printLink(az.reason === 'missing'
            ? 'The Azure CLI (az) is not installed, so falling back to an admin-consent link.'
            : `The Azure CLI has no usable admin sign-in (${az.detail}). Run "az login --tenant ${tenant === 'common' ? '<tenant>' : tenant}" as an admin and retry, or use this link:`);
        }
        token = az.token.accessToken;
        azTenant = az.token.tenant;
        if (tenant !== 'common' && azTenant && azTenant !== tenant) {
          throw new CliError(`The Azure CLI is signed in to tenant ${azTenant}, but ms365 uses ${tenant}.`,
            { code: 'usage', exitCode: ExitCode.USAGE, fix: `az login --tenant ${tenant}` });
        }
      }
      const azAuth = { getToken: async () => token } as unknown as AuthManager;
      const client = new GraphClient(azAuth, isDryRun(), g.apiVersion || 'v1.0', !!g.readOnly);
      let result;
      try {
        result = await grantScopes(client, appId, scopes);
      } catch (err) {
        if (err instanceof CliError && err.exitCode === ExitCode.PERMISSION) {
          err.fix = 'sign in to the Azure CLI as an admin who can manage app registrations and grant consent (Application Administrator or Global Administrator): az login';
        }
        throw err;
      }
      if (isDryRun()) return;
      const out: Record<string, unknown> = { ...result, tenant: azTenant ?? tenant };
      // Pick up the new grant in the user's own token right away (auth refresh), if signed in.
      try {
        await authManager.initialize();
        const before = tokenInfo(await authManager.getToken()).scopes;
        const after = tokenInfo(await authManager.getToken(true)).scopes;
        out.tokenScopesAdded = after.filter(s => !before.includes(s));
      } catch {
        info('Not signed in to ms365 yet, so the token was not refreshed. Run: ms365 auth login');
      }
      console.log(formatOutput(out, opts.output as OutputFormat));
    }));

  auth
    .command('accounts')
    .description('List all cached accounts')
    .action(handleErrors(async () => {
      rejectExternal('accounts');
      rejectAdmin();
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
      rejectExternal('select');
      rejectAdmin();
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
      rejectExternal('remove');
      rejectAdmin();
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

  const storage = auth.command('storage').description('Where ms365 keeps refresh tokens: the OS keychain or the 0600 token-cache.json');
  const readOnlyMode = (cmd: Command) => !!cmd.optsWithGlobals().readOnly
    || process.env.MS365_CLI_READ_ONLY === 'true' || process.env.MS365_CLI_READ_ONLY === '1';

  storage
    .command('show')
    .description('Show the storage setting, what it resolved to, keychain availability (live probe) and the credential files')
    .option('--no-probe', 'Skip the live keychain probe (write, read back and delete a probe item; at most 1.5 s)')
    .addOption(outputOption())
    .action(handleErrors(async (opts, cmd: Command) => {
      rejectExternal('storage show');
      // The probe writes (and deletes) a keychain item, so --dry-run and --read-only skip it.
      const skip = isDryRun() ? 'not probed (--dry-run)' : readOnlyMode(cmd) ? 'not probed (--read-only)' : undefined;
      const status = storageStatus({ probe: opts.probe !== false && !skip, noProbeReason: skip });
      console.log(formatOutput(status, opts.output as OutputFormat));
    }));

  storage
    .command('use')
    .description('Switch credential storage and move every cached account to it (no re-login): auto, keychain or file')
    .addArgument(new Argument('<mode>', 'auto (file until keychain storage is verified on this platform), keychain or file').choices(['auto', 'keychain', 'file']))
    .addHelpText('after', `
Holds the token-cache lock for the whole move; the old copy is deleted only after the new one reads back
exactly. MS365_CLI_STORAGE, when set, overrides storage.json (use it for CI); switching against it is refused.`)
    .action(handleErrors(async (mode: StorageSetting, _opts, cmd: Command) => {
      rejectExternal('storage use');
      if (isDryRun()) {
        const plan = planStorageSwitch(mode, { planOnly: true });
        return dryRunTrace(plan.from === plan.to && plan.to === 'file'
          ? `would set storage to ${mode} (credentials stay in the file; ${plan.accounts} account(s))`
          : `would set storage to ${mode} and move ${plan.accounts} account(s) from the ${plan.from} to the ${plan.to}`);
      }
      if (readOnlyMode(cmd)) {
        throw new CliError('Read-only mode: switching credential storage is blocked. Remove --read-only to proceed.', { code: 'read_only', exitCode: ExitCode.PERMISSION });
      }
      const r = switchStorage(mode);
      const where = r.to === 'keychain' ? `the OS keychain (${r.backend})` : 'the 0600 file';
      const moved = r.from === r.to ? `${r.accounts.length} account(s) kept` : `${r.accounts.length} account(s) moved from the ${r.from}, no sign-in needed`;
      info(`Credential storage: ${mode} → ${where}; ${moved}.`);
      for (const hid of r.dropped) info(`Warning: account ${hid} had no stored credentials and was left out; sign in again: ms365 auth login`);
    }));

  const profile = auth.command('profile').description('Named sign-in settings (type, client ID, tenant, scopes, org mode, account) applied with --profile');

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
      const p = readProfile(name);
      for (const k of ignoredAdminFields(p)) info(`Warning: "${k}" is ignored by admin profile "${name}" (its token comes from the Azure CLI).`);
      console.log(formatOutput(p, opts.output as OutputFormat));
    }));

  profile
    .command('set <name>')
    .description('Create a profile, or update the given fields of an existing one')
    .addOption(new Option('--type <type>', 'Identity type: user (ms365 sign-in), admin (token borrowed from the Azure CLI, never stored) or host (token supplied by an agent host in MS365_CLI_TOKEN, never stored)').choices(['user', 'admin', 'host']))
    .option('--client-id <id>', 'Application (client) ID')
    .option('--tenant-id <id>', 'Tenant ID or domain')
    .option('--tenant <id>', 'Same as --tenant-id')
    .option('--scopes <list>', 'Comma-separated scopes: requested at login, and the ceiling of every command run under a user or host profile')
    .option('--org-mode', 'Enable organization mode')
    .option('--no-org-mode', 'Disable organization mode')
    .option('--account <upn|homeAccountId>', 'Cached account to use')
    .action(handleErrors(async (name: string, _opts, cmd: Command) => {
      validateProfileName(name);
      // --account and --org-mode are also global flags, which commander may claim first.
      const opts = cmd.optsWithGlobals();
      const next: Profile = profileExists(name) ? readProfile(name) : {};
      if (opts.type !== undefined) next.type = opts.type;
      if (opts.clientId !== undefined) next.clientId = opts.clientId;
      if (opts.tenantId !== undefined) next.tenantId = opts.tenantId;
      if (opts.tenant !== undefined) next.tenantId = opts.tenant;
      if (opts.scopes !== undefined) next.scopes = String(opts.scopes).split(',').map(s => s.trim()).filter(Boolean);
      if (opts.orgMode !== undefined) next.orgMode = opts.orgMode;
      if (opts.account !== undefined) next.account = opts.account;
      validateProfile(name, next, ExitCode.USAGE);
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
