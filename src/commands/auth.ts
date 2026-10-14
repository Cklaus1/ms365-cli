import { Command } from 'commander';
import { AuthManager } from '../auth.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';

export function registerAuthCommands(program: Command, getAuth: () => AuthManager): void {
  const auth = program.command('auth').description('Authentication management');

  auth
    .command('login')
    .description('Login using device code flow')
    .action(handleErrors(async () => {
      const authManager = getAuth();
      await authManager.initialize();
      try {
        await authManager.login();
      } catch (err) {
        throw new CliError(`Login failed: ${(err as Error).message}`, { code: 'login_failed', exitCode: ExitCode.AUTH });
      }
      const result = await authManager.verifyLogin();
      if (result.success && result.user) {
        console.log(`Logged in as ${result.user.displayName} (${result.user.email})`);
      } else {
        console.log('Login completed. Run "ms365 auth status" to verify.');
      }
    }));

  auth
    .command('logout')
    .description('Log out and clear saved credentials')
    .action(handleErrors(async () => {
      const authManager = getAuth();
      await authManager.initialize();
      await authManager.logout();
      console.log('Logged out successfully.');
    }));

  auth
    .command('status')
    .description('Verify current login status')
    .action(handleErrors(async () => {
      const authManager = getAuth();
      await authManager.initialize();
      const result = await authManager.verifyLogin();
      if (result.success && result.user) {
        console.log(`Logged in as ${result.user.displayName} (${result.user.email})`);
      } else {
        console.log('Not logged in. Run: ms365 auth login');
      }
    }));

  auth
    .command('accounts')
    .description('List all cached accounts')
    .action(handleErrors(async () => {
      const authManager = getAuth();
      await authManager.initialize();
      const accounts = await authManager.listAccounts();
      const selectedId = authManager.getSelectedAccountId();

      if (accounts.length === 0) {
        console.log('No accounts found. Run: ms365 auth login');
        return;
      }

      for (const acc of accounts) {
        const marker = acc.homeAccountId === selectedId ? ' (active)' : '';
        console.log(`  ${acc.username}${marker}`);
        console.log(`    ID: ${acc.homeAccountId}`);
      }
    }));

  auth
    .command('select <accountId>')
    .description('Select an account by ID')
    .action(handleErrors(async (accountId: string) => {
      const authManager = getAuth();
      await authManager.initialize();
      const ok = await authManager.selectAccount(accountId);
      if (ok) {
        console.log('Account selected.');
      } else {
        throw new CliError('Account not found.', { code: 'not_found', exitCode: ExitCode.NOT_FOUND });
      }
    }));

  auth
    .command('remove <accountId>')
    .description('Remove a cached account')
    .action(handleErrors(async (accountId: string) => {
      const authManager = getAuth();
      await authManager.initialize();
      const ok = await authManager.removeAccount(accountId);
      if (ok) {
        console.log('Account removed.');
      } else {
        throw new CliError('Account not found.', { code: 'not_found', exitCode: ExitCode.NOT_FOUND });
      }
    }));
}
