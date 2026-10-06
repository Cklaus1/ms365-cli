import { Command } from 'commander';
import { AuthManager } from '../auth.js';
import { handleErrors } from '../errors.js';

export function registerAuthCommands(program: Command, getAuth: () => AuthManager): void {
  const auth = program.command('auth').description('Authentication management');

  auth
    .command('login')
    .description('Login using device code flow')
    .action(handleErrors(async () => {
      const authManager = getAuth();
      await authManager.initialize();
      await authManager.login(); // errors already say "Sign-in failed: <AADSTS detail>"
      try {
        const user = await authManager.whoAmI();
        console.log(`Logged in as ${user.displayName} (${user.email})`);
      } catch {
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
      // Throws "Not logged in" for a missing/expired login, a different message when Graph or the network fails.
      const user = await authManager.whoAmI();
      console.log(`Logged in as ${user.displayName} (${user.email})`);
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
      if (!(await authManager.selectAccount(accountId))) throw new Error('Account not found.');
      console.log('Account selected.');
    }));

  auth
    .command('remove <accountId>')
    .description('Remove a cached account')
    .action(handleErrors(async (accountId: string) => {
      const authManager = getAuth();
      await authManager.initialize();
      if (!(await authManager.removeAccount(accountId))) throw new Error('Account not found.');
      console.log('Account removed.');
    }));
}
