import { Command } from 'commander';
import { AuthManager } from '../auth.js';

export function registerAuthCommands(program: Command, getAuth: () => AuthManager): void {
  const auth = program.command('auth').description('Authentication management');

  auth
    .command('login')
    .description('Login using device code flow')
    .action(async () => {
      const authManager = getAuth();
      await authManager.initialize();
      try {
        await authManager.login();
        const result = await authManager.verifyLogin();
        if (result.success && result.user) {
          console.log(`Logged in as ${result.user.displayName} (${result.user.email})`);
        } else {
          console.log('Login completed. Run "ms365 auth status" to verify.');
        }
      } catch (err) {
        console.error('Login failed:', (err as Error).message);
        process.exit(1);
      }
    });

  auth
    .command('logout')
    .description('Log out and clear saved credentials')
    .action(async () => {
      const authManager = getAuth();
      await authManager.initialize();
      await authManager.logout();
      console.log('Logged out successfully.');
    });

  auth
    .command('status')
    .description('Verify current login status')
    .action(async () => {
      const authManager = getAuth();
      await authManager.initialize();
      const result = await authManager.verifyLogin();
      if (result.success && result.user) {
        console.log(`Logged in as ${result.user.displayName} (${result.user.email})`);
      } else {
        console.log('Not logged in. Run: ms365 auth login');
      }
    });

  auth
    .command('accounts')
    .description('List all cached accounts')
    .action(async () => {
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
    });

  auth
    .command('select <accountId>')
    .description('Select an account by ID')
    .action(async (accountId: string) => {
      const authManager = getAuth();
      await authManager.initialize();
      const ok = await authManager.selectAccount(accountId);
      if (ok) {
        console.log('Account selected.');
      } else {
        console.error('Account not found.');
        process.exit(1);
      }
    });

  auth
    .command('remove <accountId>')
    .description('Remove a cached account')
    .action(async (accountId: string) => {
      const authManager = getAuth();
      await authManager.initialize();
      const ok = await authManager.removeAccount(accountId);
      if (ok) {
        console.log('Account removed.');
      } else {
        console.error('Account not found.');
        process.exit(1);
      }
    });
}
