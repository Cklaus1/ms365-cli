import { Command } from 'commander';
import { handleErrors } from '../errors.js';
import { AuthManager, buildScopes } from '../auth.js';
import * as readline from 'readline';

function prompt(question: string): Promise<string> {
  const rl = readline.createInterface({ input: process.stdin, output: process.stderr });
  return new Promise((resolve) => {
    rl.question(question, (answer) => {
      rl.close();
      resolve(answer.trim());
    });
  });
}

function promptYN(question: string, defaultYes = true): Promise<boolean> {
  const hint = defaultYes ? '[Y/n]' : '[y/N]';
  return prompt(`${question} ${hint} `).then((answer) => {
    if (!answer) return defaultYes;
    return answer.toLowerCase().startsWith('y');
  });
}

const AZURE_PORTAL_APP_REG = 'https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade';

const PERMISSION_GROUPS = {
  'Mail (read/write/send)': ['Mail.ReadWrite', 'Mail.Send'],
  'Calendar (read/write)': ['Calendars.ReadWrite'],
  'Files / OneDrive (read/write)': ['Files.ReadWrite'],
  'Tasks / To Do (read/write)': ['Tasks.ReadWrite'],
  'Contacts (read/write)': ['Contacts.ReadWrite'],
  'OneNote (read/create)': ['Notes.Read', 'Notes.Create'],
  'People (read)': ['People.Read'],
  'User profile (read)': ['User.Read'],
  'Teams (read) [org]': ['Team.ReadBasic.All', 'Channel.ReadBasic.All', 'TeamMember.Read.All'],
  'Chat (read/send) [org]': ['Chat.Read', 'ChatMessage.Read', 'ChatMessage.Send'],
  'SharePoint (read) [org]': ['Sites.Read.All'],
  'Shared mailbox [org]': ['Mail.Read.Shared', 'Mail.Send.Shared'],
  'Channel messages [org]': ['ChannelMessage.Read.All', 'ChannelMessage.Send'],
  'User directory [org]': ['User.Read.All'],
};

export function registerSetupCommand(program: Command, getAuth: () => AuthManager): void {
  const auth = program.commands.find(c => c.name() === 'auth');
  if (!auth) return;

  auth
    .command('setup')
    .description('Interactive setup wizard for Azure AD app registration')
    .action(handleErrors(async () => {
      console.log('');
      console.log('╔══════════════════════════════════════════════════╗');
      console.log('║         ms365-cli — Interactive Setup            ║');
      console.log('╚══════════════════════════════════════════════════╝');
      console.log('');
      console.log('This wizard helps you configure ms365-cli with your');
      console.log('own Azure AD app registration for custom permissions.');
      console.log('');

      // Step 1: Check if already logged in
      console.log('Step 1/5: Checking existing authentication...');
      const authManager = getAuth();
      await authManager.initialize();
      const currentStatus = await authManager.verifyLogin();

      if (currentStatus.success && currentStatus.user) {
        console.log(`  Currently logged in as: ${currentStatus.user.displayName} (${currentStatus.user.email})`);
        const useExisting = await promptYN('  Use existing login?');
        if (useExisting) {
          console.log('  Using existing credentials.\n');
        } else {
          console.log('  Will set up new credentials.\n');
        }
      } else {
        console.log('  Not currently logged in.\n');
      }

      // Step 2: Choose auth mode
      console.log('Step 2/5: Choose authentication mode');
      console.log('');
      console.log('  [1] Quick setup — Use built-in app registration (no Azure Portal needed)');
      console.log('  [2] Custom app  — Register your own app in Azure Portal (more control)');
      console.log('');
      const modeChoice = await prompt('  Choose [1/2]: ');

      if (modeChoice === '1' || !modeChoice) {
        // Quick setup — just login with default client ID
        console.log('\n  Using built-in app registration.');
        console.log('  Starting device code login...\n');

        await authManager.login();
        const result = await authManager.verifyLogin();
        if (result.success && result.user) {
          console.log(`\n  Logged in as ${result.user.displayName} (${result.user.email})`);
        }
        console.log('\n  Setup complete! You can now use ms365-cli.');
        return;
      }

      // Step 3: Custom app — guide through Azure Portal
      console.log('\nStep 3/5: Register an app in Azure Portal');
      console.log('');
      console.log('  Follow these steps in the Azure Portal:');
      console.log('');
      console.log(`  1. Open: ${AZURE_PORTAL_APP_REG}`);
      console.log('  2. Click "New registration"');
      console.log('  3. Name: "ms365-cli" (or your preferred name)');
      console.log('  4. Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"');
      console.log('  5. Redirect URI: Select "Mobile and desktop applications"');
      console.log('     → Add: https://login.microsoftonline.com/common/oauth2/nativeclient');
      console.log('  6. Click "Register"');
      console.log('');

      const openBrowser = await promptYN('  Open Azure Portal in browser?');
      if (openBrowser) {
        const { exec } = await import('child_process');
        const openCmd = process.platform === 'darwin' ? 'open' : process.platform === 'win32' ? 'start' : 'xdg-open';
        exec(`${openCmd} "${AZURE_PORTAL_APP_REG}"`);
      }

      console.log('');
      const clientId = await prompt('  Enter the Application (client) ID: ');
      if (!clientId) {
        console.error('  Client ID is required.');
        process.exit(1);
      }

      // Step 4: Configure API permissions
      console.log('\nStep 4/5: API Permissions');
      console.log('');
      console.log('  Add these delegated permissions in Azure Portal:');
      console.log('  → API permissions → Add a permission → Microsoft Graph → Delegated');
      console.log('');

      const selectedScopes: string[] = [];
      for (const [group, scopes] of Object.entries(PERMISSION_GROUPS)) {
        const isOrg = group.includes('[org]');
        const include = await promptYN(`  Include ${group}?`, !isOrg);
        if (include) {
          selectedScopes.push(...scopes);
          for (const scope of scopes) {
            console.log(`    → ${scope}`);
          }
        }
      }

      console.log('');
      console.log('  Make sure to click "Grant admin consent" if you are a tenant admin.');

      // Step 5: Enable public client flow
      console.log('\nStep 5/5: Enable public client flow');
      console.log('');
      console.log('  In Azure Portal:');
      console.log('  → Authentication → Advanced settings');
      console.log('  → Set "Allow public client flows" to YES');
      console.log('  → Click Save');
      console.log('');

      await prompt('  Press Enter when done...');

      // Step 6: Test login with custom app
      console.log('\n  Testing login with your app registration...\n');

      // Set env vars for this session
      process.env.MS365_CLI_CLIENT_ID = clientId;

      const customAuth = new AuthManager(selectedScopes.length > 0 ? selectedScopes : buildScopes(false));
      await customAuth.initialize();
      await customAuth.login();

      const result = await customAuth.verifyLogin();
      if (result.success && result.user) {
        console.log(`\n  Logged in as ${result.user.displayName} (${result.user.email})`);
      }

      // Show config summary
      console.log('\n╔══════════════════════════════════════════════════╗');
      console.log('║                 Setup Complete                    ║');
      console.log('╚══════════════════════════════════════════════════╝');
      console.log('');
      console.log('  To use your custom app permanently, set:');
      console.log('');
      console.log(`    export MS365_CLI_CLIENT_ID="${clientId}"`);
      console.log('');
      console.log('  Or add it to your shell profile (~/.bashrc, ~/.zshrc).');

      if (selectedScopes.some(s => s.includes('.All') || s.includes('Shared'))) {
        console.log('');
        console.log('  For organization features, also set:');
        console.log('    export MS365_CLI_ORG_MODE=true');
      }

      console.log('');
    }));
}
