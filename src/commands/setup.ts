import { Command } from 'commander';
import fs from 'fs';
import readline from 'readline';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import { AuthManager, buildScopes, configFilePath, writeFileAtomic } from '../auth.js';
import { info } from '../utils.js';

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
  'Presence (read/write)': ['Presence.Read', 'Presence.ReadWrite'],
  'User profile (read)': ['User.Read'],
  'Teams (read) [org]': ['Team.ReadBasic.All', 'Channel.ReadBasic.All', 'TeamMember.Read.All'],
  'Chat (read/send) [org]': ['Chat.Read', 'ChatMessage.Read', 'ChatMessage.Send'],
  'SharePoint (read) [org]': ['Sites.Read.All'],
  'Shared mailbox [org]': ['Mail.Read.Shared', 'Mail.Send.Shared'],
  'Channel messages [org]': ['ChannelMessage.Read.All', 'ChannelMessage.Send'],
  'User directory [org]': ['User.Read.All'],
  'Groups (read) [org]': ['Group.Read.All'],
  'Shared calendars (read) [org]': ['Calendars.Read.Shared'],
  'Rooms / places (read) [org]': ['Place.Read.All'],
};

export function openBrowserCommand(url: string, platform: NodeJS.Platform = process.platform): string {
  // On Windows `start` takes the first quoted argument as the window title.
  if (platform === 'win32') return `start "" "${url}"`;
  return `${platform === 'darwin' ? 'open' : 'xdg-open'} "${url}"`;
}

/** Save the wizard's scopes to config.json (D14), keeping any other keys; buildScopes() reads them back. */
function saveScopes(scopes: string[]): void {
  let config: Record<string, unknown> = {};
  try {
    config = JSON.parse(fs.readFileSync(configFilePath(), 'utf8'));
  } catch { /* no config yet */ }
  writeFileAtomic(configFilePath(), JSON.stringify({ ...config, scopes }, null, 2));
}

export function registerSetupCommand(program: Command, getAuth: () => AuthManager): void {
  const auth = program.commands.find(c => c.name() === 'auth');
  if (!auth) return;

  auth
    .command('setup')
    .description('Interactive setup wizard for Azure AD app registration')
    .action(handleErrors(async () => {
      info('');
      info('╔══════════════════════════════════════════════════╗');
      info('║         ms365-cli — Interactive Setup            ║');
      info('╚══════════════════════════════════════════════════╝');
      info('');
      info('This wizard helps you configure ms365-cli with your');
      info('own Azure AD app registration for custom permissions.');
      info('');

      // Step 1: Check if already logged in. AuthManager needs a client ID, so only check when one is configured.
      info('Step 1/5: Checking existing authentication...');
      if (process.env.MS365_CLI_CLIENT_ID || process.env.MS365_MCP_CLIENT_ID) {
        const authManager = getAuth();
        await authManager.initialize();
        const currentStatus = await authManager.verifyLogin();

        if (currentStatus.success && currentStatus.user) {
          info(`  Currently logged in as: ${currentStatus.user.displayName} (${currentStatus.user.email})`);
          const useExisting = await promptYN('  Use existing login?');
          if (useExisting) {
            info('  Using existing credentials.');
            return;
          }
          info('  Will set up new credentials.\n');
        } else {
          info('  Not currently logged in.\n');
        }
      } else {
        info('  No client ID configured yet.\n');
      }

      // Step 2: Choose auth mode
      info('Step 2/5: Choose authentication mode');
      info('');
      info('  [1] Quick setup — Use built-in app registration (not available yet)');
      info('  [2] Custom app  — Register your own app in Azure Portal (more control)');
      info('');
      const modeChoice = await prompt('  Choose [1/2]: ');

      if (modeChoice === '1') {
        // No published multi-tenant app ID exists yet (D13).
        info('\n  The built-in app registration is not available yet.');
        info('  Continuing with option [2]: register your own app.');
      }

      // Step 3: Custom app — guide through Azure Portal
      info('\nStep 3/5: Register an app in Azure Portal');
      info('');
      info('  Follow these steps in the Azure Portal:');
      info('');
      info(`  1. Open: ${AZURE_PORTAL_APP_REG}`);
      info('  2. Click "New registration"');
      info('  3. Name: "ms365-cli" (or your preferred name)');
      info('  4. Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"');
      info('  5. Redirect URI: Select "Mobile and desktop applications"');
      info('     → Add: https://login.microsoftonline.com/common/oauth2/nativeclient');
      info('  6. Click "Register"');
      info('');

      const openBrowser = await promptYN('  Open Azure Portal in browser?');
      if (openBrowser) {
        const { exec } = await import('child_process');
        exec(openBrowserCommand(AZURE_PORTAL_APP_REG));
      }

      info('');
      const clientId = await prompt('  Enter the Application (client) ID: ');
      if (!clientId) {
        throw new CliError('Client ID is required.', { code: 'usage', exitCode: ExitCode.USAGE, fix: 'ms365 auth setup' });
      }

      // Step 4: Configure API permissions
      info('\nStep 4/5: API Permissions');
      info('');
      info('  Add these delegated permissions in Azure Portal:');
      info('  → API permissions → Add a permission → Microsoft Graph → Delegated');
      info('');

      const selectedScopes: string[] = [];
      for (const [group, scopes] of Object.entries(PERMISSION_GROUPS)) {
        const isOrg = group.includes('[org]');
        const include = await promptYN(`  Include ${group}?`, !isOrg);
        if (include) {
          selectedScopes.push(...scopes);
          for (const scope of scopes) {
            info(`    → ${scope}`);
          }
        }
      }

      info('');
      info('  Make sure to click "Grant admin consent" if you are a tenant admin.');

      // Step 5: Enable public client flow
      info('\nStep 5/5: Enable public client flow');
      info('');
      info('  In Azure Portal:');
      info('  → Authentication → Advanced settings');
      info('  → Set "Allow public client flows" to YES');
      info('  → Click Save');
      info('');

      await prompt('  Press Enter when done...');

      // Step 6: Test login with custom app
      info('\n  Testing login with your app registration...\n');

      if (selectedScopes.length > 0) {
        saveScopes(selectedScopes);
        info(`  Saved the chosen permissions to ${configFilePath()}\n`);
      }
      const customAuth = new AuthManager(buildScopes(false), { clientId });
      await customAuth.initialize();
      await customAuth.login();

      const result = await customAuth.verifyLogin();
      if (result.success && result.user) {
        info(`\n  Logged in as ${result.user.displayName} (${result.user.email})`);
      }

      // Show config summary
      info('\n╔══════════════════════════════════════════════════╗');
      info('║                 Setup Complete                    ║');
      info('╚══════════════════════════════════════════════════╝');
      info('');
      info('  To use your custom app permanently, set:');
      info('');
      info(`    export MS365_CLI_CLIENT_ID="${clientId}"`);
      info('');
      info('  Or add it to your shell profile (~/.bashrc, ~/.zshrc).');

      if (selectedScopes.some(s => s.includes('.All') || s.includes('Shared'))) {
        info('');
        info('  For organization features, also set:');
        info('    export MS365_CLI_ORG_MODE=true');
      }

      info('');
    }));
}
