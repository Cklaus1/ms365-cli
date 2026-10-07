import { Command } from 'commander';
import { exec, execFile } from 'node:child_process';
import * as readline from 'node:readline';
import { handleErrors } from '../errors.js';
import { AuthManager, buildScopes, DEFAULT_CLIENT_ID, NotLoggedInError, SCOPE_GROUPS } from '../auth.js';
import { info } from '../utils.js';
import { graph } from '../command-tree.js';

/**
 * Line prompts on stderr. One reader for the whole wizard: piped answers arrive in one chunk, and a
 * per-question readline would drop the lines after the first.
 */
class Prompter {
  private readonly rl = readline.createInterface({ input: process.stdin, output: process.stderr });
  private readonly lines = this.rl[Symbol.asyncIterator]();
  private inputEnded = false;

  constructor() {
    this.rl.on('close', () => { this.inputEnded = true; });
  }

  async ask(question: string): Promise<string> {
    // Piped input may end before the last question; its buffered lines are still readable from the iterator.
    if (this.inputEnded) {
      process.stderr.write(question);
    } else {
      this.rl.setPrompt(question);
      this.rl.prompt();
    }
    const next = await this.lines.next();
    if (next.done) throw new Error('Setup aborted: input ended before the wizard finished.');
    return next.value.trim();
  }

  async yesNo(question: string, defaultYes: boolean): Promise<boolean> {
    const answer = (await this.ask(`${question} ${defaultYes ? '[Y/n]' : '[y/N]'} `)).toLowerCase();
    return answer ? answer.startsWith('y') : defaultYes;
  }

  close(): void {
    this.rl.close();
  }
}

const AZURE_PORTAL_APP_REG = 'https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade';
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function openInBrowser(url: string): void {
  // Windows `start` takes the first quoted argument as the window title, so an empty title goes first.
  const child = process.platform === 'win32'
    ? exec(`start "" "${url}"`)
    : execFile(process.platform === 'darwin' ? 'open' : 'xdg-open', [url]);
  child.on('error', () => info(`  Could not open a browser; open ${url} manually.`));
  child.unref();
}

function printScopeGroups(orgMode: boolean): void {
  for (const group of Object.values(SCOPE_GROUPS)) {
    if (group.org && !orgMode) continue;
    console.log(`    ${group.label}: ${group.scopes.join(', ')}`);
  }
}

async function runWizard(prompter: Prompter, getAuth: () => AuthManager): Promise<void> {
  console.log('');
  console.log('╔══════════════════════════════════════════════════╗');
  console.log('║         ms365-cli — Interactive Setup            ║');
  console.log('╚══════════════════════════════════════════════════╝');
  console.log('');

  console.log('Step 1: Checking existing authentication...');
  const current = getAuth();
  await current.initialize();
  try {
    const user = await current.whoAmI();
    console.log(`  Currently logged in as: ${user.displayName} (${user.email})`);
    if (await prompter.yesNo('  Use existing login?', true)) {
      console.log('  Keeping the existing login. Setup complete.');
      return;
    }
    console.log('  Will set up a new login.\n');
  } catch (err) {
    if (err instanceof NotLoggedInError) console.log('  Not currently logged in.\n');
    else console.log(`  Could not check the existing login: ${err instanceof Error ? err.message : String(err)}\n`);
  }

  console.log('Step 2: Choose authentication mode');
  console.log('');
  console.log('  [1] Quick setup — built-in app (Microsoft Graph Command Line Tools), no Azure Portal needed.');
  console.log('      Some organizations block this app or require admin consent.');
  console.log('      Personal Microsoft accounts are untested with it.');
  console.log('  [2] Custom app  — register your own app in Azure Portal (more control)');
  console.log('');
  let mode = '';
  while (mode !== '1' && mode !== '2') {
    mode = (await prompter.ask('  Choose [1/2]: ')) || '1';
    if (mode !== '1' && mode !== '2') info('  Enter 1 or 2.');
  }

  console.log('\nStep 3: Permissions');
  const orgMode = await prompter.yesNo(
    '  Enable organization features (Teams, chats, SharePoint, shared mailboxes, directory)?',
    false,
  );
  console.log('  ms365-cli will request these delegated Microsoft Graph permissions:');
  printScopeGroups(orgMode);

  let clientId = DEFAULT_CLIENT_ID;
  if (mode === '1') {
    if (current.clientId !== DEFAULT_CLIENT_ID) {
      info(`  Note: MS365_CLI_CLIENT_ID is set (${current.clientId}); ms365 commands keep using it until you unset it.`);
    }
  } else {
    console.log('\nStep 4: Register an app in Azure Portal');
    console.log('');
    console.log(`  1. Open: ${AZURE_PORTAL_APP_REG}`);
    console.log('  2. Click "New registration"');
    console.log('  3. Name: "ms365-cli" (or your preferred name)');
    console.log('  4. Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"');
    console.log('  5. Redirect URI: Select "Mobile and desktop applications"');
    console.log('     → Add: https://login.microsoftonline.com/common/oauth2/nativeclient');
    console.log('  6. Click "Register"');
    console.log('  7. API permissions → Add a permission → Microsoft Graph → Delegated: add the permissions listed above');
    console.log('     (click "Grant admin consent" if you are a tenant admin)');
    console.log('  8. Authentication → Advanced settings → "Allow public client flows" = Yes → Save');
    console.log('');
    if (await prompter.yesNo('  Open Azure Portal in browser?', true)) openInBrowser(AZURE_PORTAL_APP_REG);

    clientId = await prompter.ask('  Enter the Application (client) ID: ');
    while (!GUID.test(clientId)) {
      info('  The client ID is a GUID like 00000000-0000-0000-0000-000000000000.');
      clientId = await prompter.ask('  Enter the Application (client) ID: ');
    }
    await prompter.ask('  Press Enter when the app registration is saved...');
  }

  console.log('\n  Starting device code login...\n');
  const auth = new AuthManager(buildScopes(orgMode), { clientId });
  await auth.initialize();
  await auth.login();
  try {
    const user = await auth.whoAmI();
    console.log(`\n  Logged in as ${user.displayName} (${user.email})`);
  } catch {
    console.log('\n  Login completed. Run "ms365 auth status" to verify.');
  }

  console.log('\n╔══════════════════════════════════════════════════╗');
  console.log('║                 Setup Complete                    ║');
  console.log('╚══════════════════════════════════════════════════╝');
  const exports = [
    ...(clientId !== DEFAULT_CLIENT_ID ? [`export MS365_CLI_CLIENT_ID="${clientId}"`] : []),
    ...(orgMode ? ['export MS365_CLI_ORG_MODE=true'] : []),
  ];
  if (exports.length > 0) {
    console.log('\n  Commands request the same permissions only with these settings. Add to your shell profile:\n');
    for (const line of exports) console.log(`    ${line}`);
  }
  console.log('');
}

export function registerSetupCommand(program: Command, getAuth: () => AuthManager): void {
  const auth = program.commands.find(c => c.name() === 'auth');
  if (!auth) return;

  graph(auth.command('setup'), 'none')
    .description('Interactive setup wizard: built-in app or your own Azure AD app registration')
    .action(handleErrors(async () => {
      const prompter = new Prompter();
      try {
        await runWizard(prompter, getAuth);
      } finally {
        prompter.close();
      }
    }));
}
