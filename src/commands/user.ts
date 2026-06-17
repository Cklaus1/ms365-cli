import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { handleErrors } from '../errors.js';
import { outputOption, info } from '../utils.js';
import fs from 'fs';
import path from 'path';

export function registerUserCommands(program: Command, getClient: () => GraphClient): void {
  const user = program.command('user').description('User profile and management');

  user
    .command('me')
    .description('Get current user profile')
    .option('--select <fields>', 'Fields to return')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.select) queryParams['$select'] = opts.select;

      const response = await client.request('/me', { queryParams });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  user
    .command('list')
    .description('List users in the organization (requires --org-mode)')
    .option('--top <n>', 'Number of users')
    .option('--filter <filter>', 'OData filter')
    .option('--select <fields>', 'Fields to return')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.select) queryParams['$select'] = opts.select;

      const response = opts.all
        ? await client.requestAllPages('/users', { queryParams })
        : await client.request('/users', { queryParams });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // ── Out of Office / Auto-replies ────────────────────────────
  user
    .command('ooo')
    .description('Get out-of-office / automatic reply settings')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/mailboxSettings/automaticRepliesSetting');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  user
    .command('set-ooo')
    .description('Set out-of-office / automatic replies')
    .requiredOption('--status <status>', 'Status: alwaysEnabled, scheduled, disabled')
    .option('--internal <message>', 'Auto-reply for internal senders (HTML)')
    .option('--external <message>', 'Auto-reply for external senders (HTML)')
    .option('--audience <audience>', 'External audience: all, contactsOnly, none', 'all')
    .option('--start <datetime>', 'Schedule start (ISO 8601, required if scheduled)')
    .option('--end <datetime>', 'Schedule end (ISO 8601, required if scheduled)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const settings: Record<string, unknown> = {
        status: opts.status,
        externalAudience: opts.audience,
      };

      if (opts.internal) {
        settings.internalReplyMessage = opts.internal;
      }
      if (opts.external) {
        settings.externalReplyMessage = opts.external;
      }
      if (opts.start && opts.end) {
        settings.scheduledStartDateTime = { dateTime: opts.start, timeZone: 'UTC' };
        settings.scheduledEndDateTime = { dateTime: opts.end, timeZone: 'UTC' };
      }

      const response = await client.request('/me/mailboxSettings', {
        method: 'PATCH',
        body: JSON.stringify({ automaticRepliesSetting: settings }),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // ── Presence ────────────────────────────────────────────────
  user
    .command('presence')
    .description('Get your current presence status (Available, Busy, etc.)')
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/presence');
      const data = response.data as Record<string, unknown>;

      if (opts.output === 'text') {
        console.log(`Availability: ${data.availability || 'Unknown'}`);
        console.log(`Activity:     ${data.activity || 'Unknown'}`);
      } else {
        console.log(formatOutput(response.data, opts.output as OutputFormat));
      }
    }));

  user
    .command('set-presence')
    .description('Set your presence status')
    .requiredOption('--availability <status>', 'Available, Busy, DoNotDisturb, BeRightBack, Away, Offline')
    .requiredOption('--activity <activity>', 'Available, Busy, DoNotDisturb, BeRightBack, Away, OffWork, etc.')
    .option('--duration <minutes>', 'Duration in minutes (ISO 8601 format, e.g. PT30M)', 'PT60M')
    .action(handleErrors(async (opts) => {
      const client = getClient();
      // setPresence requires application ID — use the configured client ID
      const clientId = process.env.MS365_CLI_CLIENT_ID || process.env.MS365_MCP_CLIENT_ID;
      if (!clientId) throw new Error('Client ID not configured. Set MS365_CLI_CLIENT_ID or run: ms365 auth setup');
      await client.request('/me/presence/setPresence', {
        method: 'POST',
        body: JSON.stringify({
          sessionId: clientId,
          availability: opts.availability,
          activity: opts.activity,
          expirationDuration: opts.duration,
        }),
      });
      console.log(`Presence set to ${opts.availability} (${opts.activity}).`);
    }));

  user
    .command('clear-presence')
    .description('Clear custom presence and revert to automatic')
    .action(handleErrors(async () => {
      const client = getClient();
      const clientId = process.env.MS365_CLI_CLIENT_ID || process.env.MS365_MCP_CLIENT_ID;
      if (!clientId) throw new Error('Client ID not configured. Set MS365_CLI_CLIENT_ID or run: ms365 auth setup');
      await client.request('/me/presence/clearPresence', {
        method: 'POST',
        body: JSON.stringify({ sessionId: clientId }),
      });
      info('Presence cleared.');
    }));

  // ── Profile photo ───────────────────────────────────────────
  user
    .command('photo')
    .description('Get profile photo metadata')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/photo');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  user
    .command('download-photo')
    .description('Download profile photo')
    .option('--out <path>', 'Output file path', 'profile-photo.jpg')
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/photo/$value');
      const data = response.data as Record<string, unknown>;

      if (data.contentBytes) {
        const outPath = path.resolve(opts.out);
        fs.writeFileSync(outPath, Buffer.from(data.contentBytes as string, 'base64'));
        console.log(`Photo saved to ${outPath}`);
      } else {
        console.log('No profile photo found.');
      }
    }));

  user
    .command('upload-photo <filePath>')
    .description('Upload/update profile photo (JPEG, PNG, GIF — max 4MB)')
    .action(handleErrors(async (filePath: string) => {
      const resolvedPath = path.resolve(filePath);
      if (!fs.existsSync(resolvedPath)) {
        throw new Error(`File not found: ${resolvedPath}`);
      }

      const client = getClient();
      const ext = path.extname(resolvedPath).toLowerCase();
      const mimeMap: Record<string, string> = { '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.gif': 'image/gif' };
      const contentType = mimeMap[ext] || 'image/jpeg';

      const stat = fs.statSync(resolvedPath);
      if (stat.size > 4 * 1024 * 1024) {
        throw new Error(`Photo too large (${(stat.size / 1024 / 1024).toFixed(1)} MB). Maximum is 4 MB.`);
      }

      // Upload raw binary directly — photo endpoint doesn't support upload sessions
      const response = await client.uploadFile('/me/photo/$value', resolvedPath, contentType);
      if (response.status >= 400) {
        throw new Error(`Photo upload failed (${response.status})`);
      }
      info('Profile photo updated.');
    }));

  // ── Mailbox settings ────────────────────────────────────────
  user
    .command('mailbox-settings')
    .description('Get mailbox settings (timezone, language, working hours)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/mailboxSettings');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  user
    .command('set-timezone <timezone>')
    .description('Set mailbox timezone (e.g. "America/New_York", "UTC")')
    .action(handleErrors(async (timezone: string) => {
      const client = getClient();
      await client.request('/me/mailboxSettings', {
        method: 'PATCH',
        body: JSON.stringify({ timeZone: timezone }),
      });
      console.log(`Timezone set to ${timezone}.`);
    }));

  user
    .command('set-language <locale>')
    .description('Set mailbox language (e.g. "en-US", "fr-FR")')
    .action(handleErrors(async (locale: string) => {
      const client = getClient();
      await client.request('/me/mailboxSettings', {
        method: 'PATCH',
        body: JSON.stringify({ language: { locale } }),
      });
      console.log(`Language set to ${locale}.`);
    }));

  user
    .command('working-hours')
    .description('Get working hours configuration')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/mailboxSettings/workingHours');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));
}
