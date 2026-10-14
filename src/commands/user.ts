import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import { outputOption, info, parseTime, resolveTimeZone, safeOutputPath } from '../utils.js';
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
        ? await client.requestAllPages('/users', { queryParams }, opts.pageLimit ? parseInt(opts.pageLimit) : undefined, opts.pageDelay ? parseInt(opts.pageDelay) : undefined)
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
    .option('--start <datetime>', 'Schedule start (ISO 8601 or e.g. "tomorrow 9:00"; required if scheduled)')
    .option('--end <datetime>', 'Schedule end (ISO 8601 or e.g. "+3d"; required if scheduled)')
    .option('--timezone <tz>', 'IANA time zone for --start/--end (default: MS365_CLI_TIMEZONE or system zone)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      if (opts.status === 'scheduled' && !(opts.start && opts.end)) {
        throw new CliError('--status scheduled requires both --start and --end.', { code: 'usage', exitCode: ExitCode.USAGE });
      }
      const tz = resolveTimeZone(opts.timezone);
      const start = opts.start ? parseTime(opts.start, { tz }) : undefined;
      const end = opts.end ? parseTime(opts.end, { tz }) : undefined;

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
      if (start && end) {
        settings.scheduledStartDateTime = { dateTime: start.dateTime, timeZone: start.timeZone };
        settings.scheduledEndDateTime = { dateTime: end.dateTime, timeZone: end.timeZone };
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
    .option('--force', 'Overwrite --out if it already exists')
    .action(handleErrors(async (opts) => {
      const outPath = safeOutputPath(opts.out, { out: opts.out, force: opts.force });
      const client = getClient();
      let response;
      try {
        response = await client.request('/me/photo/$value', { responseType: 'binary' });
      } catch (err) {
        if (err instanceof CliError && err.status === 404) {
          info('No profile photo found.');
          process.exitCode = ExitCode.NOT_FOUND;
          return;
        }
        throw err;
      }
      const data = response.data as { contentBytes?: string };
      if (!data.contentBytes) return; // dry-run: nothing downloaded
      fs.writeFileSync(outPath, Buffer.from(data.contentBytes, 'base64'));
      info(`Photo saved to ${outPath}`);
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
