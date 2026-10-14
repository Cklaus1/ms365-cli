import { Command, Option } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat, printList } from '../formatter.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import { addListOptions, outputOption, info, success, parseTime, resolveTimeZone, safeOutputPath } from '../utils.js';
import fs from 'fs';
import path from 'path';

const usage = (msg: string) => new CliError(msg, { code: 'usage', exitCode: ExitCode.USAGE });

/** setPresence's sessionId is the app's client ID: the one GraphClient resolved, or a placeholder when only previewing (DR-15). */
function presenceSessionId(client: GraphClient): string {
  const id = client.clientId();
  if (id) return id;
  if (client.isDryRun()) return '<client-id>';
  throw new CliError('No client ID configured. Set MS365_CLI_CLIENT_ID in your .env file or run: ms365 auth setup',
    { code: 'no_client_id', exitCode: ExitCode.AUTH, fix: 'ms365 auth setup' });
}

/** An Intl-valid IANA name or a Windows-shaped one ("Pacific Standard Time", "UTC") — Graph accepts both (F24). */
function validateMailboxTimeZone(tz: string): void {
  if (/^[A-Z][A-Za-z .+()-]* Time$|^UTC$/.test(tz)) return;
  try {
    new Intl.DateTimeFormat('en-US', { timeZone: tz });
  } catch {
    throw usage(`Unknown time zone "${tz}". Use an IANA name (America/New_York) or a Windows name (Pacific Standard Time).`);
  }
}

/** ISO 8601 duration PT[nH][nM], 5 minutes to 4 hours (setPresence's documented range †). */
function parsePresenceDuration(value: string): string {
  const m = /^PT(?:(\d+)H)?(?:(\d+)M)?$/.exec(value);
  const minutes = m && (m[1] || m[2]) ? Number(m[1] ?? 0) * 60 + Number(m[2] ?? 0) : NaN;
  if (!(minutes >= 5 && minutes <= 240)) throw usage(`Invalid --duration "${value}". Use an ISO 8601 duration from PT5M to PT4H (e.g. PT30M, PT1H30M).`);
  return value;
}

const PHOTO_TYPES: Record<string, string> = { '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.gif': 'image/gif' };

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

  addListOptions(user
    .command('list')
    .description('List users in the organization (requires --org-mode)')
    .option('--filter <filter>', 'OData filter')
    .option('--select <fields>', 'Fields to return'), { description: 'Number of users' })
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.select) queryParams['$select'] = opts.select;

      await printList(client, '/users', { queryParams }, opts);
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
    .addOption(new Option('--status <status>', 'Auto-reply status').choices(['alwaysEnabled', 'scheduled', 'disabled']).makeOptionMandatory())
    .option('--internal <message>', 'Auto-reply for internal senders (HTML)')
    .option('--external <message>', 'Auto-reply for external senders (HTML)')
    .addOption(new Option('--audience <audience>', 'External audience').choices(['all', 'contactsOnly', 'none']).default('all'))
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
    .description('Set your presence status for this app session (setPresence; values unverified live)')
    .addOption(new Option('--availability <status>', 'Availability').choices(['Available', 'Busy', 'Away', 'DoNotDisturb']).makeOptionMandatory())
    .addOption(new Option('--activity <activity>', 'Activity').choices(['Available', 'InACall', 'InAConferenceCall', 'Away', 'Presenting']).makeOptionMandatory())
    .option('--duration <duration>', 'How long it lasts: ISO 8601 duration, PT5M to PT4H', 'PT60M')
    .action(handleErrors(async (opts) => {
      const duration = parsePresenceDuration(opts.duration);
      const client = getClient();
      await client.request('/me/presence/setPresence', {
        method: 'POST',
        body: JSON.stringify({
          sessionId: presenceSessionId(client),
          availability: opts.availability,
          activity: opts.activity,
          expirationDuration: duration,
        }),
      });
      success(`Presence set to ${opts.availability} (${opts.activity}).`);
    }));

  user
    .command('clear-presence')
    .description('Clear custom presence and revert to automatic')
    .action(handleErrors(async () => {
      const client = getClient();
      await client.request('/me/presence/clearPresence', {
        method: 'POST',
        body: JSON.stringify({ sessionId: presenceSessionId(client) }),
      });
      success('Presence cleared.');
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
          throw new CliError('No profile photo found.', { code: 'not_found', exitCode: ExitCode.NOT_FOUND, status: 404 });
        }
        throw err;
      }
      if (client.isPreview(response)) return; // nothing downloaded
      const data = response.data as { contentBytes?: string };
      fs.writeFileSync(outPath, Buffer.from(data.contentBytes ?? '', 'base64'));
      info(`Photo saved to ${outPath}`);
    }));

  user
    .command('upload-photo <filePath>')
    .description('Upload/update profile photo (JPEG, PNG, GIF — max 4MB)')
    .action(handleErrors(async (filePath: string) => {
      const resolvedPath = path.resolve(filePath);
      let stat: fs.Stats;
      try {
        stat = fs.statSync(resolvedPath);
      } catch {
        throw usage(`File not found: ${resolvedPath}`);
      }
      if (!stat.isFile()) throw usage(`Not a file: ${resolvedPath}`);
      const contentType = PHOTO_TYPES[path.extname(resolvedPath).toLowerCase()];
      if (!contentType) throw usage(`Unsupported photo type "${path.extname(resolvedPath)}". Use .jpg, .jpeg, .png or .gif.`);
      if (stat.size > 4 * 1024 * 1024) {
        throw usage(`Photo too large (${(stat.size / 1024 / 1024).toFixed(1)} MB). Maximum is 4 MB.`);
      }

      // Upload raw binary directly — photo endpoint doesn't support upload sessions
      await getClient().uploadFile('/me/photo/$value', resolvedPath, contentType);
      success('Profile photo updated.');
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
    .description('Set mailbox timezone: IANA ("America/New_York") or Windows ("Pacific Standard Time") name')
    .action(handleErrors(async (timezone: string) => {
      validateMailboxTimeZone(timezone);
      const client = getClient();
      await client.request('/me/mailboxSettings', {
        method: 'PATCH',
        body: JSON.stringify({ timeZone: timezone }),
      });
      success(`Timezone set to ${timezone}.`);
    }));

  user
    .command('set-language <locale>')
    .description('Set mailbox language (e.g. "en-US", "fr-FR")')
    .action(handleErrors(async (locale: string) => {
      try {
        Intl.getCanonicalLocales(locale);
      } catch {
        throw usage(`Invalid language tag "${locale}". Use a BCP 47 tag like en-US or fr-FR.`);
      }
      const client = getClient();
      await client.request('/me/mailboxSettings', {
        method: 'PATCH',
        body: JSON.stringify({ language: { locale } }),
      });
      success(`Language set to ${locale}.`);
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
