import { Command, Option } from 'commander';
import { GraphClient, GraphError } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { handleErrors } from '../errors.js';
import { outputOption, done, info } from '../utils.js';
import { saveOrStream } from '../files.js';
import { listRequest, paginationOptions, type ListOptions } from '../list.js';
import { isCalendarDate } from '../graph-helpers.js';
import { detectMimeType } from '../mime.js';
import { graph } from '../command-tree.js';
import fs from 'fs';
import path from 'path';

/** Photo formats Graph documents for `PUT /me/photo/$value`. */
const PHOTO_TYPES = ['image/jpeg', 'image/png', 'image/gif'];

/** The availability/activity pairs `setUserPreferredPresence` supports (availability -> activity). */
const PREFERRED_PRESENCE: Record<string, string> = {
  Available: 'Available',
  Busy: 'Busy',
  DoNotDisturb: 'DoNotDisturb',
  BeRightBack: 'BeRightBack',
  Away: 'Away',
  Offline: 'OffWork',
};

const PHOTO_EXTENSIONS: Record<string, string> = {
  'image/jpeg': '.jpg',
  'image/pjpeg': '.jpg',
  'image/png': '.png',
  'image/gif': '.gif',
  'image/bmp': '.bmp',
  'image/webp': '.webp',
};

export function photoExtension(contentType: string): string {
  return PHOTO_EXTENSIONS[contentType.split(';')[0].trim().toLowerCase()] ?? '.bin';
}

const ISO_DATE_TIME = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?)?$/;

/**
 * An ISO 8601 schedule bound as a Graph dateTimeTimeZone in UTC. Values with `Z`/offset keep their
 * instant; values without one are local wall time (a bare date is local midnight). Impossible days
 * (`2026-02-30`) are rejected rather than rolled over.
 */
function scheduleTime(value: string, flag: string): { dateTime: string; timeZone: 'UTC'; ms: number } {
  const d = ISO_DATE_TIME.test(value) && isCalendarDate(value)
    ? new Date(value.includes('T') ? value : `${value}T00:00`)
    : new Date(NaN);
  if (Number.isNaN(d.getTime())) {
    throw new Error(`Invalid ${flag} "${value}". Use ISO 8601: 2026-10-07T09:00 (local time) or 2026-10-07T09:00:00Z.`);
  }
  return { dateTime: d.toISOString().slice(0, 19), timeZone: 'UTC', ms: d.getTime() };
}

export function registerUserCommands(program: Command, getClient: () => GraphClient): void {
  const user = program.command('user').description('User profile and management');

  graph(user.command('me'), ['GET /me'])
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

  paginationOptions(
    graph(user.command('list'), ['GET /users'])
      .description('List users in the organization (requires --org-mode)')
      .option('--filter <filter>', 'OData filter')
      .option('--select <fields>', 'Fields to return'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (opts: ListOptions & { filter?: string; select?: string }) => {
      const queryParams: Record<string, string> = {};
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.select) queryParams['$select'] = opts.select;
      await listRequest(getClient(), '/users', { queryParams }, opts);
    }));

  // ── Out of Office / Auto-replies ────────────────────────────
  graph(user.command('ooo'), ['GET /me/mailboxSettings/automaticRepliesSetting'])
    .description('Get out-of-office / automatic reply settings')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/mailboxSettings/automaticRepliesSetting');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  graph(user.command('set-ooo'), ['PATCH /me/mailboxSettings'])
    .description('Set out-of-office / automatic replies')
    .addOption(new Option('--status <status>', 'Status').choices(['alwaysEnabled', 'scheduled', 'disabled']).makeOptionMandatory())
    .option('--internal <message>', 'Auto-reply for internal senders (HTML)')
    .option('--external <message>', 'Auto-reply for external senders (HTML)')
    .addOption(new Option('--audience <audience>', 'External audience (unchanged when omitted)').choices(['all', 'contactsOnly', 'none']))
    .option('--start <datetime>', 'Schedule start, ISO 8601; local time unless it has Z/offset (required if scheduled)')
    .option('--end <datetime>', 'Schedule end, ISO 8601; local time unless it has Z/offset (required if scheduled)')
    .addOption(outputOption())
    .action(handleErrors(async (opts: {
      status: string; internal?: string; external?: string; audience?: string; start?: string; end?: string; output: string;
    }) => {
      const settings: Record<string, unknown> = { status: opts.status };
      if (opts.audience) settings.externalAudience = opts.audience;
      if (opts.internal) settings.internalReplyMessage = opts.internal;
      if (opts.external) settings.externalReplyMessage = opts.external;

      if (opts.status === 'scheduled') {
        if (!opts.start || !opts.end) throw new Error('--status scheduled requires --start and --end.');
        const { ms: startMs, ...start } = scheduleTime(opts.start, '--start');
        const { ms: endMs, ...end } = scheduleTime(opts.end, '--end');
        if (endMs <= startMs) throw new Error('--end must be after --start.');
        settings.scheduledStartDateTime = start;
        settings.scheduledEndDateTime = end;
      } else if (opts.start || opts.end) {
        throw new Error('--start/--end only apply with --status scheduled.');
      }

      const response = await getClient().request('/me/mailboxSettings', {
        method: 'PATCH',
        body: JSON.stringify({ automaticRepliesSetting: settings }),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // ── Presence ────────────────────────────────────────────────
  graph(user.command('presence'), ['GET /me/presence'])
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

  graph(user.command('set-presence'), ['POST /me/presence/setUserPreferredPresence'])
    .description('Set your preferred presence (shown while you have a Teams or app presence session)')
    .addOption(
      new Option('--availability <status>', 'Preferred availability')
        .choices(Object.keys(PREFERRED_PRESENCE))
        .makeOptionMandatory(),
    )
    .addOption(
      new Option('--activity <activity>', 'Activity; must match availability (Offline -> OffWork); default derived')
        .choices([...new Set(Object.values(PREFERRED_PRESENCE))]),
    )
    .option('--duration <duration>', 'Expiration as ISO 8601 duration (e.g. PT30M, PT8H)', 'PT60M')
    .action(handleErrors(async (opts: { availability: string; activity?: string; duration: string }) => {
      const activity = PREFERRED_PRESENCE[opts.availability];
      if (opts.activity && opts.activity !== activity) {
        const pairs = Object.entries(PREFERRED_PRESENCE).map(([a, b]) => `${a}/${b}`).join(', ');
        throw new Error(`Unsupported --availability/--activity pair ${opts.availability}/${opts.activity}. Supported: ${pairs}.`);
      }
      const res = await getClient().request('/me/presence/setUserPreferredPresence', {
        method: 'POST',
        body: JSON.stringify({ availability: opts.availability, activity, expirationDuration: opts.duration }),
      });
      done(res, `Presence set to ${opts.availability} (${activity}).`);
    }));

  graph(user.command('clear-presence'), ['POST /me/presence/clearUserPreferredPresence'])
    .description('Clear your preferred presence and revert to automatic')
    .action(handleErrors(async () => {
      const res = await getClient().request('/me/presence/clearUserPreferredPresence', { method: 'POST' });
      done(res, 'Presence cleared.');
    }));

  // ── Profile photo ───────────────────────────────────────────
  graph(user.command('photo'), ['GET /me/photo'])
    .description('Get profile photo metadata')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/photo');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  graph(user.command('download-photo'), ['GET /me/photo/$value'])
    .description('Download profile photo')
    .option('--out <path>', 'Output file path (default: profile-photo.<ext from the content type>)')
    .option('--force', 'Overwrite an existing file')
    .action(handleErrors(async (opts: { out?: string; force?: boolean }) => {
      let res;
      try {
        res = await getClient().requestBinary('/me/photo/$value');
      } catch (err) {
        if (err instanceof GraphError && err.status === 404) throw new Error('No profile photo found.');
        throw err;
      }
      if (res.status === 0) return; // dry-run: the preview is on stderr, nothing to save

      const out = opts.out ?? `profile-photo${photoExtension(res.contentType)}`;
      const saved = await saveOrStream(res.body, { out, force: opts.force });
      info(`Photo saved to ${path.resolve(out)} (${saved.bytes} bytes).`);
    }));

  graph(user.command('upload-photo <filePath>'), ['PUT /me/photo/$value'])
    .description('Upload/update profile photo (JPEG, PNG, GIF — max 4MB)')
    .action(handleErrors(async (filePath: string) => {
      const resolvedPath = path.resolve(filePath);
      const contentType = detectMimeType(resolvedPath);
      if (!PHOTO_TYPES.includes(contentType)) {
        throw new Error(`Unsupported photo type ${path.extname(resolvedPath) || '(no extension)'}: use a JPEG, PNG or GIF file.`);
      }
      if (!fs.existsSync(resolvedPath)) {
        throw new Error(`File not found: ${resolvedPath}`);
      }

      const stat = fs.statSync(resolvedPath);
      if (stat.size > 4 * 1024 * 1024) {
        throw new Error(`Photo too large (${(stat.size / 1024 / 1024).toFixed(1)} MB). Maximum is 4 MB.`);
      }

      // Upload raw binary directly — photo endpoint doesn't support upload sessions; putBinary throws on failure.
      const response = await getClient().putBinary('/me/photo/$value', resolvedPath, contentType);
      done(response, 'Profile photo updated.');
    }));

  // ── Mailbox settings ────────────────────────────────────────
  graph(user.command('mailbox-settings'), ['GET /me/mailboxSettings'])
    .description('Get mailbox settings (timezone, language, working hours)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/mailboxSettings');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  graph(user.command('set-timezone <timezone>'), ['PATCH /me/mailboxSettings'])
    .description('Set mailbox timezone (e.g. "America/New_York", "UTC")')
    .action(handleErrors(async (timezone: string) => {
      const res = await getClient().request('/me/mailboxSettings', {
        method: 'PATCH',
        body: JSON.stringify({ timeZone: timezone }),
      });
      done(res, `Timezone set to ${timezone}.`);
    }));

  graph(user.command('set-language <locale>'), ['PATCH /me/mailboxSettings'])
    .description('Set mailbox language (e.g. "en-US", "fr-FR")')
    .action(handleErrors(async (locale: string) => {
      const res = await getClient().request('/me/mailboxSettings', {
        method: 'PATCH',
        body: JSON.stringify({ language: { locale } }),
      });
      done(res, `Language set to ${locale}.`);
    }));

  graph(user.command('working-hours'), ['GET /me/mailboxSettings/workingHours'])
    .description('Get working hours configuration')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/mailboxSettings/workingHours');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));
}
