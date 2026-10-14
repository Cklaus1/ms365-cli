import { Command, Option } from 'commander';
import fs from 'fs';
import { once } from 'events';
import { Readable } from 'stream';
import { pipeline } from 'stream/promises';
import type { ReadableStream as WebReadableStream } from 'stream/web';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat, printList } from '../formatter.js';
import { addListOptions, outputOption, info, odataQuote, safeOutputPath, resolveTimeZone, preferTimeZone, parseTime, formatSize } from '../utils.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';

const TZ_HELP = 'IANA time zone (default: MS365_CLI_TIMEZONE or the local zone)';
const TIME_HELP = 'ISO 8601 or natural ("tomorrow 15:00", "next mon 9:30", "+2h")';
const MEETING_ARG = 'online meeting id or Teams join URL';

const isUrl = (s: string) => /^https?:\/\//i.test(s);
const isDryRun = (data: unknown) => !!(data as { dryRun?: boolean } | null)?.dryRun;

/** Look up an online meeting by its Teams join URL (the link calendar events carry). */
async function meetingByJoinUrl(client: GraphClient, joinUrl: string): Promise<Record<string, unknown>> {
  const response = await client.request('/me/onlineMeetings', {
    queryParams: { $filter: `JoinWebUrl eq '${odataQuote(joinUrl)}'` },
  });
  if (isDryRun(response.data)) return response.data as Record<string, unknown>;
  const meeting = (response.data as { value?: Array<Record<string, unknown>> }).value?.[0];
  if (!meeting) {
    throw new CliError('No online meeting found for that join URL', {
      code: 'not_found', exitCode: ExitCode.NOT_FOUND,
      fix: 'only meetings you organize or attend can be resolved; check the URL or pass the meeting id',
    });
  }
  return meeting;
}

/** The meetings path segment for a meeting id or join URL. */
async function meetingPath(client: GraphClient, idOrUrl: string): Promise<string> {
  if (!isUrl(idOrUrl)) return `/me/onlineMeetings/${encodeURIComponent(idOrUrl)}`;
  const meeting = await meetingByJoinUrl(client, idOrUrl);
  const id = typeof meeting.id === 'string' ? meeting.id : 'MEETING_ID';
  return `/me/onlineMeetings/${encodeURIComponent(id)}`;
}

/** The given child id, or the most recently created one from `collection` (transcripts/recordings). */
async function latestId(client: GraphClient, collection: string, given: string | undefined, kind: string): Promise<string> {
  if (given) return encodeURIComponent(given);
  const response = await client.request(collection);
  if (isDryRun(response.data)) return 'LATEST';
  const items = (response.data as { value?: Array<{ id: string; createdDateTime?: string }> }).value ?? [];
  const latest = items.reduce<{ id: string; createdDateTime?: string } | undefined>(
    (best, it) => (!best || (it.createdDateTime ?? '') > (best.createdDateTime ?? '') ? it : best), undefined);
  if (!latest) {
    throw new CliError(`This meeting has no ${kind}s`, { code: 'not_found', exitCode: ExitCode.NOT_FOUND });
  }
  return encodeURIComponent(latest.id);
}

/**
 * Convert a WebVTT transcript to plain "Speaker: text" lines: timestamps, cue ids,
 * NOTE blocks and markup are dropped, and consecutive cues by one speaker are merged.
 */
export function vttToText(vtt: string): string {
  const lines: Array<{ speaker: string; text: string }> = [];
  for (const block of vtt.replace(/\r\n?/g, '\n').split(/\n{2,}/)) {
    const rows = block.split('\n').filter(r => r.trim() !== '');
    const timing = rows.findIndex(r => r.includes('-->'));
    if (timing < 0) continue; // WEBVTT header, NOTE, STYLE
    const raw = rows.slice(timing + 1).join(' ');
    const voice = raw.match(/<v(?:\.[^\s>]+)*\s+([^>]+)>/);
    const speaker = voice ? voice[1].trim() : '';
    const text = raw.replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
    if (!text) continue;
    const prev = lines[lines.length - 1];
    if (prev && prev.speaker === speaker) prev.text += ` ${text}`;
    else lines.push({ speaker, text });
  }
  return lines.map(l => (l.speaker ? `${l.speaker}: ${l.text}` : l.text)).join('\n') + (lines.length ? '\n' : '');
}

async function streamTo(stream: unknown, target: string | null): Promise<void> {
  const body = Readable.fromWeb(stream as WebReadableStream<Uint8Array>);
  if (target === null) {
    for await (const chunk of body) {
      if (!process.stdout.write(chunk)) await once(process.stdout, 'drain');
    }
    return;
  }
  try {
    await pipeline(body, fs.createWriteStream(target));
  } catch (err) {
    fs.rmSync(target, { force: true });
    throw err;
  }
  info(`Downloaded to ${target} (${formatSize(fs.statSync(target).size)})`);
}

export function registerMeetingsCommands(program: Command, getClient: () => GraphClient): void {
  const meetings = program.command('meetings').alias('meet').description('Teams online meetings, transcripts and recordings');

  addListOptions(meetings
    .command('list')
    .description('List your online meetings in a time window (calendar events with a Teams link)')
    .option('--start <datetime>', `Window start (default: now): ${TIME_HELP}`)
    .option('--end <datetime>', `Window end (default: 7 days after the start): ${TIME_HELP}`)
    .option('--timezone <tz>', `Zone for the window and returned times; ${TZ_HELP}`))
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const tz = resolveTimeZone(opts.timezone);
      const start = parseTime(opts.start ?? 'now', { tz });
      const end = opts.end ? parseTime(opts.end, { tz }) : parseTime('+7d', { tz, now: start.date });
      // /me/onlineMeetings can't be listed; the calendar is the index of meetings, filtered server-side.
      await printList(getClient(), '/me/calendarView', {
        queryParams: {
          startDateTime: start.date.toISOString(),
          endDateTime: end.date.toISOString(),
          $filter: 'isOnlineMeeting eq true',
          $select: 'subject,start,end,organizer,onlineMeeting,onlineMeetingProvider',
          $orderby: 'start/dateTime',
        },
        headers: preferTimeZone(tz),
      }, opts);
    }));

  meetings
    .command('get <joinUrlOrId>')
    .description('Get an online meeting by Teams join URL or id')
    .addOption(outputOption())
    .action(handleErrors(async (joinUrlOrId: string, opts) => {
      const client = getClient();
      const data = isUrl(joinUrlOrId)
        ? await meetingByJoinUrl(client, joinUrlOrId)
        : (await client.request(`/me/onlineMeetings/${encodeURIComponent(joinUrlOrId)}`)).data;
      console.log(formatOutput(data, opts.output as OutputFormat));
    }));

  meetings
    .command('transcripts <meeting>')
    .description(`List a meeting's transcripts (<meeting>: ${MEETING_ARG})`)
    .addOption(outputOption())
    .action(handleErrors(async (meeting: string, opts) => {
      const client = getClient();
      const response = await client.request(`${await meetingPath(client, meeting)}/transcripts`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  meetings
    .command('transcript <meeting> [transcriptId]')
    .description(`Download a transcript (default: the latest) as VTT or "Speaker: text" (<meeting>: ${MEETING_ARG})`)
    .addOption(new Option('--format <fmt>', 'vtt or txt').choices(['vtt', 'txt']).default('vtt'))
    .option('--out <path>', 'Output file (default: stdout)')
    .option('--force', 'Overwrite an existing file')
    .action(handleErrors(async (meeting: string, transcriptId: string | undefined, opts) => {
      const target = opts.out ? safeOutputPath(opts.out, { out: opts.out, force: opts.force }) : null;
      const client = getClient();
      const base = `${await meetingPath(client, meeting)}/transcripts`;
      const tid = await latestId(client, base, transcriptId, 'transcript');
      const response = await client.request(`${base}/${tid}/content`, {
        queryParams: { $format: 'text/vtt' },
        responseType: 'text',
      });
      if (typeof response.data !== 'string') {
        console.log(formatOutput(response.data, 'json')); // --dry-run
        return;
      }
      const text = opts.format === 'txt' ? vttToText(response.data) : response.data;
      if (target === null) {
        process.stdout.write(text);
        return;
      }
      fs.writeFileSync(target, text);
      info(`Saved transcript to ${target}`);
    }));

  meetings
    .command('recordings <meeting>')
    .description(`List a meeting's recordings (<meeting>: ${MEETING_ARG})`)
    .addOption(outputOption())
    .action(handleErrors(async (meeting: string, opts) => {
      const client = getClient();
      const response = await client.request(`${await meetingPath(client, meeting)}/recordings`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  meetings
    .command('recording <meeting> [recordingId]')
    .description(`Download a recording (default: the latest) to a file (<meeting>: ${MEETING_ARG})`)
    .requiredOption('--out <path>', 'Output file')
    .option('--force', 'Overwrite an existing file')
    .action(handleErrors(async (meeting: string, recordingId: string | undefined, opts) => {
      const target = safeOutputPath(opts.out, { out: opts.out, force: opts.force });
      const client = getClient();
      const base = `${await meetingPath(client, meeting)}/recordings`;
      const rid = await latestId(client, base, recordingId, 'recording');
      const response = await client.request(`${base}/${rid}/content`, { responseType: 'stream' });
      if (!response.stream) {
        console.log(formatOutput(response.data, 'json')); // --dry-run
        return;
      }
      await streamTo(response.stream, target);
    }));

  addListOptions(meetings
    .command('all-transcripts')
    .description('List transcripts of all meetings you organized (beta API; needs --api-version beta)')
    .option('--start <datetime>', `Only transcripts created after: ${TIME_HELP}`)
    .option('--end <datetime>', `Only transcripts created before: ${TIME_HELP}`)
    .option('--timezone <tz>', `Zone for --start/--end; ${TZ_HELP}`))
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      if (!client.buildUrl('').includes('/beta')) {
        throw new CliError('meetings all-transcripts uses the Graph beta API (getAllTranscripts); pass --api-version beta', {
          code: 'usage', exitCode: ExitCode.USAGE, fix: 'pass --api-version beta',
        });
      }
      const tz = resolveTimeZone(opts.timezone);
      const range: string[] = [];
      if (opts.start) range.push(`startDateTime=${parseTime(opts.start, { tz }).date.toISOString()}`);
      if (opts.end) range.push(`endDateTime=${parseTime(opts.end, { tz }).date.toISOString()}`);
      const me = await client.request('/me', { queryParams: { $select: 'id' } });
      const id = (me.data as { id?: string }).id ?? 'ME_ID';
      const args = [`meetingOrganizerUserId='${odataQuote(id)}'`, ...range];
      await printList(client, `/users/${encodeURIComponent(id)}/onlineMeetings/getAllTranscripts(${args.join(',')})`, {}, opts);
    }));
}
