import { Command, InvalidArgumentError, Option } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat, printList } from '../formatter.js';
import {
  addListOptions, outputOption, confirm, info, success, validateEmail, resolveTimeZone, preferTimeZone, parseTime, parseGraphDateTime, toIsoWithOffset, fromGet,
} from '../utils.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import { fingerprint, isDryRun, withIdempotency } from '../idempotency.js';

const TZ_HELP = 'IANA time zone (default: MS365_CLI_TIMEZONE or the local zone)';
const TIME_HELP = 'ISO 8601 or natural ("tomorrow 15:00", "next mon 9:30", "+2h")';

/** A CLI time expression as a Graph dateTimeTimeZone in `tz`. */
function graphTime(input: string, tz: string): { dateTime: string; timeZone: string } {
  const { dateTime, timeZone } = parseTime(input, { tz });
  return { dateTime, timeZone };
}

/** Graph reports a UTC event zone as this URI rather than an IANA name. */
const GRAPH_UTC = 'tzone://Microsoft/Utc';

/**
 * Time for an update when the caller gave no --timezone: keep the event's own zone.
 * Windows zone names ("Eastern Standard Time") can't drive parseTime, so only a
 * plain wall-clock ISO time is accepted for those, and sent as-is.
 */
function eventZoneTime(input: string, eventTz: string): { dateTime: string; timeZone: string } {
  const tz = eventTz === GRAPH_UTC ? 'UTC' : eventTz;
  try {
    resolveTimeZone(tz);
  } catch {
    if (/^\d{4}-\d\d-\d\dT\d\d:\d\d(:\d\d(\.\d+)?)?$/.test(input.trim())) return { dateTime: input.trim(), timeZone: tz };
    throw new CliError(`The event's time zone "${tz}" is not an IANA name; give an ISO time (2026-10-20T09:00) or pass --timezone.`,
      { code: 'usage', exitCode: ExitCode.USAGE });
  }
  return graphTime(input, tz);
}

/** getSchedule's availabilityViewInterval: whole minutes, 5 to 1440 (†). */
function parseInterval(value: string): string {
  const n = Number(value);
  if (!/^\d+$/.test(value.trim()) || n < 5 || n > 1440) throw new InvalidArgumentError('must be an integer from 5 to 1440.');
  return String(n);
}


const CALENDAR_COLORS = ['auto', 'lightBlue', 'lightGreen', 'lightOrange', 'lightGray', 'lightYellow', 'lightTeal', 'lightPink', 'lightBrown', 'lightRed', 'maxColor'];
const colorOption = () => new Option('--color <color>', 'Calendar color').choices(CALENDAR_COLORS);

/** "30m", "1h", "1h30m" as an ISO 8601 duration (PT1H30M). */
function isoDuration(input: string): string {
  const m = input.trim().toLowerCase().match(/^(?:(\d+)h)?(?:(\d+)m)?$/);
  if (!m || (!m[1] && !m[2]) || (+(m[1] ?? 0) === 0 && +(m[2] ?? 0) === 0)) {
    throw new CliError(`Invalid duration "${input}". Use minutes and/or hours: 30m, 1h, 1h30m.`, { code: 'usage', exitCode: ExitCode.USAGE });
  }
  return `PT${m[1] ? `${+m[1]}H` : ''}${m[2] ? `${+m[2]}M` : ''}`;
}

interface MeetingTimeSuggestion {
  confidence?: number;
  organizerAvailability?: string;
  meetingTimeSlot: { start: { dateTime: string; timeZone?: string }; end: { dateTime: string; timeZone?: string } };
  attendeeAvailability?: Array<{ attendee?: { emailAddress?: { address?: string } }; availability?: string }>;
}

/** Rooms need Place.Read.All, which only the org-mode login requests. */
async function withPlacesHint(fn: () => Promise<void>): Promise<void> {
  try {
    await fn();
  } catch (err) {
    if (err instanceof CliError && err.status === 403) err.fix = 'requires --org-mode login with Place.Read.All';
    throw err;
  }
}

export function registerCalendarCommands(program: Command, getClient: () => GraphClient): void {
  const cal = program.command('calendar').alias('cal').description('Calendar operations');

  addListOptions(cal
    .command('list')
    .description('List calendars'))
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      await printList(getClient(), '/me/calendars', {}, opts);
    }));

  addListOptions(cal
    .command('events')
    .description('List calendar events')
    .option('--filter <filter>', 'OData filter')
    .option('--select <fields>', 'Fields to return')
    .option('--orderby <field>', 'Order by field')
    .option('--calendar <calendarId>', 'Specific calendar ID')
    .option('--timezone <tz>', `Zone for returned times; ${TZ_HELP}`), { description: 'Number of events', default: '10' })
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const headers = preferTimeZone(resolveTimeZone(opts.timezone));
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.select) queryParams['$select'] = opts.select;
      if (opts.orderby) queryParams['$orderby'] = opts.orderby;

      const endpoint = opts.calendar
        ? `/me/calendars/${encodeURIComponent(opts.calendar)}/events`
        : '/me/events';

      await printList(client, endpoint, { queryParams, headers }, opts);
    }));

  cal
    .command('get <eventId>')
    .description('Get a specific event')
    .option('--calendar <calendarId>', 'Specific calendar ID')
    .option('--select <fields>', 'Fields to return')
    .option('--timezone <tz>', `Zone for returned times; ${TZ_HELP}`)
    .addOption(outputOption())
    .action(handleErrors(async (eventId: string, opts) => {
      const client = getClient();
      const headers = preferTimeZone(resolveTimeZone(opts.timezone));
      const queryParams: Record<string, string> = {};
      if (opts.select) queryParams['$select'] = opts.select;

      const endpoint = opts.calendar
        ? `/me/calendars/${encodeURIComponent(opts.calendar)}/events/${encodeURIComponent(eventId)}`
        : `/me/events/${encodeURIComponent(eventId)}`;

      const response = await client.request(endpoint, { queryParams, headers });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  cal
    .command('create')
    .description('Create a calendar event')
    .requiredOption('--subject <subject>', 'Event subject')
    .requiredOption('--start <datetime>', `Start: ${TIME_HELP}`)
    .requiredOption('--end <datetime>', `End: ${TIME_HELP}`)
    .option('--timezone <tz>', TZ_HELP)
    .option('--body <body>', 'Event body/description')
    .option('--location <location>', 'Event location')
    .option('--attendees <emails>', 'Comma-separated attendee emails')
    .option('--calendar <calendarId>', 'Calendar ID')
    .option('--html', 'Body is HTML')
    .option('--teams', 'Make it a Teams online meeting (adds a join link)')
    .option('--idempotency-key <key>', 'Retry-safe: sent as the event transactionId; a repeat within 24 h returns the first event')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const tz = resolveTimeZone(opts.timezone);
      const event: Record<string, unknown> = {
        subject: opts.subject,
        start: graphTime(opts.start, tz),
        end: graphTime(opts.end, tz),
      };

      if (opts.body) {
        event.body = { contentType: opts.html ? 'HTML' : 'Text', content: opts.body };
      }
      if (opts.location) {
        event.location = { displayName: opts.location };
      }
      if (opts.attendees) {
        event.attendees = opts.attendees.split(',').map((e: string) => ({
          emailAddress: { address: validateEmail(e) },
          type: 'required',
        }));
      }
      if (opts.teams) {
        event.isOnlineMeeting = true;
        event.onlineMeetingProvider = 'teamsForBusiness';
      }

      const client = getClient();
      const endpoint = opts.calendar
        ? `/me/calendars/${encodeURIComponent(opts.calendar)}/events`
        : '/me/events';

      const create = () => client.request(endpoint, { method: 'POST', body: JSON.stringify(event) });
      const key: string | undefined = opts.idempotencyKey;
      if (!key) {
        console.log(formatOutput((await create()).data, opts.output as OutputFormat));
        return;
      }
      // Graph de-duplicates on transactionId server-side; the ledger also skips the request on a local retry (D23).
      event.transactionId = key;
      if (isDryRun(client)) {
        console.log(formatOutput((await create()).data, opts.output as OutputFormat));
        return;
      }
      let created: unknown;
      const outcome = await withIdempotency(key, 'cal create', fingerprint({ endpoint, event }), async () => {
        created = (await create()).data;
        const { id, webLink } = created as { id?: string; webLink?: string };
        return { id, webLink };
      });
      if (outcome.status === 'replay') info(`Event already created at ${outcome.ts}; not created again.`);
      console.log(formatOutput(outcome.status === 'replay' ? outcome.result : created, opts.output as OutputFormat));
    }));

  cal
    .command('update <eventId>')
    .description('Update a calendar event')
    .option('--subject <subject>', 'Event subject')
    .option('--start <datetime>', `Start: ${TIME_HELP}`)
    .option('--end <datetime>', `End: ${TIME_HELP}`)
    .option('--timezone <tz>', 'IANA time zone for --start/--end (default: the event\'s own zone)')
    .option('--body <body>', 'Event body')
    .option('--location <location>', 'Event location')
    .option('--calendar <calendarId>', 'Specific calendar ID')
    .addOption(outputOption())
    .action(handleErrors(async (eventId: string, opts) => {
      const client = getClient();
      const endpoint = opts.calendar
        ? `/me/calendars/${encodeURIComponent(opts.calendar)}/events/${encodeURIComponent(eventId)}`
        : `/me/events/${encodeURIComponent(eventId)}`;

      const patch: Record<string, unknown> = {};
      if (opts.subject) patch.subject = opts.subject;
      if (opts.start || opts.end) {
        let toGraph: (input: string) => { dateTime: string; timeZone: string };
        if (opts.timezone) {
          const tz = resolveTimeZone(opts.timezone);
          toGraph = (input) => graphTime(input, tz);
        } else {
          // Without a Prefer header start.timeZone is always UTC; originalStartTimeZone is the event's real zone.
          const existing = await client.request(endpoint, { queryParams: { $select: 'originalStartTimeZone' } });
          const eventTz = (existing.data as { originalStartTimeZone?: string }).originalStartTimeZone;
          if (client.isPreview(existing)) {
            toGraph = (input) => ({ dateTime: graphTime(input, resolveTimeZone()).dateTime, timeZone: fromGet() });
          } else {
            toGraph = eventTz ? (input) => eventZoneTime(input, eventTz) : (input) => graphTime(input, resolveTimeZone());
          }
        }
        if (opts.start) patch.start = toGraph(opts.start);
        if (opts.end) patch.end = toGraph(opts.end);
      }
      if (opts.body) patch.body = { contentType: 'Text', content: opts.body };
      if (opts.location) patch.location = { displayName: opts.location };

      const response = await client.request(endpoint, {
        method: 'PATCH',
        body: JSON.stringify(patch),
      });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  cal
    .command('delete <eventId>')
    .description('Delete a calendar event')
    .option('--calendar <calendarId>', 'Specific calendar ID')
    .action(handleErrors(async (eventId: string, opts) => {
      if (!await confirm(`Delete event ${eventId.substring(0, 16)}...?`)) return;
      const client = getClient();
      const endpoint = opts.calendar
        ? `/me/calendars/${encodeURIComponent(opts.calendar)}/events/${encodeURIComponent(eventId)}`
        : `/me/events/${encodeURIComponent(eventId)}`;

      await client.request(endpoint, { method: 'DELETE' });
      success('Event deleted.');
    }));

  addListOptions(cal
    .command('view')
    .description('Get calendar view (time range)')
    .requiredOption('--start <datetime>', `Start: ${TIME_HELP}`)
    .requiredOption('--end <datetime>', `End: ${TIME_HELP}`)
    .option('--select <fields>', 'Fields to return')
    .option('--timezone <tz>', `Zone for --start/--end and returned times; ${TZ_HELP}`))
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const tz = resolveTimeZone(opts.timezone);
      const queryParams: Record<string, string> = {
        startDateTime: parseTime(opts.start, { tz }).date.toISOString(),
        endDateTime: parseTime(opts.end, { tz }).date.toISOString(),
      };
      if (opts.select) queryParams['$select'] = opts.select;
      const headers = preferTimeZone(tz);

      const client = getClient();
      await printList(client, '/me/calendarView', { queryParams, headers }, opts);
    }));

  // ── Respond to event (accept/decline/tentative) ─────────────
  cal
    .command('accept <eventId>')
    .description('Accept a meeting invite')
    .option('--comment <text>', 'Response comment')
    .option('--no-notify', 'Don\'t send response to organizer')
    .action(handleErrors(async (eventId: string, opts) => {
      const client = getClient();
      await client.request(`/me/events/${encodeURIComponent(eventId)}/accept`, {
        method: 'POST',
        body: JSON.stringify({
          comment: opts.comment || '',
          sendResponse: opts.notify !== false,
        }),
      });
      success('Event accepted.');
    }));

  cal
    .command('decline <eventId>')
    .description('Decline a meeting invite')
    .option('--comment <text>', 'Response comment')
    .option('--no-notify', 'Don\'t send response to organizer')
    .option('--propose-start <datetime>', `Propose new start time: ${TIME_HELP}`)
    .option('--propose-end <datetime>', `Propose new end time: ${TIME_HELP}`)
    .option('--timezone <tz>', `Zone for proposed times; ${TZ_HELP}`)
    .action(handleErrors(async (eventId: string, opts) => {
      if (opts.proposeStart || opts.proposeEnd) {
        if (!opts.proposeStart || !opts.proposeEnd) {
          throw new CliError('--propose-start and --propose-end must be given together', { code: 'usage', exitCode: ExitCode.USAGE });
        }
        // Graph rejects a proposed time unless the response is sent to the organizer.
        if (opts.notify === false) {
          throw new CliError('--no-notify cannot be combined with --propose-start/--propose-end', { code: 'usage', exitCode: ExitCode.USAGE });
        }
        const tz = resolveTimeZone(opts.timezone);
        const proposedNewTime = { start: graphTime(opts.proposeStart, tz), end: graphTime(opts.proposeEnd, tz) };
        const client = getClient();
        await client.request(`/me/events/${encodeURIComponent(eventId)}/decline`, {
          method: 'POST',
          body: JSON.stringify({
            comment: opts.comment || '',
            sendResponse: true,
            proposedNewTime,
          }),
        });
        success('Event declined with proposed new time.');
      } else {
        const client = getClient();
        await client.request(`/me/events/${encodeURIComponent(eventId)}/decline`, {
          method: 'POST',
          body: JSON.stringify({
            comment: opts.comment || '',
            sendResponse: opts.notify !== false,
          }),
        });
        success('Event declined.');
      }
    }));

  cal
    .command('tentative <eventId>')
    .description('Tentatively accept a meeting invite')
    .option('--comment <text>', 'Response comment')
    .option('--no-notify', 'Don\'t send response to organizer')
    .action(handleErrors(async (eventId: string, opts) => {
      const client = getClient();
      await client.request(`/me/events/${encodeURIComponent(eventId)}/tentativelyAccept`, {
        method: 'POST',
        body: JSON.stringify({
          comment: opts.comment || '',
          sendResponse: opts.notify !== false,
        }),
      });
      success('Tentatively accepted.');
    }));

  // ── Free/Busy (scheduling) ─────────────────────────────────
  cal
    .command('free-busy')
    .description('Get free/busy schedule for one or more users')
    .requiredOption('--emails <addresses>', 'Comma-separated email addresses to check')
    .requiredOption('--start <datetime>', `Start: ${TIME_HELP}`)
    .requiredOption('--end <datetime>', `End: ${TIME_HELP}`)
    .option('--timezone <tz>', TZ_HELP)
    .option('--interval <min>', 'Time slot interval in minutes (5-1440)', parseInterval, '30')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const schedules = opts.emails.split(',').map((e: string) => validateEmail(e));
      const tz = resolveTimeZone(opts.timezone);
      const client = getClient();

      const response = await client.request('/me/calendar/getSchedule', {
        method: 'POST',
        body: JSON.stringify({
          schedules,
          startTime: graphTime(opts.start, tz),
          endTime: graphTime(opts.end, tz),
          availabilityViewInterval: parseInt(opts.interval),
        }),
      });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  cal
    .command('find-time')
    .description('Suggest meeting times when the attendees are free (findMeetingTimes)')
    .requiredOption('--attendees <emails>', 'Comma-separated attendee emails')
    .requiredOption('--duration <d>', 'Meeting length: 30m, 1h, 1h30m')
    .option('--start <datetime>', `Window start (default: now): ${TIME_HELP}`)
    .option('--end <datetime>', `Window end (default: 7 days after the start): ${TIME_HELP}`)
    .option('--timezone <tz>', `Zone for the window and returned times; ${TZ_HELP}`)
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const tz = resolveTimeZone(opts.timezone);
      const meetingDuration = isoDuration(opts.duration);
      const attendees = opts.attendees.split(',').map((e: string) => ({
        type: 'required',
        emailAddress: { address: validateEmail(e) },
      }));
      const start = opts.start ? parseTime(opts.start, { tz }) : parseTime('now', { tz });
      const end = opts.end ? parseTime(opts.end, { tz }) : parseTime('+7d', { tz, now: start.date });

      const client = getClient();
      const response = await client.request('/me/findMeetingTimes', {
        method: 'POST',
        headers: preferTimeZone(tz),
        body: JSON.stringify({
          attendees,
          timeConstraint: {
            activityDomain: 'work',
            timeSlots: [{ start: { dateTime: start.dateTime, timeZone: tz }, end: { dateTime: end.dateTime, timeZone: tz } }],
          },
          meetingDuration,
        }),
      });

      if (client.isPreview(response)) {
        console.log(formatOutput(response.data, opts.output as OutputFormat));
        return;
      }
      const data = response.data as { meetingTimeSuggestions?: MeetingTimeSuggestion[]; emptySuggestionsReason?: string };
      // Normalize to ISO with the user's offset, whether or not Graph honored the Prefer header.
      const iso = (t: { dateTime: string; timeZone?: string }) => toIsoWithOffset(parseGraphDateTime(t, tz), tz);
      const value = (data.meetingTimeSuggestions ?? []).map(s => ({
        start: iso(s.meetingTimeSlot.start),
        end: iso(s.meetingTimeSlot.end),
        confidence: s.confidence,
        organizerAvailability: s.organizerAvailability,
        attendeeAvailability: (s.attendeeAvailability ?? []).map(a => ({
          attendee: a.attendee?.emailAddress?.address,
          availability: a.availability,
        })),
      }));
      const out: Record<string, unknown> = { value };
      if (data.emptySuggestionsReason) out.emptySuggestionsReason = data.emptySuggestionsReason;
      console.log(formatOutput(out, opts.output as OutputFormat));
    }));

  // ── Room finder ─────────────────────────────────────────────
  addListOptions(cal
    .command('rooms')
    .description('List meeting rooms (requires --org-mode)'))
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      await withPlacesHint(() => printList(getClient(), '/places/microsoft.graph.room', {}, opts));
    }));

  addListOptions(cal
    .command('room-lists')
    .description('List room lists / buildings (requires --org-mode)'))
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      await withPlacesHint(() => printList(getClient(), '/places/microsoft.graph.roomlist', {}, opts));
    }));

  // ── Calendar CRUD ───────────────────────────────────────────
  cal
    .command('create-calendar')
    .description('Create a secondary calendar')
    .requiredOption('--name <name>', 'Calendar display name')
    .addOption(colorOption())
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const body: Record<string, unknown> = { name: opts.name };
      if (opts.color) body.color = opts.color;

      const response = await client.request('/me/calendars', {
        method: 'POST',
        body: JSON.stringify(body),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  cal
    .command('update-calendar <calendarId>')
    .description('Update a calendar')
    .option('--name <name>', 'Display name')
    .addOption(colorOption())
    .addOption(outputOption())
    .action(handleErrors(async (calendarId: string, opts) => {
      const client = getClient();
      const patch: Record<string, unknown> = {};
      if (opts.name) patch.name = opts.name;
      if (opts.color) patch.color = opts.color;

      const response = await client.request(`/me/calendars/${encodeURIComponent(calendarId)}`, {
        method: 'PATCH',
        body: JSON.stringify(patch),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

}
