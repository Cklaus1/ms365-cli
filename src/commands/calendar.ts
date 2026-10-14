import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat, printList } from '../formatter.js';
import { outputOption, confirm, info, validateEmail, resolveTimeZone, preferTimeZone, parseTime } from '../utils.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';

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

/** Rooms need Place.Read.All, which only the org-mode login requests. */
async function placesRequest(client: GraphClient, endpoint: string) {
  try {
    return await client.request(endpoint);
  } catch (err) {
    if (err instanceof CliError && err.status === 403) err.fix = 'requires --org-mode login with Place.Read.All';
    throw err;
  }
}

export function registerCalendarCommands(program: Command, getClient: () => GraphClient): void {
  const cal = program.command('calendar').alias('cal').description('Calendar operations');

  cal
    .command('list')
    .description('List calendars')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/calendars');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  cal
    .command('events')
    .description('List calendar events')
    .option('--top <n>', 'Number of events', '10')
    .option('--filter <filter>', 'OData filter')
    .option('--select <fields>', 'Fields to return')
    .option('--orderby <field>', 'Order by field')
    .option('--calendar <calendarId>', 'Specific calendar ID')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .option('--timezone <tz>', `Zone for returned times; ${TZ_HELP}`)
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

      const client = getClient();
      const endpoint = opts.calendar
        ? `/me/calendars/${encodeURIComponent(opts.calendar)}/events`
        : '/me/events';

      const response = await client.request(endpoint, {
        method: 'POST',
        body: JSON.stringify(event),
      });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
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
          toGraph = eventTz ? (input) => eventZoneTime(input, eventTz) : (input) => graphTime(input, resolveTimeZone());
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
      info('Event deleted.');
    }));

  cal
    .command('view')
    .description('Get calendar view (time range)')
    .requiredOption('--start <datetime>', `Start: ${TIME_HELP}`)
    .requiredOption('--end <datetime>', `End: ${TIME_HELP}`)
    .option('--select <fields>', 'Fields to return')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .option('--timezone <tz>', `Zone for --start/--end and returned times; ${TZ_HELP}`)
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
      info('Event accepted.');
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
        info('Event declined with proposed new time.');
      } else {
        const client = getClient();
        await client.request(`/me/events/${encodeURIComponent(eventId)}/decline`, {
          method: 'POST',
          body: JSON.stringify({
            comment: opts.comment || '',
            sendResponse: opts.notify !== false,
          }),
        });
        info('Event declined.');
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
      info('Tentatively accepted.');
    }));

  // ── Free/Busy (scheduling) ─────────────────────────────────
  cal
    .command('free-busy')
    .description('Get free/busy schedule for one or more users')
    .requiredOption('--emails <addresses>', 'Comma-separated email addresses to check')
    .requiredOption('--start <datetime>', `Start: ${TIME_HELP}`)
    .requiredOption('--end <datetime>', `End: ${TIME_HELP}`)
    .option('--timezone <tz>', TZ_HELP)
    .option('--interval <min>', 'Time slot interval in minutes', '30')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const schedules = opts.emails.split(',').map((e: string) => e.trim());
      const tz = resolveTimeZone(opts.timezone);

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

  // ── Room finder ─────────────────────────────────────────────
  cal
    .command('rooms')
    .description('List meeting rooms (requires --org-mode)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const response = await placesRequest(getClient(), '/places/microsoft.graph.room');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  cal
    .command('room-lists')
    .description('List room lists / buildings (requires --org-mode)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const response = await placesRequest(getClient(), '/places/microsoft.graph.roomlist');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // ── Calendar CRUD ───────────────────────────────────────────
  cal
    .command('create-calendar')
    .description('Create a secondary calendar')
    .requiredOption('--name <name>', 'Calendar display name')
    .option('--color <color>', 'Color: auto, lightBlue, lightGreen, lightOrange, lightGray, lightYellow, lightTeal, lightPink, lightBrown, lightRed, maxColor')
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
    .option('--color <color>', 'Color')
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
