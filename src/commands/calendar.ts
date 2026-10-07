import { Command, Option } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption, confirm, done, parseRecipients, validateEmail } from '../utils.js';
import { handleErrors } from '../errors.js';
import { integer, paginationOptions, listRequest, type ListOptions } from '../list.js';
import { eventTime, preferTimezone } from '../graph-helpers.js';
import { graph } from '../command-tree.js';

/** Graph `calendarColor` values (https://learn.microsoft.com/en-us/graph/api/resources/calendar). */
const CALENDAR_COLORS = ['auto', 'lightBlue', 'lightGreen', 'lightOrange', 'lightGray', 'lightYellow', 'lightTeal', 'lightPink', 'lightBrown', 'lightRed', 'maxColor'];

/**
 * `--timezone` for event times. No Commander default: the local zone is resolved in the action
 * (`eventTime`), so the command tree and generated skills stay machine-independent.
 */
const timezoneOption = (what: string) => new Option('--timezone <tz>', `Time zone of ${what} (default: local time zone)`);

interface EventsOptions extends ListOptions {
  filter?: string;
  select?: string;
  orderby?: string;
  calendar?: string;
}

interface ViewOptions extends ListOptions {
  start: string;
  end: string;
  select?: string;
}

interface UpdateOptions {
  subject?: string;
  start?: string;
  end?: string;
  timezone?: string;
  body?: string;
  location?: string;
  calendar?: string;
  output: string;
}

interface DeclineOptions {
  comment?: string;
  notify: boolean;
  proposeStart?: string;
  proposeEnd?: string;
  timezone?: string;
}

export function registerCalendarCommands(program: Command, getClient: () => GraphClient): void {
  const cal = program.command('calendar').alias('cal').description('Calendar operations');

  paginationOptions(
    graph(cal.command('list'), ['GET /me/calendars'])
      .description('List calendars'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (opts: ListOptions) => {
      await listRequest(getClient(), '/me/calendars', {}, opts);
    }));

  // Event reads ask Graph for times in the local zone (`Prefer: outlook.timezone`) instead of UTC.
  paginationOptions(
    graph(cal.command('events'), ['GET /me/events', 'GET /me/calendars/{calendar-id}/events'])
      .description('List calendar events (times in the local time zone)')
      .option('--filter <filter>', 'OData filter')
      .option('--select <fields>', 'Fields to return')
      .option('--orderby <field>', 'Order by field')
      .option('--calendar <calendarId>', 'Specific calendar ID'),
    10,
  )
    .addOption(outputOption())
    .action(handleErrors(async (opts: EventsOptions) => {
      const queryParams: Record<string, string> = {};
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.select) queryParams['$select'] = opts.select;
      if (opts.orderby) queryParams['$orderby'] = opts.orderby;

      const endpoint = opts.calendar
        ? `/me/calendars/${encodeURIComponent(opts.calendar)}/events`
        : '/me/events';

      await listRequest(getClient(), endpoint, { queryParams, headers: preferTimezone() }, opts);
    }));

  graph(cal.command('get <eventId>'), [
    'GET /me/events/{event-id}',
    'GET /me/calendars/{calendar-id}/events/{event-id}',
  ])
    .description('Get a specific event (times in the local time zone)')
    .option('--calendar <calendarId>', 'Specific calendar ID')
    .option('--select <fields>', 'Fields to return')
    .addOption(outputOption())
    .action(handleErrors(async (eventId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.select) queryParams['$select'] = opts.select;

      const endpoint = opts.calendar
        ? `/me/calendars/${encodeURIComponent(opts.calendar)}/events/${encodeURIComponent(eventId)}`
        : `/me/events/${encodeURIComponent(eventId)}`;

      const response = await client.request(endpoint, { queryParams, headers: preferTimezone() });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  graph(cal.command('create'), ['POST /me/events', 'POST /me/calendars/{calendar-id}/events'])
    .description('Create a calendar event')
    .requiredOption('--subject <subject>', 'Event subject')
    .requiredOption('--start <datetime>', 'Start datetime (ISO 8601)')
    .requiredOption('--end <datetime>', 'End datetime (ISO 8601)')
    .addOption(timezoneOption('--start/--end'))
    .option('--body <body>', 'Event body/description')
    .option('--location <location>', 'Event location')
    .option('--attendees <emails>', 'Comma-separated attendee emails')
    .option('--calendar <calendarId>', 'Calendar ID')
    .option('--html', 'Body is HTML')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const event: Record<string, unknown> = {
        subject: opts.subject,
        start: eventTime(opts.start, opts.timezone),
        end: eventTime(opts.end, opts.timezone),
      };

      if (opts.body) {
        event.body = { contentType: opts.html ? 'HTML' : 'Text', content: opts.body };
      }
      if (opts.location) {
        event.location = { displayName: opts.location };
      }
      if (opts.attendees) {
        event.attendees = parseRecipients(opts.attendees).map((r) => ({ ...r, type: 'required' }));
      }

      const endpoint = opts.calendar
        ? `/me/calendars/${encodeURIComponent(opts.calendar)}/events`
        : '/me/events';

      const response = await getClient().request(endpoint, {
        method: 'POST',
        headers: preferTimezone(),
        body: JSON.stringify(event),
      });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  graph(cal.command('update <eventId>'), [
    'PATCH /me/events/{event-id}',
    'PATCH /me/calendars/{calendar-id}/events/{event-id}',
  ])
    .description('Update a calendar event')
    .option('--subject <subject>', 'Event subject')
    .option('--start <datetime>', 'Start datetime')
    .option('--end <datetime>', 'End datetime')
    .addOption(timezoneOption('--start/--end'))
    .option('--body <body>', 'Event body')
    .option('--location <location>', 'Event location')
    .option('--calendar <calendarId>', 'Specific calendar ID')
    .addOption(outputOption())
    .action(handleErrors(async (eventId: string, opts: UpdateOptions) => {
      if (opts.timezone && !opts.start && !opts.end) {
        throw new Error('--timezone applies to --start/--end; give at least one of them.');
      }
      const patch: Record<string, unknown> = {};
      if (opts.subject) patch.subject = opts.subject;
      if (opts.start) patch.start = eventTime(opts.start, opts.timezone);
      if (opts.end) patch.end = eventTime(opts.end, opts.timezone);
      if (opts.body) patch.body = { contentType: 'Text', content: opts.body };
      if (opts.location) patch.location = { displayName: opts.location };
      if (Object.keys(patch).length === 0) {
        throw new Error('Nothing to update: give at least one of --subject, --start, --end, --body, --location.');
      }

      const endpoint = opts.calendar
        ? `/me/calendars/${encodeURIComponent(opts.calendar)}/events/${encodeURIComponent(eventId)}`
        : `/me/events/${encodeURIComponent(eventId)}`;

      const response = await getClient().request(endpoint, {
        method: 'PATCH',
        headers: preferTimezone(),
        body: JSON.stringify(patch),
      });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  graph(cal.command('delete <eventId>'), [
    'DELETE /me/events/{event-id}',
    'DELETE /me/calendars/{calendar-id}/events/{event-id}',
  ])
    .description('Delete a calendar event')
    .option('--calendar <calendarId>', 'Specific calendar ID')
    .action(handleErrors(async (eventId: string, opts) => {
      await confirm(`Delete event ${eventId.substring(0, 16)}...?`);
      const client = getClient();
      const endpoint = opts.calendar
        ? `/me/calendars/${encodeURIComponent(opts.calendar)}/events/${encodeURIComponent(eventId)}`
        : `/me/events/${encodeURIComponent(eventId)}`;

      const res = await client.request(endpoint, { method: 'DELETE' });
      done(res, 'Event deleted.');
    }));

  paginationOptions(
    graph(cal.command('view'), ['GET /me/calendarView'])
      .description('Get calendar view (time range; times in the local time zone)')
      .requiredOption('--start <datetime>', 'Start datetime (ISO 8601)')
      .requiredOption('--end <datetime>', 'End datetime (ISO 8601)')
      .option('--select <fields>', 'Fields to return'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (opts: ViewOptions) => {
      const queryParams: Record<string, string> = {
        startDateTime: opts.start,
        endDateTime: opts.end,
      };
      if (opts.select) queryParams['$select'] = opts.select;
      await listRequest(getClient(), '/me/calendarView', { queryParams, headers: preferTimezone() }, opts);
    }));

  // ── Respond to event (accept/decline/tentative) ─────────────
  graph(cal.command('accept <eventId>'), ['POST /me/events/{event-id}/accept'])
    .description('Accept a meeting invite')
    .option('--comment <text>', 'Response comment')
    .option('--no-notify', 'Don\'t send response to organizer')
    .action(handleErrors(async (eventId: string, opts) => {
      const client = getClient();
      const res = await client.request(`/me/events/${encodeURIComponent(eventId)}/accept`, {
        method: 'POST',
        body: JSON.stringify({
          comment: opts.comment || '',
          sendResponse: opts.notify !== false,
        }),
      });
      done(res, 'Event accepted.');
    }));

  graph(cal.command('decline <eventId>'), ['POST /me/events/{event-id}/decline'])
    .description('Decline a meeting invite')
    .option('--comment <text>', 'Response comment')
    .option('--no-notify', 'Don\'t send response to organizer (not with --propose-start)')
    .option('--propose-start <datetime>', 'Propose new start time (ISO 8601, without offset)')
    .option('--propose-end <datetime>', 'Propose new end time (ISO 8601, without offset)')
    .addOption(timezoneOption('the proposed times'))
    .action(handleErrors(async (eventId: string, opts: DeclineOptions) => {
      const body: Record<string, unknown> = {
        comment: opts.comment || '',
        sendResponse: opts.notify !== false,
      };
      if (opts.proposeStart || opts.proposeEnd) {
        if (!opts.proposeStart || !opts.proposeEnd) {
          throw new Error('--propose-start and --propose-end must be given together.');
        }
        // Graph only delivers a proposed time inside the response to the organizer.
        if (opts.notify === false) {
          throw new Error('--propose-start cannot be combined with --no-notify: a proposed time is sent with the response.');
        }
        body.proposedNewTime = {
          start: eventTime(opts.proposeStart, opts.timezone),
          end: eventTime(opts.proposeEnd, opts.timezone),
        };
      }
      const res = await getClient().request(`/me/events/${encodeURIComponent(eventId)}/decline`, {
        method: 'POST',
        body: JSON.stringify(body),
      });
      done(res, body.proposedNewTime ? 'Event declined with proposed new time.' : 'Event declined.');
    }));

  graph(cal.command('tentative <eventId>'), ['POST /me/events/{event-id}/tentativelyAccept'])
    .description('Tentatively accept a meeting invite')
    .option('--comment <text>', 'Response comment')
    .option('--no-notify', 'Don\'t send response to organizer')
    .action(handleErrors(async (eventId: string, opts) => {
      const client = getClient();
      const res = await client.request(`/me/events/${encodeURIComponent(eventId)}/tentativelyAccept`, {
        method: 'POST',
        body: JSON.stringify({
          comment: opts.comment || '',
          sendResponse: opts.notify !== false,
        }),
      });
      done(res, 'Tentatively accepted.');
    }));

  // ── Free/Busy (scheduling) ─────────────────────────────────
  graph(cal.command('free-busy'), ['POST /me/calendar/getSchedule'])
    .description('Get free/busy schedule for one or more users')
    .requiredOption('--emails <addresses>', 'Comma-separated email addresses to check')
    .requiredOption('--start <datetime>', 'Start datetime (ISO 8601)')
    .requiredOption('--end <datetime>', 'End datetime (ISO 8601)')
    .addOption(timezoneOption('--start/--end'))
    .addOption(new Option('--interval <min>', 'Time slot interval in minutes (5-1440)').argParser(integer(5, 1440)).default(30))
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const schedules = opts.emails.split(',').map(validateEmail);

      const response = await getClient().request('/me/calendar/getSchedule', {
        method: 'POST',
        body: JSON.stringify({
          schedules,
          startTime: eventTime(opts.start, opts.timezone),
          endTime: eventTime(opts.end, opts.timezone),
          availabilityViewInterval: opts.interval,
        }),
      });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // ── Room finder (beta-only endpoints; work/school accounts) ─
  graph(cal.command('rooms'), ['GET /me/findRooms'])
    .description('List meeting rooms in your organization (Graph beta endpoint)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const response = await getClient().request('/me/findRooms', { apiVersion: 'beta' });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  graph(cal.command('room-lists'), ['GET /me/findRoomLists'])
    .description('List room lists / buildings in your organization (Graph beta endpoint)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const response = await getClient().request('/me/findRoomLists', { apiVersion: 'beta' });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // ── Calendar CRUD ───────────────────────────────────────────
  graph(cal.command('create-calendar'), ['POST /me/calendars'])
    .description('Create a secondary calendar')
    .requiredOption('--name <name>', 'Calendar display name')
    .addOption(new Option('--color <color>', 'Color').choices(CALENDAR_COLORS))
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

  graph(cal.command('update-calendar <calendarId>'), ['PATCH /me/calendars/{calendar-id}'])
    .description('Update a calendar')
    .option('--name <name>', 'Display name')
    .addOption(new Option('--color <color>', 'Color').choices(CALENDAR_COLORS))
    .addOption(outputOption())
    .action(handleErrors(async (calendarId: string, opts) => {
      const patch: Record<string, unknown> = {};
      if (opts.name) patch.name = opts.name;
      if (opts.color) patch.color = opts.color;
      if (Object.keys(patch).length === 0) {
        throw new Error('Nothing to update: give at least one of --name, --color.');
      }

      const response = await getClient().request(`/me/calendars/${encodeURIComponent(calendarId)}`, {
        method: 'PATCH',
        body: JSON.stringify(patch),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

}
