import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption, confirm, info } from '../utils.js';
import { handleErrors } from '../errors.js';

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
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.select) queryParams['$select'] = opts.select;
      if (opts.orderby) queryParams['$orderby'] = opts.orderby;

      const endpoint = opts.calendar
        ? `/me/calendars/${encodeURIComponent(opts.calendar)}/events`
        : '/me/events';

      const response = opts.all
        ? await client.requestAllPages(endpoint, { queryParams }, opts.pageLimit ? parseInt(opts.pageLimit) : undefined, opts.pageDelay ? parseInt(opts.pageDelay) : undefined)
        : await client.request(endpoint, { queryParams });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  cal
    .command('get <eventId>')
    .description('Get a specific event')
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

      const response = await client.request(endpoint, { queryParams });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  cal
    .command('create')
    .description('Create a calendar event')
    .requiredOption('--subject <subject>', 'Event subject')
    .requiredOption('--start <datetime>', 'Start datetime (ISO 8601)')
    .requiredOption('--end <datetime>', 'End datetime (ISO 8601)')
    .option('--timezone <tz>', 'Timezone', 'UTC')
    .option('--body <body>', 'Event body/description')
    .option('--location <location>', 'Event location')
    .option('--attendees <emails>', 'Comma-separated attendee emails')
    .option('--calendar <calendarId>', 'Calendar ID')
    .option('--html', 'Body is HTML')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const event: Record<string, unknown> = {
        subject: opts.subject,
        start: { dateTime: opts.start, timeZone: opts.timezone },
        end: { dateTime: opts.end, timeZone: opts.timezone },
      };

      if (opts.body) {
        event.body = { contentType: opts.html ? 'HTML' : 'Text', content: opts.body };
      }
      if (opts.location) {
        event.location = { displayName: opts.location };
      }
      if (opts.attendees) {
        event.attendees = opts.attendees.split(',').map((e: string) => ({
          emailAddress: { address: e.trim() },
          type: 'required',
        }));
      }

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
    .option('--start <datetime>', 'Start datetime')
    .option('--end <datetime>', 'End datetime')
    .option('--timezone <tz>', 'Timezone')
    .option('--body <body>', 'Event body')
    .option('--location <location>', 'Event location')
    .option('--calendar <calendarId>', 'Specific calendar ID')
    .addOption(outputOption())
    .action(handleErrors(async (eventId: string, opts) => {
      const client = getClient();
      const patch: Record<string, unknown> = {};
      if (opts.subject) patch.subject = opts.subject;
      if (opts.start) patch.start = { dateTime: opts.start, timeZone: opts.timezone || 'UTC' };
      if (opts.end) patch.end = { dateTime: opts.end, timeZone: opts.timezone || 'UTC' };
      if (opts.body) patch.body = { contentType: 'Text', content: opts.body };
      if (opts.location) patch.location = { displayName: opts.location };

      const endpoint = opts.calendar
        ? `/me/calendars/${encodeURIComponent(opts.calendar)}/events/${encodeURIComponent(eventId)}`
        : `/me/events/${encodeURIComponent(eventId)}`;

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
    .requiredOption('--start <datetime>', 'Start datetime (ISO 8601)')
    .requiredOption('--end <datetime>', 'End datetime (ISO 8601)')
    .option('--select <fields>', 'Fields to return')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {
        startDateTime: opts.start,
        endDateTime: opts.end,
      };
      if (opts.select) queryParams['$select'] = opts.select;

      const response = opts.all
        ? await client.requestAllPages('/me/calendarView', { queryParams })
        : await client.request('/me/calendarView', { queryParams });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
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
    .option('--propose-start <datetime>', 'Propose new start time (ISO 8601)')
    .option('--propose-end <datetime>', 'Propose new end time (ISO 8601)')
    .action(handleErrors(async (eventId: string, opts) => {
      const client = getClient();
      if (opts.proposeStart) {
        if (!opts.proposeEnd) {
          throw new Error('--propose-end is required when using --propose-start');
        }
        await client.request(`/me/events/${encodeURIComponent(eventId)}/decline`, {
          method: 'POST',
          body: JSON.stringify({
            comment: opts.comment || '',
            sendResponse: opts.notify !== false,
            proposedNewTime: {
              start: { dateTime: opts.proposeStart, timeZone: 'UTC' },
              end: { dateTime: opts.proposeEnd, timeZone: 'UTC' },
            },
          }),
        });
        info('Event declined with proposed new time.');
      } else {
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
    .requiredOption('--start <datetime>', 'Start datetime (ISO 8601)')
    .requiredOption('--end <datetime>', 'End datetime (ISO 8601)')
    .option('--timezone <tz>', 'Timezone', 'UTC')
    .option('--interval <min>', 'Time slot interval in minutes', '30')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const schedules = opts.emails.split(',').map((e: string) => e.trim());

      const response = await client.request('/me/calendar/getSchedule', {
        method: 'POST',
        body: JSON.stringify({
          schedules,
          startTime: { dateTime: opts.start, timeZone: opts.timezone },
          endTime: { dateTime: opts.end, timeZone: opts.timezone },
          availabilityViewInterval: parseInt(opts.interval),
        }),
      });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // ── Room finder ─────────────────────────────────────────────
  cal
    .command('rooms')
    .description('List available meeting rooms (requires --org-mode)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/findRooms');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  cal
    .command('room-lists')
    .description('List room lists / buildings (requires --org-mode)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/findRoomLists');
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
