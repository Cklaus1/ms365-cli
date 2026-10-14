import { Command, Option } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput } from '../formatter.js';
import { CliError, ExitCode, handleErrors } from '../errors.js';
import {
  parseDate, outputOption, info, success, resolveTimeZone, preferTimeZone, parseGraphDateTime, dayRange, toGraphDateTime, fromGet,
} from '../utils.js';

/** Under --dry-run, stands in for the list the real run picks from GET /me/todo/lists (DR-10). */
const FIRST_LIST = '<first-list>';

/**
 * Read-aggregate workflows under --dry-run: the reads are traced, but a report built from placeholders
 * would be invented, so print nothing (text) or exactly {"dryRun":true} (json) (DR-12, F6).
 */
function previewReport(output: string): void {
  if (output === 'json') console.log(JSON.stringify({ dryRun: true }));
}

/** Parse a work-day bound "H", "HH" or "HH:MM" (24h) into minutes since midnight; usage error otherwise (V-20). */
function workMinute(flag: string, input: string): number {
  const m = /^(\d{1,2})(?::([0-5]\d))?$/.exec(input);
  const minutes = m ? Number(m[1]) * 60 + Number(m[2] ?? 0) : NaN;
  if (!(minutes <= 24 * 60)) {
    throw new CliError(`Invalid ${flag}: "${input}". Use an hour 0-24 or HH:MM (24h).`, { code: 'usage', exitCode: ExitCode.USAGE });
  }
  return minutes;
}

const WEEKDAY_INDEX: Record<string, number> = { Mon: 0, Tue: 1, Wed: 2, Thu: 3, Fri: 4, Sat: 5, Sun: 6 };

/** Monday 00:00 to next Monday 00:00 (exclusive) of the current week in `tz`, as UTC ISO strings. */
function weekRange(tz: string): { start: string; end: string } {
  const now = new Date();
  const weekday = new Intl.DateTimeFormat('en-US', { timeZone: tz, weekday: 'short' }).format(now);
  // Step back from midday so a DST shift can't move us into a neighbouring day.
  const midday = new Date(new Date(dayRange(now, tz).start).getTime() + 12 * 3600e3);
  const monday = new Date(midday.getTime() - WEEKDAY_INDEX[weekday] * 86400e3);
  return dayRange(monday, tz, 7);
}

type GraphTime = { dateTime: string; timeZone?: string };

const clock = (d: Date, tz: string) => d.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', timeZone: tz });

/** Minutes since midnight of the wall-clock time of `d` in `tz`. */
function minuteOfDay(d: Date, tz: string): number {
  const [h, m] = toGraphDateTime(d, tz).slice(11, 16).split(':').map(Number);
  return h * 60 + m;
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

async function inboxUnreadCount(client: GraphClient): Promise<number> {
  const res = await client.request('/me/mailFolders/inbox', { queryParams: { '$select': 'unreadItemCount' } });
  return ((res.data as Record<string, unknown>).unreadItemCount as number) ?? 0;
}

export function registerWorkflowCommands(program: Command, getClient: () => GraphClient): void {
  const wf = program.command('workflow').alias('wf').description('Cross-service productivity workflows');

  // ── standup-report ──────────────────────────────────────────────
  wf
    .command('standup')
    .description("Today's standup: calendar events + open tasks + unread email count")
    .option('--timezone <tz>', 'IANA time zone (default: MS365_CLI_TIMEZONE or system zone)')
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const tz = resolveTimeZone(opts.timezone);
      const { start, end } = dayRange(new Date(), tz);

      // Fetch in parallel: today's events and the inbox unread count
      const [eventsRes, unreadCount] = await Promise.all([
        client.request('/me/calendarView', {
          headers: preferTimeZone(tz),
          queryParams: {
            startDateTime: start,
            endDateTime: end,
            '$select': 'subject,start,end,location,isAllDay,organizer',
            '$orderby': 'start/dateTime',
            '$top': '50',
          },
        }),
        inboxUnreadCount(client),
      ]);
      if (client.isDryRun()) return previewReport(opts.output);

      const events = ((eventsRes.data as Record<string, unknown>).value as Array<Record<string, unknown>>) || [];

      if (opts.output === 'json') {
        console.log(JSON.stringify({ events, unreadEmails: unreadCount }, null, 2));
        return;
      }

      // Human-readable standup report
      const dateStr = new Date().toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric', timeZone: tz });

      console.log(`Standup Report — ${dateStr}`);
      console.log('='.repeat(50));

      // Calendar
      console.log(`\nCalendar (${events.length} events today):`);
      if (events.length === 0) {
        console.log('  No meetings today.');
      } else {
        for (const ev of events) {
          const startTime = clock(parseGraphDateTime(ev.start as GraphTime, tz), tz);
          const endTime = clock(parseGraphDateTime(ev.end as GraphTime, tz), tz);
          const loc = (ev.location as Record<string, unknown>)?.displayName;
          const locStr = loc ? ` @ ${loc}` : '';
          if (ev.isAllDay) {
            console.log(`  [all day] ${ev.subject}${locStr}`);
          } else {
            console.log(`  ${startTime}–${endTime}  ${ev.subject}${locStr}`);
          }
        }
      }

      // Unread emails
      console.log(`\nUnread Emails: ${unreadCount}`);

      console.log('');
    }));

  // ── meeting-prep ────────────────────────────────────────────────
  wf
    .command('meeting-prep')
    .description('Get details for the next upcoming meeting: attendees, agenda, location')
    .option('--event-id <id>', 'Specific event ID (default: next upcoming event)')
    .option('--timezone <tz>', 'IANA time zone (default: MS365_CLI_TIMEZONE or system zone)')
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const tz = resolveTimeZone(opts.timezone);

      let event: Record<string, unknown>;

      if (opts.eventId) {
        const res = await client.request(`/me/events/${encodeURIComponent(opts.eventId)}`, {
          headers: preferTimeZone(tz),
          queryParams: {
            '$select': 'subject,start,end,location,body,attendees,organizer,onlineMeeting,isAllDay',
          },
        });
        if (client.isPreview(res)) return previewReport(opts.output);
        event = res.data as Record<string, unknown>;
      } else {
        // Find the next upcoming real meeting: not all-day, not already started, not declined
        const now = new Date();
        const tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
        const res = await client.request('/me/calendarView', {
          headers: preferTimeZone(tz),
          queryParams: {
            startDateTime: now.toISOString(),
            endDateTime: tomorrow.toISOString(),
            '$select': 'subject,start,end,location,body,attendees,organizer,onlineMeeting,isAllDay,responseStatus',
            '$orderby': 'start/dateTime',
            '$top': '50',
          },
        });
        if (client.isPreview(res)) return previewReport(opts.output);
        const events = ((res.data as Record<string, unknown>).value as Array<Record<string, unknown>>) || [];
        const next = events.find(ev =>
          !ev.isAllDay
          && (ev.responseStatus as Record<string, unknown> | undefined)?.response !== 'declined'
          && parseGraphDateTime(ev.start as GraphTime, tz) >= now);
        if (!next) {
          info('No upcoming meetings in the next 24 hours.');
          return;
        }
        event = next;
      }

      if (opts.output === 'json') {
        console.log(formatOutput(event, 'json'));
        return;
      }

      // Human-readable meeting prep
      const startTime = parseGraphDateTime(event.start as GraphTime, tz).toLocaleString('en-US', {
        weekday: 'short', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit', timeZone: tz,
      });
      const endTime = clock(parseGraphDateTime(event.end as GraphTime, tz), tz);
      const loc = (event.location as Record<string, unknown>)?.displayName;
      const organizer = (event.organizer as Record<string, unknown>)?.emailAddress as Record<string, unknown>;
      const attendees = (event.attendees as Array<Record<string, unknown>>) || [];
      const body = event.body as Record<string, unknown>;
      const onlineMeeting = event.onlineMeeting as Record<string, unknown>;

      console.log(`Meeting Prep`);
      console.log('='.repeat(50));
      console.log(`Subject:   ${event.subject}`);
      console.log(`When:      ${startTime} – ${endTime}`);
      if (loc) console.log(`Where:     ${loc}`);
      if (onlineMeeting?.joinUrl) console.log(`Join:      ${onlineMeeting.joinUrl}`);
      if (organizer) console.log(`Organizer: ${organizer.name || ''} <${organizer.address || ''}>`);

      if (attendees.length > 0) {
        console.log(`\nAttendees (${attendees.length}):`);
        for (const att of attendees) {
          const email = att.emailAddress as Record<string, unknown>;
          const status = (att.status as Record<string, unknown>)?.response || 'none';
          const type = att.type || 'required';
          console.log(`  ${email?.name || ''} <${email?.address || ''}> [${type}, ${status}]`);
        }
      }

      if (body?.content) {
        let content = body.content as string;
        if (body.contentType === 'html') {
          content = content.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/\n\s*\n/g, '\n').trim();
        }
        if (content.trim()) {
          console.log(`\nAgenda/Notes:`);
          console.log(content.trim());
        }
      }

      console.log('');
    }));

  // ── email-to-task ───────────────────────────────────────────────
  wf
    .command('email-to-task <messageId>')
    .description('Convert an email into a To Do task')
    .option('--list <listId>', 'To Do list ID (default: first list)')
    .option('--due <date>', 'Due date (YYYY-MM-DD)')
    .addOption(new Option('--importance <level>', 'Importance').choices(['low', 'normal', 'high']).default('normal'))
    .option('--timezone <tz>', 'IANA time zone (default: MS365_CLI_TIMEZONE or system zone)')
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      const tz = resolveTimeZone(opts.timezone);
      // Validate before any request so a bad date exits 2 without touching Graph (V-13)
      const due = opts.due ? parseDate(opts.due) : undefined;

      // Fetch the email; under dry-run its fields are only known to the real run (DR-11)
      const msgRes = await client.request(`/me/messages/${encodeURIComponent(messageId)}`, {
        queryParams: { '$select': 'subject,from,bodyPreview,webLink,receivedDateTime' },
      });
      const msg: Record<string, unknown> = client.isPreview(msgRes)
        ? { subject: fromGet(), from: { emailAddress: { name: fromGet(), address: fromGet() } }, receivedDateTime: fromGet(), bodyPreview: fromGet(), webLink: fromGet() }
        : msgRes.data as Record<string, unknown>;
      const from = (msg.from as Record<string, unknown>)?.emailAddress as Record<string, unknown>;

      // Determine target list
      let listId = opts.list;
      if (!listId) {
        const listsRes = await client.request('/me/todo/lists', {
          queryParams: { '$top': '1' },
        });
        if (client.isPreview(listsRes)) {
          listId = FIRST_LIST;
        } else {
          const lists = ((listsRes.data as Record<string, unknown>).value as Array<Record<string, unknown>>) || [];
          if (lists.length === 0) {
            throw new CliError('No To Do lists found. Create one first.', { code: 'not_found', exitCode: ExitCode.NOT_FOUND });
          }
          listId = lists[0].id as string;
        }
      }

      // Create the task
      const task: Record<string, unknown> = {
        title: `[Email] ${msg.subject || '(no subject)'}`,
        importance: opts.importance,
        body: {
          contentType: 'text',
          content: [
            `From: ${from?.name || ''} <${from?.address || ''}>`,
            `Date: ${msg.receivedDateTime}`,
            `Preview: ${(msg.bodyPreview as string || '').substring(0, 200)}`,
            msg.webLink ? `Link: ${msg.webLink}` : '',
          ].filter(Boolean).join('\n'),
        },
      };

      if (due) {
        task.dueDateTime = { dateTime: due + 'T00:00:00', timeZone: tz };
      }

      const taskRes = await client.request(`/me/todo/lists/${encodeURIComponent(listId)}/tasks`, {
        method: 'POST',
        body: JSON.stringify(task),
      });

      // Prose only for -o text, and as a status line; every other format gets the data (V-21, DR-1)
      if (opts.output === 'text') {
        const created = taskRes.data as Record<string, unknown>;
        success(`Task created: "${created.title}"`);
        if (created.id) success(`ID: ${created.id}`);
      } else {
        console.log(formatOutput(taskRes.data, opts.output));
      }
    }));

  // ── weekly-digest ───────────────────────────────────────────────
  wf
    .command('digest')
    .description('Weekly digest: meetings this week, unread emails, tasks due')
    .option('--timezone <tz>', 'IANA time zone (default: MS365_CLI_TIMEZONE or system zone)')
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const tz = resolveTimeZone(opts.timezone);
      const { start, end } = weekRange(tz);

      // Fetch in parallel
      const [eventsRes, unreadCount, todoListsRes] = await Promise.all([
        client.requestAllPages('/me/calendarView', {
          headers: preferTimeZone(tz),
          queryParams: {
            startDateTime: start,
            endDateTime: end,
            '$select': 'subject,start,end,isAllDay',
            '$orderby': 'start/dateTime',
            '$top': '100',
          },
        }),
        inboxUnreadCount(client),
        client.request('/me/todo/lists', {
          queryParams: { '$top': '10' },
        }),
      ]);

      const events = ((eventsRes.data as Record<string, unknown>).value as Array<Record<string, unknown>>) || [];
      const todoLists = ((todoListsRes.data as Record<string, unknown>).value as Array<Record<string, unknown>>) || [];

      // Count tasks due this week across all lists
      let tasksDueCount = 0;
      let overdueTasks: Array<{ title: string; due: string }> = [];
      for (const list of todoLists.slice(0, 5)) {
        const tasksRes = await client.request(
          `/me/todo/lists/${encodeURIComponent(list.id as string)}/tasks`,
          {
            queryParams: {
              '$filter': `status ne 'completed'`,
              '$select': 'title,dueDateTime,status',
              '$top': '50',
            },
          }
        );
        const tasks = ((tasksRes.data as Record<string, unknown>).value as Array<Record<string, unknown>>) || [];
        for (const t of tasks) {
          const due = t.dueDateTime as GraphTime | undefined;
          if (due?.dateTime) {
            const dueDate = parseGraphDateTime(due, 'UTC');
            if (dueDate < new Date(end)) {
              tasksDueCount++;
              if (dueDate < new Date()) {
                overdueTasks.push({
                  title: t.title as string,
                  due: dueDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric', timeZone: tz }),
                });
              }
            }
          }
        }
      }

      if (client.isDryRun()) return previewReport(opts.output);

      if (opts.output === 'json') {
        console.log(JSON.stringify({
          week: { start, end },
          meetings: events.length,
          unreadEmails: unreadCount,
          tasksDueThisWeek: tasksDueCount,
          overdueTasks,
        }, null, 2));
        return;
      }

      // Group events by day
      const eventsByDay: Record<string, Array<Record<string, unknown>>> = {};
      for (const ev of events) {
        const day = parseGraphDateTime(ev.start as GraphTime, tz).toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', timeZone: tz });
        if (!eventsByDay[day]) eventsByDay[day] = [];
        eventsByDay[day].push(ev);
      }

      const weekStart = new Date(start).toLocaleDateString('en-US', { month: 'short', day: 'numeric', timeZone: tz });
      const weekEnd = new Date(new Date(end).getTime() - 1).toLocaleDateString('en-US', { month: 'short', day: 'numeric', timeZone: tz });

      console.log(`Weekly Digest — ${weekStart} to ${weekEnd}`);
      console.log('='.repeat(50));
      console.log(`Meetings:     ${events.length}`);
      console.log(`Unread Email: ${unreadCount}`);
      console.log(`Tasks Due:    ${tasksDueCount}`);
      if (overdueTasks.length > 0) {
        console.log(`Overdue:      ${overdueTasks.length}`);
      }

      if (Object.keys(eventsByDay).length > 0) {
        console.log('\nCalendar:');
        for (const [day, dayEvents] of Object.entries(eventsByDay)) {
          console.log(`  ${day}:`);
          for (const ev of dayEvents) {
            if (ev.isAllDay) {
              console.log(`    [all day] ${ev.subject}`);
            } else {
              const st = clock(parseGraphDateTime(ev.start as GraphTime, tz), tz);
              const et = clock(parseGraphDateTime(ev.end as GraphTime, tz), tz);
              console.log(`    ${st}–${et}  ${ev.subject}`);
            }
          }
        }
      }

      if (overdueTasks.length > 0) {
        console.log('\nOverdue Tasks:');
        for (const t of overdueTasks) {
          console.log(`  - ${t.title} (due ${t.due})`);
        }
      }

      console.log('');
    }));

  // ── file-announce ───────────────────────────────────────────────
  wf
    .command('file-announce <driveId> <itemId>')
    .description('Get file info and post announcement to a Teams channel (requires --org-mode)')
    .requiredOption('--team <teamId>', 'Team ID')
    .requiredOption('--channel <channelId>', 'Channel ID')
    .option('--message <text>', 'Custom announcement message')
    .action(handleErrors(async (driveId: string, itemId: string, opts) => {
      const client = getClient();

      // Get file metadata
      const fileRes = await client.request(
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`,
        { queryParams: { '$select': 'name,webUrl,size,lastModifiedDateTime,createdBy' } }
      );
      const file = fileRes.data as Record<string, unknown>;
      // Under dry-run the file's fields are only known to the real run (DR-11)
      const preview = client.isPreview(fileRes);
      const field = (v: unknown) => preview ? fromGet() : escapeHtml(String(v ?? ''));

      const sizeKb = preview ? fromGet() : String(Math.round((file.size as number || 0) / 1024));
      const name = field(file.name);
      const intro = opts.message ? escapeHtml(opts.message) : `New file shared: <b>${name}</b>`;
      const body = [
        intro,
        '',
        `File: <a href="${field(file.webUrl)}">${name}</a>`,
        `Size: ${sizeKb} KB`,
        `Modified: ${field(file.lastModifiedDateTime)}`,
      ].join('<br>');

      // Post to channel
      await client.request(
        `/teams/${encodeURIComponent(opts.team)}/channels/${encodeURIComponent(opts.channel)}/messages`,
        {
          method: 'POST',
          body: JSON.stringify({
            body: { contentType: 'html', content: body },
          }),
        }
      );

      success(`Announced "${file.name}" to channel.`);
    }));

  // ── focus-time ──────────────────────────────────────────────────
  wf
    .command('focus-time')
    .description('Find gaps in your calendar today for focus time')
    .option('--work-start <time>', 'Work day start, 24h hour or HH:MM', '9')
    .option('--work-end <time>', 'Work day end, 24h hour or HH:MM', '17')
    .option('--timezone <tz>', 'IANA time zone (default: MS365_CLI_TIMEZONE or system zone)')
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const tz = resolveTimeZone(opts.timezone);
      // Validate before the read so a bad bound exits 2, also under dry-run (V-20)
      const workStartMin = workMinute('--work-start', opts.workStart);
      const endMinute = workMinute('--work-end', opts.workEnd);
      if (workStartMin >= endMinute) {
        throw new CliError(`--work-start (${opts.workStart}) must be before --work-end (${opts.workEnd}).`, { code: 'usage', exitCode: ExitCode.USAGE });
      }
      const now = new Date();
      const { start, end } = dayRange(now, tz);
      const dayStart = new Date(start);
      const dayEnd = new Date(end);

      const eventsRes = await client.request('/me/calendarView', {
        headers: preferTimeZone(tz),
        queryParams: {
          startDateTime: start,
          endDateTime: end,
          '$select': 'subject,start,end,isAllDay,showAs',
          '$orderby': 'start/dateTime',
          '$top': '50',
        },
      });

      if (client.isPreview(eventsRes)) return previewReport(opts.output);

      const events = ((eventsRes.data as Record<string, unknown>).value as Array<Record<string, unknown>>) || [];

      // Build busy intervals (skip all-day and free events)
      const busy: Array<{ start: number; end: number; subject: string }> = [];
      for (const ev of events) {
        if (ev.isAllDay) continue;
        if (ev.showAs === 'free') continue;
        const s = parseGraphDateTime(ev.start as GraphTime, tz);
        const e = parseGraphDateTime(ev.end as GraphTime, tz);
        // Events that cross midnight are clamped to today's bounds
        let startMin = s <= dayStart ? 0 : minuteOfDay(s, tz);
        let endMin = e >= dayEnd ? 24 * 60 : minuteOfDay(e, tz);
        // Clamp to work window
        startMin = Math.max(startMin, workStartMin);
        endMin = Math.min(endMin, endMinute);
        if (startMin >= endMin) continue;
        busy.push({ start: startMin, end: endMin, subject: ev.subject as string });
      }

      // Sort by start time
      busy.sort((a, b) => a.start - b.start);

      // Merge overlapping intervals then find gaps
      const merged: typeof busy = [];
      for (const block of busy) {
        const last = merged[merged.length - 1];
        if (last && block.start <= last.end) {
          last.end = Math.max(last.end, block.end);
        } else {
          merged.push({ ...block });
        }
      }

      const gaps: Array<{ start: number; end: number; minutes: number }> = [];
      let cursor = Math.max(workStartMin, minuteOfDay(now, tz));

      for (const block of merged) {
        if (block.start > cursor) {
          const gapEnd = Math.min(block.start, endMinute);
          if (gapEnd - cursor >= 15) {
            gaps.push({ start: cursor, end: gapEnd, minutes: gapEnd - cursor });
          }
        }
        cursor = Math.max(cursor, block.end);
      }
      if (cursor < endMinute) {
        gaps.push({ start: cursor, end: endMinute, minutes: endMinute - cursor });
      }

      if (opts.output === 'json') {
        console.log(JSON.stringify({
          gaps: gaps.map(g => ({
            start: `${Math.floor(g.start / 60)}:${String(g.start % 60).padStart(2, '0')}`,
            end: `${Math.floor(g.end / 60)}:${String(g.end % 60).padStart(2, '0')}`,
            minutes: g.minutes,
          })),
          busyBlocks: busy.length,
          totalFocusMinutes: gaps.reduce((s, g) => s + g.minutes, 0),
        }, null, 2));
        return;
      }

      const fmtTime = (mins: number) => {
        const h = Math.floor(mins / 60);
        const m = mins % 60;
        const period = h >= 12 ? 'PM' : 'AM';
        const h12 = h > 12 ? h - 12 : (h === 0 ? 12 : h);
        return `${h12}:${String(m).padStart(2, '0')} ${period}`;
      };

      const totalFocus = gaps.reduce((s, g) => s + g.minutes, 0);

      console.log(`Focus Time — ${now.toLocaleDateString('en-US', { weekday: 'long', month: 'short', day: 'numeric', timeZone: tz })}`);
      console.log('='.repeat(50));
      console.log(`Meetings: ${busy.length}  |  Available focus: ${Math.floor(totalFocus / 60)}h ${totalFocus % 60}m\n`);

      if (gaps.length === 0) {
        console.log('No focus time available today. Your calendar is fully booked.');
      } else {
        for (const gap of gaps) {
          const dur = gap.minutes >= 60 ? `${Math.floor(gap.minutes / 60)}h ${gap.minutes % 60}m` : `${gap.minutes}m`;
          console.log(`  ${fmtTime(gap.start)} – ${fmtTime(gap.end)}  (${dur})`);
        }
      }

      if (busy.length > 0) {
        console.log('\nMeetings blocking time:');
        for (const b of busy) {
          console.log(`  ${fmtTime(b.start)} – ${fmtTime(b.end)}  ${b.subject}`);
        }
      }

      console.log('');
    }));
}
