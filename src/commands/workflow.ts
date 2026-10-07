import { Command, InvalidArgumentError, Option } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { handleErrors } from '../errors.js';
import { outputOption, info, done } from '../utils.js';
import { calendarDate, collectionItems, preferTimezone } from '../graph-helpers.js';

type Item = Record<string, unknown>;

/** A Graph `dateTimeTimeZone`: offset-less wall-clock `dateTime` in `timeZone`. */
interface DateTimeTimeZone {
  dateTime?: unknown;
  timeZone?: unknown;
}

interface ReportOptions {
  output: string;
}

const UTC_ZONES: Record<string, true> = { utc: true, 'etc/utc': true, gmt: true, 'etc/gmt': true, z: true, 'coordinated universal time': true };

/** Offset of `timeZone` from UTC at `instant`, in ms (wall clock in the zone minus UTC). */
function zoneOffset(instant: number, timeZone: string): number {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone, hourCycle: 'h23', year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric', second: 'numeric',
  }).formatToParts(instant);
  const part = (type: Intl.DateTimeFormatPartTypes) => Number(parts.find((p) => p.type === type)?.value);
  const wall = Date.UTC(part('year'), part('month') - 1, part('day'), part('hour'), part('minute'), part('second'));
  return wall - (instant - (instant % 1000));
}

/**
 * The instant a Graph `dateTimeTimeZone` denotes. Graph sends offset-less wall-clock times in
 * `timeZone` — the zone asked for with `Prefer: outlook.timezone`, UTC otherwise. A zone name the
 * runtime does not know (e.g. a Windows name) is read as local time, the zone the CLI asks for.
 */
export function eventInstant(value: unknown): Date {
  const { dateTime, timeZone } = (value ?? {}) as DateTimeTimeZone;
  const raw = typeof dateTime === 'string' ? dateTime : '';
  if (/(?:Z|[+-]\d{2}:?\d{2})$/i.test(raw)) return new Date(raw);
  const asUtc = Date.parse(`${raw}Z`);
  const zone = typeof timeZone === 'string' ? timeZone : 'UTC';
  if (Number.isNaN(asUtc) || UTC_ZONES[zone.toLowerCase()]) return new Date(asUtc);
  try {
    // Two passes so a DST change between the guess and the answer is accounted for.
    const guess = asUtc - zoneOffset(asUtc, zone);
    return new Date(asUtc - zoneOffset(guess, zone));
  } catch {
    return new Date(raw);
  }
}

/** Local calendar day an event starts on; all-day events carry a date, not an instant. */
function eventStartDay(ev: Item): Date {
  const start = ev.start as DateTimeTimeZone | undefined;
  if (ev.isAllDay && typeof start?.dateTime === 'string') {
    const [y, m, d] = start.dateTime.slice(0, 10).split('-').map(Number);
    return new Date(y, m - 1, d);
  }
  return eventInstant(start);
}

/** `YYYY-MM-DD` of `d` in the local zone. */
function localDate(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

/**
 * The local `YYYY-MM-DD` of a To Do `dueDateTime`. Graph stores a due date as midnight in the zone it
 * was set in (Outlook in Berlin: `2026-10-06T22:00:00` UTC for Oct 7; `todo --due`: `T00:00:00` UTC),
 * or returns local midnight when the Prefer zone is honoured. The instant is rounded to the nearest
 * local midnight, which is that day whenever the setting zone is within 12 hours of the local one.
 */
function dueDay(value: unknown): string | undefined {
  if (typeof (value as DateTimeTimeZone | undefined)?.dateTime !== 'string') return undefined;
  const ms = eventInstant(value).getTime();
  return Number.isNaN(ms) ? undefined : localDate(new Date(ms + 12 * 60 * 60 * 1000));
}

/** Local midnight of `now`'s day plus `days`. */
function localMidnight(now: Date, days = 0): Date {
  return new Date(now.getFullYear(), now.getMonth(), now.getDate() + days);
}

function values(data: unknown): Item[] {
  return (collectionItems(data) ?? []) as Item[];
}

/** Every event overlapping [start, end), with times in the local zone, ordered by start. */
async function calendarView(client: GraphClient, start: Date, end: Date, select: string): Promise<Item[]> {
  const res = await client.requestAllPages('/me/calendarView', {
    headers: preferTimezone(),
    queryParams: {
      startDateTime: start.toISOString(),
      endDateTime: end.toISOString(),
      '$select': select,
      '$orderby': 'start/dateTime',
      '$top': '100',
    },
  });
  return values(res.data);
}

/** Unread messages in the inbox (the folder's own counter, not a page length). */
async function inboxUnreadCount(client: GraphClient): Promise<number> {
  const res = await client.request('/me/mailFolders/inbox', { queryParams: { '$select': 'unreadItemCount' } });
  const count = (res.data as Item).unreadItemCount;
  return typeof count === 'number' ? count : 0;
}

/** `HH:MM` (or `HH`) on a 24h clock → minutes since midnight; `24:00` allowed as end of day. */
function clockMinutes(value: string): number {
  const m = /^(\d{1,2})(?::(\d{2}))?$/.exec(value);
  const hours = Number(m?.[1]);
  const minutes = Number(m?.[2] ?? 0);
  if (!m || hours > 24 || minutes > 59 || (hours === 24 && minutes > 0)) {
    throw new InvalidArgumentError('Expected a 24h time HH:MM, e.g. 09:00.');
  }
  return hours * 60 + minutes;
}

const time = (d: Date) => d.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
const HTML_ENTITIES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
const escapeHtml = (s: string) => s.replace(/[&<>"']/g, (c) => HTML_ENTITIES[c]);

export function registerWorkflowCommands(program: Command, getClient: () => GraphClient): void {
  const wf = program.command('workflow').alias('wf').description('Cross-service productivity workflows');

  // ── standup-report ──────────────────────────────────────────────
  wf
    .command('standup')
    .description("Today's standup: calendar events + inbox unread count")
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts: ReportOptions) => {
      const client = getClient();
      const now = new Date();
      const [events, unreadEmails] = await Promise.all([
        calendarView(client, localMidnight(now), localMidnight(now, 1), 'subject,start,end,location,isAllDay,organizer'),
        inboxUnreadCount(client),
      ]);

      if (opts.output !== 'text') {
        console.log(formatOutput({ events, unreadEmails }, opts.output as OutputFormat));
        return;
      }

      const dateStr = now.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
      console.log(`Standup Report — ${dateStr}`);
      console.log('='.repeat(50));

      console.log(`\nCalendar (${events.length} events today):`);
      if (events.length === 0) {
        console.log('  No meetings today.');
      } else {
        for (const ev of events) {
          const loc = (ev.location as Item | undefined)?.displayName;
          const locStr = loc ? ` @ ${loc}` : '';
          if (ev.isAllDay) {
            console.log(`  [all day] ${ev.subject}${locStr}`);
          } else {
            console.log(`  ${time(eventInstant(ev.start))}–${time(eventInstant(ev.end))}  ${ev.subject}${locStr}`);
          }
        }
      }

      console.log(`\nUnread Emails: ${unreadEmails}`);
      console.log('');
    }));

  // ── meeting-prep ────────────────────────────────────────────────
  wf
    .command('meeting-prep')
    .description('Get details for the next upcoming meeting: attendees, agenda, location')
    .option('--event-id <id>', 'Specific event ID (default: next meeting starting in the next 24 hours)')
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts: ReportOptions & { eventId?: string }) => {
      const client = getClient();
      const select = 'subject,start,end,location,body,attendees,organizer,onlineMeeting,isAllDay,isCancelled';

      let event: Item | undefined;
      if (opts.eventId) {
        const res = await client.request(`/me/events/${encodeURIComponent(opts.eventId)}`, {
          headers: preferTimezone(),
          queryParams: { '$select': select },
        });
        event = res.data as Item;
      } else {
        // The view also returns all-day and already-running events; the next meeting is the first timed one still ahead.
        const now = new Date();
        const events = await calendarView(client, now, new Date(now.getTime() + 24 * 60 * 60 * 1000), select);
        event = events.find((ev) => !ev.isAllDay && !ev.isCancelled && eventInstant(ev.start) >= now);
      }

      if (!event) {
        // Under --dry-run the view is only previewed, so "no meeting" would be a guess.
        if (!client.isDryRun) info('No upcoming meetings in the next 24 hours.');
        if (opts.output !== 'text') console.log(formatOutput(null, opts.output as OutputFormat));
        return;
      }

      if (opts.output !== 'text') {
        console.log(formatOutput(event, opts.output as OutputFormat));
        return;
      }

      const startTime = eventInstant(event.start).toLocaleString('en-US', {
        weekday: 'short', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit',
      });
      const endTime = time(eventInstant(event.end));
      const loc = (event.location as Item | undefined)?.displayName;
      const organizer = (event.organizer as Item | undefined)?.emailAddress as Item | undefined;
      const attendees = (event.attendees as Item[] | undefined) || [];
      const body = event.body as Item | undefined;
      const onlineMeeting = event.onlineMeeting as Item | undefined;

      console.log(`Meeting Prep`);
      console.log('='.repeat(50));
      console.log(`Subject:   ${event.subject}`);
      console.log(`When:      ${event.isAllDay ? `${eventStartDay(event).toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' })} (all day)` : `${startTime} – ${endTime}`}`);
      if (loc) console.log(`Where:     ${loc}`);
      if (onlineMeeting?.joinUrl) console.log(`Join:      ${onlineMeeting.joinUrl}`);
      if (organizer) console.log(`Organizer: ${organizer.name || ''} <${organizer.address || ''}>`);

      if (attendees.length > 0) {
        console.log(`\nAttendees (${attendees.length}):`);
        for (const att of attendees) {
          const email = att.emailAddress as Item | undefined;
          const status = (att.status as Item | undefined)?.response || 'none';
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
    .option('--list <listId>', 'To Do list ID (default: your default "Tasks" list)')
    .option('--due <date>', 'Due date (YYYY-MM-DD)')
    .addOption(new Option('--importance <level>', 'Importance').choices(['low', 'normal', 'high']).default('normal'))
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts: { list?: string; due?: string; importance: string; output: string }) => {
      const due = opts.due ? calendarDate(opts.due) : undefined;
      const client = getClient();

      const msgRes = await client.request(`/me/messages/${encodeURIComponent(messageId)}`, {
        queryParams: { '$select': 'subject,from,bodyPreview,webLink,receivedDateTime' },
      });
      const msg = msgRes.data as Item;
      const from = (msg.from as Item | undefined)?.emailAddress as Item | undefined;

      // Under --dry-run nothing is read: placeholders stand in for the default list and the subject.
      let listId = opts.list;
      if (!listId && client.isDryRun) listId = '<default-list>';
      if (!listId) {
        const listsRes = await client.request('/me/todo/lists', {
          queryParams: { '$filter': "wellknownListName eq 'defaultList'" },
        });
        const list = values(listsRes.data)[0];
        if (!list) throw new Error('No default To Do list found; pass --list <listId>.');
        listId = list.id as string;
      }

      const subject = client.isDryRun ? '<subject>' : msg.subject || '(no subject)';
      const task: Item = {
        title: `[Email] ${subject}`,
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
      if (due) task.dueDateTime = { dateTime: `${due}T00:00:00`, timeZone: 'UTC' };

      const taskRes = await client.request(`/me/todo/lists/${encodeURIComponent(listId)}/tasks`, {
        method: 'POST',
        body: JSON.stringify(task),
      });
      console.log(formatOutput(taskRes.data, opts.output as OutputFormat));
    }));

  // ── weekly-digest ───────────────────────────────────────────────
  wf
    .command('digest')
    .description('Weekly digest: meetings this week, inbox unread count, open tasks due this week and overdue (all To Do lists)')
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts: ReportOptions) => {
      const client = getClient();
      const now = new Date();
      const monday = localMidnight(now, -((now.getDay() + 6) % 7));
      const nextMonday = localMidnight(monday, 7);
      const sunday = localMidnight(monday, 6);

      const [events, unreadEmails, listsRes] = await Promise.all([
        calendarView(client, monday, nextMonday, 'subject,start,end,isAllDay'),
        inboxUnreadCount(client),
        client.requestAllPages('/me/todo/lists', { queryParams: { '$select': 'id,displayName' } }),
      ]);

      // To Do due dates are dates stored as midnight in the zone they were set in (see dueDay).
      const today = localDate(now);
      const weekEnd = localDate(sunday);
      let tasksDueThisWeek = 0;
      const overdueTasks: Array<{ title: string; due: string }> = [];
      for (const list of values(listsRes.data)) {
        const tasksRes = await client.requestAllPages(`/me/todo/lists/${encodeURIComponent(list.id as string)}/tasks`, {
          queryParams: {
            '$filter': `status ne 'completed'`,
            '$select': 'title,dueDateTime,status',
            '$top': '100',
          },
          headers: preferTimezone(),
        });
        for (const t of values(tasksRes.data)) {
          const due = dueDay(t.dueDateTime);
          if (!due) continue;
          if (due < today) overdueTasks.push({ title: t.title as string, due });
          else if (due <= weekEnd) tasksDueThisWeek++;
        }
      }

      if (opts.output !== 'text') {
        console.log(formatOutput({
          week: { start: localDate(monday), end: weekEnd },
          meetings: events.length,
          unreadEmails,
          tasksDueThisWeek,
          overdueTasks,
        }, opts.output as OutputFormat));
        return;
      }

      const eventsByDay = new Map<string, Item[]>();
      for (const ev of events) {
        const day = eventStartDay(ev).toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
        const dayEvents = eventsByDay.get(day);
        if (dayEvents) dayEvents.push(ev);
        else eventsByDay.set(day, [ev]);
      }

      const shortDate = (d: Date) => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
      console.log(`Weekly Digest — ${shortDate(monday)} to ${shortDate(sunday)}`);
      console.log('='.repeat(50));
      console.log(`Meetings:     ${events.length}`);
      console.log(`Unread Email: ${unreadEmails}`);
      console.log(`Tasks Due:    ${tasksDueThisWeek}`);
      if (overdueTasks.length > 0) {
        console.log(`Overdue:      ${overdueTasks.length}`);
      }

      if (eventsByDay.size > 0) {
        console.log('\nCalendar:');
        for (const [day, dayEvents] of eventsByDay) {
          console.log(`  ${day}:`);
          for (const ev of dayEvents) {
            if (ev.isAllDay) {
              console.log(`    [all day] ${ev.subject}`);
            } else {
              console.log(`    ${time(eventInstant(ev.start))}–${time(eventInstant(ev.end))}  ${ev.subject}`);
            }
          }
        }
      }

      if (overdueTasks.length > 0) {
        console.log('\nOverdue Tasks:');
        for (const t of overdueTasks) {
          const [y, m, d] = t.due.split('-').map(Number);
          console.log(`  - ${t.title} (due ${shortDate(new Date(y, m - 1, d))})`);
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
    .option('--message <text>', 'Custom announcement text (plain text)')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, opts: { team: string; channel: string; message?: string; output: string }) => {
      const client = getClient();

      const fileRes = await client.request(
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`,
        { queryParams: { '$select': 'name,webUrl,size,lastModifiedDateTime,createdBy' } }
      );
      const file = fileRes.data as Item;

      // The message body is HTML: every value from the file or the user is escaped.
      const name = escapeHtml(String(file.name ?? ''));
      const sizeKb = Math.round((file.size as number || 0) / 1024);
      const content = [
        opts.message ? escapeHtml(opts.message) : `New file shared: <b>${name}</b>`,
        '',
        `File: ${typeof file.webUrl === 'string' ? `<a href="${escapeHtml(file.webUrl)}">${name}</a>` : name}`,
        `Size: ${sizeKb} KB`,
        `Modified: ${escapeHtml(String(file.lastModifiedDateTime ?? ''))}`,
      ].join('<br>');

      const postRes = await client.request(
        `/teams/${encodeURIComponent(opts.team)}/channels/${encodeURIComponent(opts.channel)}/messages`,
        {
          method: 'POST',
          body: JSON.stringify({ body: { contentType: 'html', content } }),
        }
      );

      done(postRes, `Announced "${file.name}" to channel.`);
      console.log(formatOutput(postRes.data, opts.output as OutputFormat));
    }));

  // ── focus-time ──────────────────────────────────────────────────
  wf
    .command('focus-time')
    .description('Find gaps in your calendar today for focus time')
    .addOption(new Option('--work-start <HH:MM>', 'Work day start (24h)').argParser(clockMinutes).default(9 * 60, '09:00'))
    .addOption(new Option('--work-end <HH:MM>', 'Work day end (24h)').argParser(clockMinutes).default(17 * 60, '17:00'))
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts: ReportOptions & { workStart: number; workEnd: number }) => {
      const workStartMin = opts.workStart;
      const endMinute = opts.workEnd;
      if (endMinute <= workStartMin) throw new Error('--work-end must be later than --work-start.');

      const client = getClient();
      const now = new Date();
      const dayStart = localMidnight(now);
      const dayEnd = localMidnight(now, 1);
      const events = await calendarView(client, dayStart, dayEnd, 'subject,start,end,isAllDay,showAs');

      // Local wall-clock minute of today; events reaching into other days are clamped to today.
      const minuteOfDay = (d: Date) => (d < dayStart ? 0 : d >= dayEnd ? 24 * 60 : d.getHours() * 60 + d.getMinutes());

      // Build busy intervals (skip all-day and free events)
      const busy: Array<{ start: number; end: number; subject: string }> = [];
      for (const ev of events) {
        if (ev.isAllDay) continue;
        if (ev.showAs === 'free') continue;
        const startMin = Math.max(minuteOfDay(eventInstant(ev.start)), workStartMin);
        const endMin = Math.min(minuteOfDay(eventInstant(ev.end)), endMinute);
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
      let cursor = Math.max(workStartMin, minuteOfDay(now));

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

      if (opts.output !== 'text') {
        console.log(formatOutput({
          gaps: gaps.map(g => ({
            start: `${Math.floor(g.start / 60)}:${String(g.start % 60).padStart(2, '0')}`,
            end: `${Math.floor(g.end / 60)}:${String(g.end % 60).padStart(2, '0')}`,
            minutes: g.minutes,
          })),
          busyBlocks: busy.length,
          totalFocusMinutes: gaps.reduce((s, g) => s + g.minutes, 0),
        }, opts.output as OutputFormat));
        return;
      }

      const fmtTime = (mins: number) => {
        const h = Math.floor(mins / 60) % 24;
        const m = mins % 60;
        const period = h >= 12 ? 'PM' : 'AM';
        const h12 = h > 12 ? h - 12 : (h === 0 ? 12 : h);
        return `${h12}:${String(m).padStart(2, '0')} ${period}`;
      };

      const totalFocus = gaps.reduce((s, g) => s + g.minutes, 0);

      console.log(`Focus Time — ${now.toLocaleDateString('en-US', { weekday: 'long', month: 'short', day: 'numeric' })}`);
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
