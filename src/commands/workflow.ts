import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { handleErrors } from '../errors.js';
import { parseDate, truncate, outputOption } from '../utils.js';

function todayRange(): { start: string; end: string } {
  const now = new Date();
  const start = new Date(now);
  start.setUTCHours(0, 0, 0, 0);
  const end = new Date(now);
  end.setUTCHours(23, 59, 59, 999);
  return {
    start: start.toISOString(),
    end: end.toISOString(),
  };
}

function weekRange(): { start: string; end: string } {
  const now = new Date();
  const dayOfWeek = now.getDay();
  const monday = new Date(now);
  monday.setDate(now.getDate() - (dayOfWeek === 0 ? 6 : dayOfWeek - 1));
  monday.setHours(0, 0, 0, 0);
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  sunday.setHours(23, 59, 59, 999);
  return { start: monday.toISOString(), end: sunday.toISOString() };
}

export function registerWorkflowCommands(program: Command, getClient: () => GraphClient): void {
  const wf = program.command('workflow').alias('wf').description('Cross-service productivity workflows');

  // ── standup-report ──────────────────────────────────────────────
  wf
    .command('standup')
    .description("Today's standup: calendar events + open tasks + unread email count")
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const { start, end } = todayRange();

      // Fetch in parallel: today's events, my planner tasks, todo tasks, unread count
      const [eventsRes, unreadRes] = await Promise.all([
        client.request('/me/calendarView', {
          queryParams: {
            startDateTime: start,
            endDateTime: end,
            '$select': 'subject,start,end,location,isAllDay,organizer',
            '$orderby': 'start/dateTime',
            '$top': '50',
          },
        }),
        client.request('/me/messages', {
          queryParams: {
            '$filter': 'isRead eq false',
            '$select': 'id',
            '$top': '1',
            '$count': 'true',
          },
        }),
      ]);

      const events = ((eventsRes.data as Record<string, unknown>).value as Array<Record<string, unknown>>) || [];
      const unreadData = unreadRes.data as Record<string, unknown>;
      // Graph doesn't always return @odata.count, estimate from value array
      const unreadCount = (unreadData.value as unknown[])?.length ?? 0;
      const hasMoreUnread = unreadRes.nextLink ? '+' : '';

      if (opts.output === 'json') {
        console.log(JSON.stringify({ events, unreadEmails: `${unreadCount}${hasMoreUnread}` }, null, 2));
        return;
      }

      // Human-readable standup report
      const now = new Date();
      const dateStr = now.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });

      console.log(`Standup Report — ${dateStr}`);
      console.log('='.repeat(50));

      // Calendar
      console.log(`\nCalendar (${events.length} events today):`);
      if (events.length === 0) {
        console.log('  No meetings today.');
      } else {
        for (const ev of events) {
          const s = ev.start as Record<string, unknown>;
          const e = ev.end as Record<string, unknown>;
          const startTime = new Date(s.dateTime as string).toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
          const endTime = new Date(e.dateTime as string).toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
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
      console.log(`\nUnread Emails: ${unreadCount}${hasMoreUnread}`);

      console.log('');
    }));

  // ── meeting-prep ────────────────────────────────────────────────
  wf
    .command('meeting-prep')
    .description('Get details for the next upcoming meeting: attendees, agenda, location')
    .option('--event-id <id>', 'Specific event ID (default: next upcoming event)')
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts) => {
      const client = getClient();

      let event: Record<string, unknown>;

      if (opts.eventId) {
        const res = await client.request(`/me/events/${encodeURIComponent(opts.eventId)}`, {
          queryParams: {
            '$select': 'subject,start,end,location,body,attendees,organizer,onlineMeeting,isAllDay',
          },
        });
        event = res.data as Record<string, unknown>;
      } else {
        // Find the next upcoming event
        const now = new Date().toISOString();
        const tomorrow = new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString();
        const res = await client.request('/me/calendarView', {
          queryParams: {
            startDateTime: now,
            endDateTime: tomorrow,
            '$select': 'subject,start,end,location,body,attendees,organizer,onlineMeeting,isAllDay',
            '$orderby': 'start/dateTime',
            '$top': '1',
          },
        });
        const events = ((res.data as Record<string, unknown>).value as Array<Record<string, unknown>>) || [];
        if (events.length === 0) {
          console.log('No upcoming meetings in the next 24 hours.');
          return;
        }
        event = events[0];
      }

      if (opts.output === 'json') {
        console.log(formatOutput(event, 'json'));
        return;
      }

      // Human-readable meeting prep
      const start = event.start as Record<string, unknown>;
      const end = event.end as Record<string, unknown>;
      const startTime = new Date(start.dateTime as string).toLocaleString('en-US', {
        weekday: 'short', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit',
      });
      const endTime = new Date(end.dateTime as string).toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
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
    .option('--importance <level>', 'Importance: low, normal, high', 'normal')
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();

      // Fetch the email
      const msgRes = await client.request(`/me/messages/${encodeURIComponent(messageId)}`, {
        queryParams: { '$select': 'subject,from,bodyPreview,webLink,receivedDateTime' },
      });
      const msg = msgRes.data as Record<string, unknown>;
      const from = (msg.from as Record<string, unknown>)?.emailAddress as Record<string, unknown>;

      // Determine target list
      let listId = opts.list;
      if (!listId) {
        const listsRes = await client.request('/me/todo/lists', {
          queryParams: { '$top': '1' },
        });
        const lists = ((listsRes.data as Record<string, unknown>).value as Array<Record<string, unknown>>) || [];
        if (lists.length === 0) {
          throw new Error('No To Do lists found. Create one first.');
        }
        listId = lists[0].id as string;
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

      if (opts.due) {
        task.dueDateTime = { dateTime: parseDate(opts.due) + 'T00:00:00', timeZone: 'UTC' };
      }

      const taskRes = await client.request(`/me/todo/lists/${encodeURIComponent(listId)}/tasks`, {
        method: 'POST',
        body: JSON.stringify(task),
      });

      if (opts.output === 'json') {
        console.log(formatOutput(taskRes.data, 'json'));
      } else {
        const created = taskRes.data as Record<string, unknown>;
        console.log(`Task created: "${created.title}"`);
        if (created.id) console.log(`ID: ${created.id}`);
      }
    }));

  // ── weekly-digest ───────────────────────────────────────────────
  wf
    .command('digest')
    .description('Weekly digest: meetings this week, unread emails, tasks due')
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const { start, end } = weekRange();

      // Fetch in parallel
      const [eventsRes, unreadRes, todoListsRes] = await Promise.all([
        client.requestAllPages('/me/calendarView', {
          queryParams: {
            startDateTime: start,
            endDateTime: end,
            '$select': 'subject,start,end,isAllDay',
            '$orderby': 'start/dateTime',
            '$top': '100',
          },
        }),
        client.request('/me/messages', {
          queryParams: {
            '$filter': 'isRead eq false',
            '$select': 'id',
            '$top': '1',
          },
        }),
        client.request('/me/todo/lists', {
          queryParams: { '$top': '10' },
        }),
      ]);

      const events = ((eventsRes.data as Record<string, unknown>).value as Array<Record<string, unknown>>) || [];
      const unreadCount = ((unreadRes.data as Record<string, unknown>).value as unknown[])?.length || 0;
      const hasMoreUnread = unreadRes.nextLink ? '+' : '';
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
          const due = t.dueDateTime as Record<string, unknown>;
          if (due?.dateTime) {
            const dueDate = new Date(due.dateTime as string);
            const endDate = new Date(end);
            if (dueDate <= endDate) {
              tasksDueCount++;
              if (dueDate < new Date()) {
                overdueTasks.push({
                  title: t.title as string,
                  due: dueDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }),
                });
              }
            }
          }
        }
      }

      if (opts.output === 'json') {
        console.log(JSON.stringify({
          week: { start, end },
          meetings: events.length,
          unreadEmails: `${unreadCount}${hasMoreUnread}`,
          tasksDueThisWeek: tasksDueCount,
          overdueTasks,
        }, null, 2));
        return;
      }

      // Group events by day
      const eventsByDay: Record<string, Array<Record<string, unknown>>> = {};
      for (const ev of events) {
        const s = ev.start as Record<string, unknown>;
        const day = new Date(s.dateTime as string).toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
        if (!eventsByDay[day]) eventsByDay[day] = [];
        eventsByDay[day].push(ev);
      }

      const weekStart = new Date(start).toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
      const weekEnd = new Date(end).toLocaleDateString('en-US', { month: 'short', day: 'numeric' });

      console.log(`Weekly Digest — ${weekStart} to ${weekEnd}`);
      console.log('='.repeat(50));
      console.log(`Meetings:     ${events.length}`);
      console.log(`Unread Email: ${unreadCount}${hasMoreUnread}`);
      console.log(`Tasks Due:    ${tasksDueCount}`);
      if (overdueTasks.length > 0) {
        console.log(`Overdue:      ${overdueTasks.length}`);
      }

      if (Object.keys(eventsByDay).length > 0) {
        console.log('\nCalendar:');
        for (const [day, dayEvents] of Object.entries(eventsByDay)) {
          console.log(`  ${day}:`);
          for (const ev of dayEvents) {
            const s = ev.start as Record<string, unknown>;
            const e = ev.end as Record<string, unknown>;
            if (ev.isAllDay) {
              console.log(`    [all day] ${ev.subject}`);
            } else {
              const st = new Date(s.dateTime as string).toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
              const et = new Date(e.dateTime as string).toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
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

      const sizeKb = Math.round((file.size as number || 0) / 1024);
      const customMsg = opts.message || `New file shared: **${file.name}**`;
      const body = [
        customMsg,
        '',
        `File: [${file.name}](${file.webUrl})`,
        `Size: ${sizeKb} KB`,
        `Modified: ${file.lastModifiedDateTime}`,
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

      console.log(`Announced "${file.name}" to channel.`);
    }));

  // ── focus-time ──────────────────────────────────────────────────
  wf
    .command('focus-time')
    .description('Find gaps in your calendar today for focus time')
    .option('--work-start <hour>', 'Work day start hour (24h)', '9')
    .option('--work-end <hour>', 'Work day end hour (24h)', '17')
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const { start, end } = todayRange();

      const eventsRes = await client.request('/me/calendarView', {
        queryParams: {
          startDateTime: start,
          endDateTime: end,
          '$select': 'subject,start,end,isAllDay,showAs',
          '$orderby': 'start/dateTime',
          '$top': '50',
        },
      });

      const events = ((eventsRes.data as Record<string, unknown>).value as Array<Record<string, unknown>>) || [];

      const workStart = parseInt(opts.workStart);
      const workEnd = parseInt(opts.workEnd);
      const now = new Date();
      const today = new Date(now);

      // Build busy intervals (skip all-day and free events)
      const workStartMin = workStart * 60;
      const endMinute = workEnd * 60;
      const busy: Array<{ start: number; end: number; subject: string }> = [];
      for (const ev of events) {
        if (ev.isAllDay) continue;
        if (ev.showAs === 'free') continue;
        const s = new Date((ev.start as Record<string, unknown>).dateTime as string);
        const e = new Date((ev.end as Record<string, unknown>).dateTime as string);
        let startMin = s.getHours() * 60 + s.getMinutes();
        let endMin = e.getHours() * 60 + e.getMinutes();
        // Handle midnight-crossing meetings: clamp to work window end
        if (endMin <= startMin) endMin = endMinute;
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
      let cursor = Math.max(workStartMin, now.getHours() * 60 + now.getMinutes());

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

      console.log(`Focus Time — ${today.toLocaleDateString('en-US', { weekday: 'long', month: 'short', day: 'numeric' })}`);
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
