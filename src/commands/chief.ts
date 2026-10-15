import { Command, InvalidArgumentError, Option } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption, resolveTimeZone, preferTimeZone, dayRange, parseGraphDateTime, toIsoWithOffset, truncate } from '../utils.js';
import { handleErrors } from '../errors.js';

/**
 * Chief-of-staff briefing for the signed-in person: their own calendar, mail, tasks and Teams in
 * one report, ordered by what needs attention. Everything is read with the person's own
 * delegated sign-in, so it never shows anything they couldn't already open themselves.
 * Message text from other people is marked untrusted and truncated.
 */

type Row = Record<string, unknown>;
const PREVIEW_CHARS = 200;

interface Section<T> { ok: boolean; items: T[]; error?: string }

/** Run one section; a failure in one source (say Teams without consent) never sinks the briefing. */
async function section<T>(fn: () => Promise<T[]>): Promise<Section<T>> {
  try {
    return { ok: true, items: await fn() };
  } catch (err) {
    return { ok: false, items: [], error: err instanceof Error ? err.message : String(err) };
  }
}

const list = (res: { data: unknown }): Row[] => ((res.data as { value?: Row[] }).value ?? []);
const untrusted = (preview: unknown) => ({ untrusted: true, preview: truncate(String(preview ?? '').replace(/\s+/g, ' ').trim(), PREVIEW_CHARS) });

export interface Briefing {
  generatedAt: string;
  timeZone: string;
  attention: string[];
  calendar: Section<Row>;
  mail: { unreadCount: number | null; important: Section<Row>; flagged: Section<Row> };
  tasks: Section<Row>;
  teams: Section<Row>;
}

/** Plain-language lines summarising what most needs attention, from the collected sections. */
export function attentionLines(b: Pick<Briefing, 'calendar' | 'mail' | 'tasks'>, now = new Date()): string[] {
  const lines: string[] = [];
  const upcoming = b.calendar.items.filter(e => !e.isAllDay && String(e.start) >= toIso(now));
  const next = upcoming[0];
  if (next) lines.push(`Next meeting: ${next.subject} at ${String(next.start).slice(11, 16)}`);
  const back = b.calendar.items.filter(e => e.overlapsPrevious).length;
  if (back) lines.push(`${back} back-to-back or overlapping meeting${back > 1 ? 's' : ''} today`);
  if (b.mail.important.items.length) lines.push(`${b.mail.important.items.length} unread high-importance email${b.mail.important.items.length > 1 ? 's' : ''}`);
  if (b.mail.flagged.items.length) lines.push(`${b.mail.flagged.items.length} flagged email${b.mail.flagged.items.length > 1 ? 's' : ''} waiting`);
  const overdue = b.tasks.items.filter(t => t.overdue).length;
  if (overdue) lines.push(`${overdue} overdue task${overdue > 1 ? 's' : ''}`);
  return lines;
}
const toIso = (d: Date) => d.toISOString().slice(0, 19);

export function registerChiefCommands(program: Command, getClient: () => GraphClient): void {
  const chief = program.command('chief').description('Chief-of-staff view of your own work: what needs your attention today');

  chief
    .command('brief')
    .description("Today's briefing: meetings, high-importance and flagged mail, overdue tasks, recent Teams chats (your own data)")
    .option('--timezone <tz>', 'IANA time zone (default: MS365_CLI_TIMEZONE or the local zone)')
    .addOption(new Option('--days <n>', 'Look-ahead for the calendar, in days').argParser(v => {
      const n = Number(v);
      if (!Number.isInteger(n) || n < 1 || n > 14) throw new InvalidArgumentError('must be 1-14.');
      return n;
    }).default(1))
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const tz = resolveTimeZone(opts.timezone);
      const now = new Date();
      const { start, end } = dayRange(now, tz, opts.days);

      const [calendar, important, flagged, tasks, teams, unread] = await Promise.all([
        section(async () => {
          const res = await client.request('/me/calendarView', {
            headers: preferTimeZone(tz),
            queryParams: { startDateTime: start, endDateTime: end, $select: 'subject,start,end,isAllDay,location,organizer,responseStatus,onlineMeeting', $orderby: 'start/dateTime', $top: '50' },
          });
          let prevEnd = 0;
          return list(res)
            .filter(e => (e.responseStatus as { response?: string } | undefined)?.response !== 'declined')
            .map(e => {
              const s = parseGraphDateTime(e.start as { dateTime: string; timeZone?: string }, tz);
              const en = parseGraphDateTime(e.end as { dateTime: string; timeZone?: string }, tz);
              const overlapsPrevious = !e.isAllDay && s.getTime() < prevEnd;
              if (!e.isAllDay) prevEnd = Math.max(prevEnd, en.getTime());
              return {
                subject: e.subject, start: toIsoWithOffset(s, tz), end: toIsoWithOffset(en, tz), isAllDay: e.isAllDay,
                organizer: (e.organizer as { emailAddress?: { name?: string } } | undefined)?.emailAddress?.name,
                joinUrl: (e.onlineMeeting as { joinUrl?: string } | null | undefined)?.joinUrl, overlapsPrevious,
              } as Row;
            });
        }),
        section(async () => list(await client.request('/me/mailFolders/inbox/messages', {
          queryParams: { $filter: "isRead eq false and importance eq 'high'", $select: 'subject,from,receivedDateTime,bodyPreview', $top: '10' },
        })).map(m => ({ subject: m.subject, from: (m.from as { emailAddress?: { address?: string } } | undefined)?.emailAddress?.address, receivedDateTime: m.receivedDateTime, ...untrusted(m.bodyPreview) }))),
        section(async () => list(await client.request('/me/messages', {
          queryParams: { $filter: "flag/flagStatus eq 'flagged'", $select: 'subject,from,receivedDateTime,bodyPreview', $top: '10' },
        })).map(m => ({ subject: m.subject, from: (m.from as { emailAddress?: { address?: string } } | undefined)?.emailAddress?.address, receivedDateTime: m.receivedDateTime, ...untrusted(m.bodyPreview) }))),
        section(async () => {
          const lists = list(await client.request('/me/todo/lists', { queryParams: { $top: '10' } }));
          const out: Row[] = [];
          for (const l of lists.slice(0, 5)) {
            const tasks = list(await client.request(`/me/todo/lists/${encodeURIComponent(String(l.id))}/tasks`, { queryParams: { $filter: "status ne 'completed'", $top: '25' } }));
            for (const t of tasks) {
              const due = (t.dueDateTime as { dateTime?: string } | undefined)?.dateTime;
              out.push({ title: t.title, list: l.displayName, importance: t.importance, due, overdue: !!due && new Date(due + 'Z') < now });
            }
          }
          return out.sort((a, b) => Number(b.overdue) - Number(a.overdue)).slice(0, 15);
        }),
        section(async () => list(await client.request('/me/chats', {
          queryParams: { $expand: 'lastMessagePreview', $orderby: 'lastMessagePreview/createdDateTime desc', $top: '8' },
        })).map(c => ({
          topic: c.topic ?? '(direct chat)', chatType: c.chatType,
          lastActivity: (c.lastMessagePreview as { createdDateTime?: string } | undefined)?.createdDateTime,
          ...untrusted(((c.lastMessagePreview as { body?: { content?: string } } | undefined)?.body?.content ?? '').replace(/<[^>]+>/g, ' ')),
        }))),
        section(async () => {
          const res = await client.request('/me/mailFolders/inbox', { queryParams: { $select: 'unreadItemCount' } });
          return [((res.data as Row).unreadItemCount as number) ?? 0] as unknown as Row[];
        }),
      ]);
      if (client.isDryRun()) return; // requests were traced; nothing to report

      const briefing: Briefing = {
        generatedAt: now.toISOString(), timeZone: tz, attention: [],
        calendar, mail: { unreadCount: unread.ok ? (unread.items[0] as unknown as number) : null, important, flagged }, tasks, teams,
      };
      briefing.attention = attentionLines(briefing, now);
      console.log(opts.output === 'text' ? renderText(briefing) : formatOutput(briefing, opts.output as OutputFormat));
    }));
}

function renderText(b: Briefing): string {
  const out: string[] = [`Briefing for ${b.generatedAt.slice(0, 10)} (${b.timeZone})`, ''];
  out.push('Needs attention');
  out.push(...(b.attention.length ? b.attention.map(l => `  • ${l}`) : ['  • nothing urgent']));
  const block = (title: string, s: Section<Row>, line: (r: Row) => string, empty: string) => {
    out.push('', title);
    if (!s.ok) out.push(`  (unavailable: ${s.error})`);
    else out.push(...(s.items.length ? s.items.map(r => `  ${line(r)}`) : [`  ${empty}`]));
  };
  block('Calendar', b.calendar, e => `${String(e.start).slice(11, 16)}–${String(e.end).slice(11, 16)}  ${e.subject}${e.overlapsPrevious ? '  (overlaps)' : ''}`, 'no meetings');
  block(`Mail (${b.mail.unreadCount ?? '?'} unread) — high importance`, b.mail.important, m => `${m.from}: ${m.subject}`, 'none');
  block('Mail — flagged', b.mail.flagged, m => `${m.from}: ${m.subject}`, 'none');
  block('Tasks', b.tasks, t => `${t.overdue ? '[overdue] ' : ''}${t.title} (${t.list})`, 'no open tasks');
  block('Recent Teams chats', b.teams, c => `${c.topic}: ${c.preview}`, 'none');
  return out.join('\n');
}
