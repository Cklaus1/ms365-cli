import { Command, Option } from 'commander';
import fs from 'fs';
import path from 'path';
import { registerMailCommands } from '../../src/commands/mail.js';
import { registerCalendarCommands } from '../../src/commands/calendar.js';
import { registerDriveCommands } from '../../src/commands/drive.js';
import { registerTodoCommands } from '../../src/commands/todo.js';
import { registerContactsCommands } from '../../src/commands/contacts.js';
import { registerTeamsCommands } from '../../src/commands/teams.js';
import { registerSharePointCommands } from '../../src/commands/sharepoint.js';
import { registerPlannerCommands } from '../../src/commands/planner.js';
import { registerOneNoteCommands } from '../../src/commands/onenote.js';
import { registerExcelCommands } from '../../src/commands/excel.js';
import { registerSearchCommands } from '../../src/commands/search.js';
import { registerUserCommands } from '../../src/commands/user.js';
import { registerWorkflowCommands } from '../../src/commands/workflow.js';
import { registerGroupsCommands } from '../../src/commands/groups.js';
import { registerMeetingsCommands } from '../../src/commands/meetings.js';
import { registerReportsCommands } from '../../src/commands/reports.js';
import { registerAuditTrailCommands } from '../../src/commands/audit-trail.js';
import { registerWatchCommands } from '../../src/commands/watch.js';
import { registerNotesCommands } from '../../src/commands/notes.js';
import type { GraphClient } from '../../src/graph-client.js';

/** Every Graph-calling command module, with synthetic argument/option samplers (snapshot + lint guards). */
export type Register = (p: Command, g: () => GraphClient) => void;
export const MODULES: Register[] = [
  registerMailCommands, registerCalendarCommands, registerDriveCommands, registerTodoCommands,
  registerContactsCommands, registerTeamsCommands, registerSharePointCommands, registerPlannerCommands,
  registerOneNoteCommands, registerExcelCommands, registerSearchCommands, registerUserCommands,
  registerWorkflowCommands, registerGroupsCommands, registerMeetingsCommands, registerWatchCommands, registerNotesCommands, registerReportsCommands,
  registerAuditTrailCommands,
];

export function leafCommands(register: Register): Array<{ chain: string[]; cmd: Command }> {
  const program = new Command();
  register(program, () => { throw new Error('unused'); });
  const out: Array<{ chain: string[]; cmd: Command }> = [];
  const walk = (c: Command, chain: string[]) => {
    if (c.commands.length === 0) out.push({ chain, cmd: c });
    for (const sub of c.commands) walk(sub, [...chain, sub.name()]);
  };
  for (const top of program.commands) walk(top, [top.name()]);
  return out;
}

/** Sample input files, written by writeSampleFiles() into a temp dir. */
export const samples = { file: '', photo: '' };

export function writeSampleFiles(dir: string): void {
  samples.file = path.join(dir, 'sample.txt');
  samples.photo = path.join(dir, 'photo.jpg');
  fs.writeFileSync(samples.file, 'hello');
  fs.writeFileSync(samples.photo, Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0, 0x10, 0x4a, 0x46, 0x49, 0x46, 0, 0xff, 0xd9]));
}

export function sampleArg(name: string, chain: string[] = [], choices?: readonly string[]): string {
  if (choices?.length) return choices[0];
  if (/remote|drive|^path$|sitePath/i.test(name)) return '/Docs/a.txt';
  if (chain.join(' ') === 'user upload-photo') return samples.photo;
  if (/file|path/i.test(name)) return samples.file;
  if (/^upn$/i.test(name)) return 'a@contoso.com';
  if (/ids$/i.test(name)) return 'ID1,ID2';
  if (/^timezone$/i.test(name)) return 'UTC';
  return name.toUpperCase();
}

export function sampleOption(o: Option): string[] {
  const flag = o.long!;
  if (!o.required && !o.optional) return [];
  if (o.argChoices?.length) return [flag, o.argChoices[0]];
  if (/to\b|cc|bcc|attendees|emails|assigned|mailbox|address|forward|members|mention|^--by$/i.test(flag)) return [flag, 'a@contoso.com'];
  if (/work-start/.test(flag)) return [flag, '9'];
  if (/work-end/.test(flag)) return [flag, '17'];
  if (/due/i.test(flag)) return [flag, '2026-10-20'];
  if (/until/i.test(flag)) return [flag, '2026-10-21T10:00'];
  if (/start|end|date|propose|since/i.test(flag)) return [flag, '2026-10-20T10:00'];
  if (/percent/i.test(flag)) return [flag, '50'];
  if (/interval/i.test(flag)) return [flag, '15'];
  if (flag === '--from') return [flag, '0'];
  if (/top|limit|delay|priority|size|max-|timeout/i.test(flag)) return [flag, '5'];
  if (/attach/i.test(flag)) return [flag, samples.file];
  if (/entity/i.test(flag)) return [flag, 'message'];
  if (/duration/i.test(flag)) return [flag, '30m'];
  if (/fields/i.test(flag)) return [flag, '[{"key":0,"ascending":true}]'];
  if (/categories|series|values|body-json|json/i.test(flag)) return [flag, '{}'];
  if (/timezone/i.test(flag)) return [flag, 'America/Los_Angeles'];
  if (/file/i.test(flag)) return [flag, samples.file];
  if (/content/i.test(flag) && !/type/i.test(flag)) return [flag, 'aGVsbG8='];
  if (/type$/i.test(flag)) return [flag, 'text/plain'];
  return [flag, 'x'];
}

/** Positional args and mandatory options for one synthetic run of `cmd`. */
export function sampleArgv(chain: string[], cmd: Command): string[] {
  const args = cmd.registeredArguments.filter(a => a.required).map(a => sampleArg(a.name(), chain, a.argChoices));
  const opts = cmd.options.filter(o => o.mandatory).flatMap(sampleOption);
  return [...chain, ...args, ...opts];
}

/** Every .ts file under src/, as { rel: 'commands/mail.ts', lines } (for source-scan guards). */
export function sourceFiles(): Array<{ rel: string; lines: string[] }> {
  const root = path.resolve(__dirname, '../../src');
  const walk = (dir: string): string[] => fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) => {
    const p = path.join(dir, e.name);
    return e.isDirectory() ? walk(p) : e.name.endsWith('.ts') ? [p] : [];
  });
  return walk(root).map(p => ({ rel: path.relative(root, p).split(path.sep).join('/'), lines: fs.readFileSync(p, 'utf8').split('\n') }));
}
