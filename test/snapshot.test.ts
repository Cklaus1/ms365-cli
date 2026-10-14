import { describe, it, expect, beforeAll, afterAll, vi } from 'vitest';
import { Command, Option } from 'commander';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { runCommand } from './helpers/cli.js';
import { registerMailCommands } from '../src/commands/mail.js';
import { registerCalendarCommands } from '../src/commands/calendar.js';
import { registerDriveCommands } from '../src/commands/drive.js';
import { registerTodoCommands } from '../src/commands/todo.js';
import { registerContactsCommands } from '../src/commands/contacts.js';
import { registerTeamsCommands } from '../src/commands/teams.js';
import { registerSharePointCommands } from '../src/commands/sharepoint.js';
import { registerPlannerCommands } from '../src/commands/planner.js';
import { registerOneNoteCommands } from '../src/commands/onenote.js';
import { registerExcelCommands } from '../src/commands/excel.js';
import { registerSearchCommands } from '../src/commands/search.js';
import { registerUserCommands } from '../src/commands/user.js';
import { registerWorkflowCommands } from '../src/commands/workflow.js';
import { registerGroupsCommands } from '../src/commands/groups.js';
import { registerMeetingsCommands } from '../src/commands/meetings.js';
import { registerWatchCommands } from '../src/commands/watch.js';
import { registerNotesCommands } from '../src/commands/notes.js';
import type { GraphClient } from '../src/graph-client.js';

/**
 * Dry-run endpoint snapshot: every Graph-calling command is invoked with synthetic
 * arguments under --dry-run, and the requests it would send (method + URL) are recorded.
 * A diff in the snapshot means a command's wire behavior changed — review it deliberately.
 */
type Register = (p: Command, g: () => GraphClient) => void;
const MODULES: Register[] = [
  registerMailCommands, registerCalendarCommands, registerDriveCommands, registerTodoCommands,
  registerContactsCommands, registerTeamsCommands, registerSharePointCommands, registerPlannerCommands,
  registerOneNoteCommands, registerExcelCommands, registerSearchCommands, registerUserCommands,
  registerWorkflowCommands, registerGroupsCommands, registerMeetingsCommands, registerWatchCommands, registerNotesCommands,
];

let tmp: string;
let sampleFile: string;
const cwd = process.cwd();

function leafCommands(register: Register): Array<{ chain: string[]; cmd: Command }> {
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

function sampleArg(name: string): string {
  if (/remote|drive|^path$|sitePath/i.test(name)) return '/Docs/a.txt';
  if (/file|path/i.test(name)) return sampleFile;
  if (/ids$/i.test(name)) return 'ID1,ID2';
  return name.toUpperCase();
}

function sampleOption(o: Option): string[] {
  const flag = o.long!;
  if (!o.required && !o.optional) return [];
  if (o.argChoices?.length) return [flag, o.argChoices[0]];
  if (/to|cc|bcc|attendees|emails|assigned|mailbox|address|forward/i.test(flag)) return [flag, 'a@contoso.com'];
  if (/start|end|due|date|propose|since/i.test(flag)) return [flag, '2026-10-20T10:00'];
  if (/top|limit|delay|percent|priority|size/i.test(flag)) return [flag, '5'];
  if (/duration/i.test(flag)) return [flag, '30m'];
  if (/fields/i.test(flag)) return [flag, '[{"key":0,"ascending":true}]'];
  if (/categories|series|values|body-json|json/i.test(flag)) return [flag, '{}'];
  if (/timezone/i.test(flag)) return [flag, 'America/Los_Angeles'];
  if (/file/i.test(flag)) return [flag, sampleFile];
  if (/content/i.test(flag) && !/type/i.test(flag)) return [flag, 'aGVsbG8='];
  if (/type$/i.test(flag)) return [flag, 'text/plain'];
  return [flag, 'x'];
}

beforeAll(() => {
  tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-snap-'));
  sampleFile = path.join(tmp, 'sample.txt');
  fs.writeFileSync(sampleFile, 'hello');
  process.chdir(tmp);
  vi.stubEnv('MS365_CLI_TIMEZONE', 'America/Los_Angeles');
  vi.setSystemTime(new Date('2026-10-14T17:00:00Z'));
});
afterAll(() => { process.chdir(cwd); vi.unstubAllEnvs(); vi.useRealTimers(); });

describe('dry-run endpoint snapshot (every command)', () => {
  for (const register of MODULES) {
    for (const { chain, cmd } of leafCommands(register)) {
      it(chain.join(' '), async () => {
        const args = cmd.registeredArguments.filter(a => a.required).map(a => sampleArg(a.name()));
        const opts = cmd.options.filter(o => o.mandatory).flatMap(sampleOption);
        const r = await runCommand(register, [...chain, ...args, ...opts], { dryRun: true, tty: true, answer: 'y' });
        const requests = r.stdout.split('\n').filter(l => l.startsWith('[dry-run] ') && !l.startsWith('[dry-run] Body') && !l.startsWith('[dry-run] Headers') && !l.startsWith('[dry-run] File'));
        const firstErr = r.stderr.split('\n').find(l => l.startsWith('Error:')) ?? '';
        expect({ argv: [...chain, ...args, ...opts].join(' ').replaceAll(tmp, '<tmp>'), exitCode: r.exitCode, requests: requests.map(l => l.replaceAll(tmp, '<tmp>')), error: firstErr.replaceAll(tmp, '<tmp>') }).toMatchSnapshot();
      });
    }
  }
});
