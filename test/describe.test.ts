import { describe, it, expect } from 'vitest';
import fs from 'fs';
import { Command, Option } from 'commander';
import { runCommand } from './helpers/cli.js';
import { registerAuthCommands } from '../src/commands/auth.js';
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
import { registerSchemaCommands } from '../src/commands/schema.js';
import { registerGenerateSkillsCommand } from '../src/commands/generate-skills.js';
import { registerSetupCommand } from '../src/commands/setup.js';
import { registerMeetingsCommands } from '../src/commands/meetings.js';
import { registerWatchCommands } from '../src/commands/watch.js';
import { registerAuditCommands } from '../src/commands/audit.js';
import { registerIdempotencyCommands } from '../src/commands/idempotency.js';
import { registerDescribeCommand } from '../src/commands/describe.js';
import { registerMcpCommands } from '../src/commands/mcp.js';
import { registerResolveCommands } from '../src/commands/resolve.js';
import { registerNotesCommands } from '../src/commands/notes.js';
import { registerCompletionCommand } from '../src/commands/completion.js';
import { commandMeta, catalogPaths } from '../src/catalog.js';
import type { GraphClient } from '../src/graph-client.js';
import type { AuthManager } from '../src/auth.js';

/** Every command module, registered exactly as src/index.ts does. `file` is its source, for the confirm() scan. */
type Module = { file: string; register: (p: Command, g: () => GraphClient, a: () => AuthManager) => void };
const MODULES: Module[] = [
  { file: 'auth.ts', register: (p, _g, a) => registerAuthCommands(p, a) },
  { file: 'mail.ts', register: (p, g) => registerMailCommands(p, g) },
  { file: 'calendar.ts', register: (p, g) => registerCalendarCommands(p, g) },
  { file: 'drive.ts', register: (p, g) => registerDriveCommands(p, g) },
  { file: 'todo.ts', register: (p, g) => registerTodoCommands(p, g) },
  { file: 'contacts.ts', register: (p, g) => registerContactsCommands(p, g) },
  { file: 'teams.ts', register: (p, g) => registerTeamsCommands(p, g) },
  { file: 'sharepoint.ts', register: (p, g) => registerSharePointCommands(p, g) },
  { file: 'planner.ts', register: (p, g) => registerPlannerCommands(p, g) },
  { file: 'onenote.ts', register: (p, g) => registerOneNoteCommands(p, g) },
  { file: 'excel.ts', register: (p, g) => registerExcelCommands(p, g) },
  { file: 'search.ts', register: (p, g) => registerSearchCommands(p, g) },
  { file: 'user.ts', register: (p, g) => registerUserCommands(p, g) },
  { file: 'workflow.ts', register: (p, g) => registerWorkflowCommands(p, g) },
  { file: 'groups.ts', register: (p, g) => registerGroupsCommands(p, g) },
  { file: 'schema.ts', register: (p) => registerSchemaCommands(p) },
  { file: 'generate-skills.ts', register: (p) => registerGenerateSkillsCommand(p) },
  { file: 'setup.ts', register: (p, _g, a) => registerSetupCommand(p, a) },
  { file: 'meetings.ts', register: (p, g) => registerMeetingsCommands(p, g) },
  { file: 'watch.ts', register: (p, g) => registerWatchCommands(p, g) },
  { file: 'audit.ts', register: (p) => registerAuditCommands(p) },
  { file: 'idempotency.ts', register: (p) => registerIdempotencyCommands(p) },
  { file: 'describe.ts', register: (p) => registerDescribeCommand(p) },
  { file: 'mcp.ts', register: (p) => registerMcpCommands(p) },
  { file: 'resolve.ts', register: (p, g) => registerResolveCommands(p, g) },
  { file: 'notes.ts', register: (p, g) => registerNotesCommands(p, g) },
  { file: 'completion.ts', register: (p) => registerCompletionCommand(p) },
];

const noAuth = (() => { throw new Error('describe must not touch auth'); }) as () => AuthManager;

/** The full ms365 program: index.ts's global options plus every module. */
function registerAll(program: Command, getClient: () => GraphClient): void {
  program
    .version('0.1.0')
    .option('--org-mode', 'Enable organization mode (Teams, SharePoint, shared mailboxes)')
    .option('--dry-run', 'Preview the API request without executing it')
    .addOption(new Option('--api-version <version>', 'Graph API version').choices(['v1.0', 'beta']));
  for (const m of MODULES) m.register(program, getClient, noAuth);
}

function leafPaths(program: Command): string[] {
  const out: string[] = [];
  const walk = (c: Command, chain: string[]) => {
    if (c.commands.length === 0) out.push(chain.join(' '));
    for (const s of c.commands) walk(s, [...chain, s.name()]);
  };
  for (const top of program.commands) walk(top, [top.name()]);
  return out;
}

async function describeJson(...args: string[]) {
  const r = await runCommand(registerAll, ['describe', ...args, '--json']);
  return { ...r, json: JSON.parse(r.stdout) };
}

describe('ms365 describe --json', () => {
  it('prints a parseable catalogue of every leaf command', async () => {
    const { json, exitCode, calls } = await describeJson();
    expect(exitCode).not.toBe(1);
    expect(calls).toHaveLength(0);
    expect(json.version).toBe('0.1.0');
    expect(json.commands.length).toBeGreaterThanOrEqual(150);

    const program = new Command();
    registerAll(program, () => { throw new Error('unused'); });
    expect(json.commands.map((c: { path: string }) => c.path).sort()).toEqual(leafPaths(program).sort());
  });

  it('describes arguments, options and catalog metadata', async () => {
    const { json } = await describeJson('mail', 'get');
    expect(json.commands).toHaveLength(1);
    const cmd = json.commands[0];
    expect(cmd.path).toBe('mail get');
    expect(cmd.description).toMatch(/message/i);
    expect(cmd.arguments).toEqual([{ name: 'messageId', required: true, variadic: false }]);
    const output = cmd.options.find((o: { long: string }) => o.long === '--output');
    expect(output).toMatchObject({ short: '-o', takesValue: true, required: false });
    expect(output.choices).toContain('json');
    expect(cmd).toMatchObject({ scopes: ['Mail.Read'], destructive: false, idempotent: true, offline: false });
  });

  it('reports option defaults and mandatory options', async () => {
    const { json } = await describeJson('search', 'query');
    const entity = json.commands[0].options.find((o: { long: string }) => o.long === '--entity');
    expect(entity).toMatchObject({ flags: '--entity <types>', takesValue: true, default: 'message,event,driveItem', choices: null });
    const all = (await describeJson()).json.commands.flatMap((c: { options: Array<{ required: boolean }> }) => c.options);
    expect(all.some((o: { required: boolean }) => o.required)).toBe(true);
  });

  it('shows per-option and per-argument scope clauses (scopesBy)', async () => {
    const { json } = await describeJson();
    const byPath = new Map(json.commands.map((c: { path: string }) => [c.path, c]));
    expect(byPath.get('search query')).toMatchObject({ scopesBy: { option: '--entity', map: { driveItem: ['Files.Read.All', 'Sites.Read.All'] } } });
    expect(byPath.get('watch')).toMatchObject({ scopesBy: { argument: 'resource', map: { files: ['Files.Read'] } } });
    expect(byPath.get('resolve')).toMatchObject({
      scopesBy: { argument: 'type', map: { 'mail-folder': ['Mail.Read'] }, withOption: { option: '--user', map: { 'mail-folder': ['Mail.Read.Shared'] } } },
    });
    expect(byPath.get('mail list')).toMatchObject({ scopesBy: null });
    expect((await runCommand(registerAll, ['describe', 'watch'])).stdout).toMatch(/by <resource>/);
  });

  it('marks deletes destructive and offline commands offline', async () => {
    const { json } = await describeJson();
    const byPath = new Map(json.commands.map((c: { path: string }) => [c.path, c]));
    expect(byPath.get('mail delete')).toMatchObject({ destructive: true, scopes: ['Mail.ReadWrite'] });
    expect(byPath.get('mail send')).toMatchObject({ destructive: false, idempotent: false, scopes: ['Mail.Send'] });
    expect(byPath.get('describe')).toMatchObject({ offline: true, scopes: [] });
    expect(byPath.get('schema')).toMatchObject({ offline: true });
    expect(byPath.get('calendar list')).toMatchObject({ offline: false });
  });

  it('filters by command path prefix', async () => {
    const { json } = await describeJson('planner');
    expect(json.commands.length).toBeGreaterThan(10);
    expect(json.commands.every((c: { path: string }) => c.path.startsWith('planner '))).toBe(true);
    const exact = await describeJson('mail', 'list');
    expect(exact.json.commands.map((c: { path: string }) => c.path)).toEqual(['mail list']);
  });

  it('rejects an unknown command path with a usage error', async () => {
    const r = await runCommand(registerAll, ['describe', 'nope', '--json']);
    expect(r.exitCode).toBe(2);
    expect(r.stdout).toBe('');
    expect(r.stderr).toMatch(/nope/);
  });

  it('lists the global options once at the top level', async () => {
    const { json } = await describeJson();
    const longs = json.globalOptions.map((o: { long: string }) => o.long);
    expect(longs).toEqual(expect.arrayContaining(['--org-mode', '--dry-run', '--api-version', '--version']));
    expect(json.globalOptions.find((o: { long: string }) => o.long === '--api-version').choices).toEqual(['v1.0', 'beta']);
    expect(json.commands[0].options.map((o: { long: string }) => o.long)).not.toContain('--org-mode');
  });

  it('prints a readable summary without --json', async () => {
    const r = await runCommand(registerAll, ['describe', 'mail']);
    expect(() => JSON.parse(r.stdout)).toThrow();
    expect(r.stdout).toMatch(/^mail delete <messageId>\s+Delete a message.*destructive/m);
    expect(r.stdout).toMatch(/^mail list\b.*Mail\.Read/m);
  });
});

describe('command catalog coverage guard', () => {
  const program = new Command();
  registerAll(program, () => { throw new Error('unused'); });

  it('covers every module src/index.ts registers (add new modules to MODULES above)', () => {
    const index = fs.readFileSync(new URL('../src/index.ts', import.meta.url), 'utf8');
    const registered = [...index.matchAll(/^(register\w+)\(program\b/gm)].map(m => m[1]);
    const covered = MODULES.map(m => /register\w+/.exec(m.register.toString())![0]);
    expect(registered.length).toBeGreaterThan(10);
    expect(registered.filter(r => !covered.includes(r))).toEqual([]);
  });

  it('has metadata for every leaf command (add new commands to src/catalog.ts)', () => {
    const missing = leafPaths(program).filter(p => !commandMeta(p));
    expect(missing).toEqual([]);
  });

  it('has no metadata for commands that do not exist', () => {
    const leaves = new Set(leafPaths(program));
    expect(catalogPaths().filter(p => !leaves.has(p))).toEqual([]);
  });

  it('marks every command whose source calls confirm() as destructive', () => {
    const flagged: string[] = [];
    for (const m of MODULES) {
      const src = fs.readFileSync(new URL(`../src/commands/${m.file}`, import.meta.url), 'utf8');
      const own = new Command();
      m.register(own, () => { throw new Error('unused'); }, noAuth);
      const leaves = leafPaths(own);
      // Each `.command('name …')` starts a segment that runs until the next one.
      for (const seg of src.split(/\.command\(/).slice(1)) {
        if (!/\bconfirm\(/.test(seg)) continue;
        const name = /^['"`]([\w-]+)/.exec(seg)![1];
        const matches = leaves.filter(p => p.split(' ').at(-1) === name);
        expect(matches, `${m.file}: ${name}`).toHaveLength(1);
        flagged.push(matches[0]);
      }
    }
    expect(flagged.length).toBeGreaterThanOrEqual(8);
    expect(flagged.filter(p => !commandMeta(p)?.destructive)).toEqual([]);
  });
});
