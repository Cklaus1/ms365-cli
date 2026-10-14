import { describe, it, expect, afterEach } from 'vitest';
import { runCommand } from './helpers/cli.js';
import { GraphClient } from '../src/graph-client.js';
import { CliError, ExitCode } from '../src/errors.js';
import { resolveId } from '../src/resolve.js';
import { registerResolveCommands } from '../src/commands/resolve.js';
import { registerTodoCommands } from '../src/commands/todo.js';
import { registerTeamsCommands } from '../src/commands/teams.js';
import { registerPlannerCommands } from '../src/commands/planner.js';

const LIST_ID = 'AAMkAGI2THVSAAA=';
const TEAM_ID = '02bd9fd6-8f93-4758-87c3-1fb73740a315';
const CHANNEL_ID = '19:09fc54a3141a45d0bc769cf506d2e079@thread.tacv2';
const PLAN_ID = 'xqQg5FS2LkCp935s-FIFm2QAFkHM';
const BUCKET_ID = 'hsOf2dhOJkqyYYZEtdzDe2QAIUCR';

const lists = { value: [
  { id: 'L-groceries', displayName: 'Groceries' },
  { id: 'L-tasks', displayName: 'Tasks' },
  { id: 'L-work-a', displayName: 'Work Alpha' },
  { id: 'L-work-b', displayName: 'Work Beta' },
] };

const realFetch = globalThis.fetch;
afterEach(() => { globalThis.fetch = realFetch; });

/** A client whose fetch answers from `routes` and records every request path. */
function fakeClient(routes: Record<string, unknown>) {
  const paths: string[] = [];
  globalThis.fetch = (async (input: string | URL) => {
    const p = decodeURIComponent(String(input).replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/, ''));
    paths.push(p);
    const key = Object.keys(routes).find(k => p.startsWith(k));
    return new Response(JSON.stringify(key ? routes[key] : {}), { status: 200, headers: { 'content-type': 'application/json' } });
  }) as typeof fetch;
  const client = new GraphClient({ getToken: async () => 't' } as never);
  return { client, paths };
}

async function caught(p: Promise<unknown>): Promise<CliError> {
  try { await p; } catch (e) { return e as CliError; }
  throw new Error('expected rejection');
}

describe('resolveId', () => {
  it('passes IDs through without a request', async () => {
    const { client, paths } = fakeClient({});
    expect((await resolveId(client, 'todo-list', LIST_ID)).id).toBe(LIST_ID);
    expect((await resolveId(client, 'team', TEAM_ID)).id).toBe(TEAM_ID);
    expect((await resolveId(client, 'channel', CHANNEL_ID, TEAM_ID)).id).toBe(CHANNEL_ID);
    expect((await resolveId(client, 'plan', PLAN_ID)).id).toBe(PLAN_ID);
    expect((await resolveId(client, 'bucket', BUCKET_ID)).id).toBe(BUCKET_ID);
    expect(paths).toHaveLength(0);
  });

  it('matches names exactly, case-insensitively, then by unique prefix or substring', async () => {
    const { client } = fakeClient({ '/me/todo/lists': lists });
    expect(await resolveId(client, 'todo-list', 'groceries')).toEqual({ type: 'todo-list', id: 'L-groceries', displayName: 'Groceries' });
    expect((await resolveId(client, 'todo-list', 'Tas')).id).toBe('L-tasks');
    expect((await resolveId(client, 'todo-list', 'beta')).id).toBe('L-work-b');
  });

  it('prefers an exact match over prefix matches', async () => {
    const { client } = fakeClient({ '/me/todo/lists': { value: [{ id: 'a', displayName: 'Work' }, { id: 'b', displayName: 'Work Beta' }] } });
    expect((await resolveId(client, 'todo-list', 'work')).id).toBe('a');
  });

  it('several matches → exit 2 ambiguous_name listing candidates with IDs', async () => {
    const { client } = fakeClient({ '/me/todo/lists': lists });
    const err = await caught(resolveId(client, 'todo-list', 'work'));
    expect(err).toMatchObject({ code: 'ambiguous_name', exitCode: ExitCode.USAGE });
    expect(err.message).toContain('Work Alpha (L-work-a)');
    expect(err.message).toContain('Work Beta (L-work-b)');
  });

  it('no match → exit 5 not_found listing up to 10 available names', async () => {
    const many = { value: Array.from({ length: 12 }, (_, i) => ({ id: `id${i}`, displayName: `List ${i}` })) };
    const { client } = fakeClient({ '/me/todo/lists': many });
    const err = await caught(resolveId(client, 'todo-list', 'nope'));
    expect(err).toMatchObject({ code: 'not_found', exitCode: ExitCode.NOT_FOUND });
    expect(err.message).toContain('List 9');
    expect(err.message).not.toContain('List 10');
    expect(err.message).toContain('2 more');
  });

  it('follows nextLink pages when listing candidates', async () => {
    const { client } = fakeClient({
      '/me/todo/lists?page=2': { value: [{ id: 'L-late', displayName: 'Late' }] },
      '/me/todo/lists': { value: [], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/todo/lists?page=2' },
    });
    expect((await resolveId(client, 'todo-list', 'late')).id).toBe('L-late');
  });

  it('caches resolutions per client', async () => {
    const { client, paths } = fakeClient({ '/me/todo/lists': lists });
    await resolveId(client, 'todo-list', 'Groceries');
    await resolveId(client, 'todo-list', 'Groceries');
    expect(paths).toEqual(['/me/todo/lists']);
  });

  it('resolves a channel by name within a team, and a bucket within a plan', async () => {
    const { client, paths } = fakeClient({
      [`/teams/${TEAM_ID}/channels`]: { value: [{ id: CHANNEL_ID, displayName: 'General' }] },
      [`/planner/plans/${PLAN_ID}/buckets`]: { value: [{ id: BUCKET_ID, name: 'To do' }] },
    });
    expect((await resolveId(client, 'channel', 'general', TEAM_ID)).id).toBe(CHANNEL_ID);
    expect((await resolveId(client, 'bucket', 'to do', PLAN_ID)).id).toBe(BUCKET_ID);
    expect(paths).toEqual([`/teams/${TEAM_ID}/channels`, `/planner/plans/${PLAN_ID}/buckets`]);
  });

  it('plan names match on title', async () => {
    const { client } = fakeClient({ '/me/planner/plans': { value: [{ id: PLAN_ID, title: 'Q4 Launch' }] } });
    expect(await resolveId(client, 'plan', 'q4')).toEqual({ type: 'plan', id: PLAN_ID, displayName: 'Q4 Launch' });
  });

  it('a bucket name without a plan is a usage error', async () => {
    const { client, paths } = fakeClient({});
    const err = await caught(resolveId(client, 'bucket', 'To do'));
    expect(err).toMatchObject({ exitCode: ExitCode.USAGE });
    expect(paths).toHaveLength(0);
  });
});

const FOLDER_ID = 'AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAA=';
const TOP = "/me/mailFolders?$select=id,displayName,childFolderCount&$top=100";
const top = { value: [
  { id: 'F-inbox', displayName: 'Inbox', childFolderCount: 2 },
  { id: 'F-archive', displayName: 'Archive', childFolderCount: 0 },
  { id: 'F-receipts', displayName: 'Receipts', childFolderCount: 1 },
] };
const inboxKids = { value: [
  { id: 'F-projects', displayName: 'Projects', childFolderCount: 1 },
  { id: 'F-misc', displayName: 'Misc', childFolderCount: 0 },
] };
const projectKids = { value: [{ id: 'F-alpha', displayName: 'Alpha', childFolderCount: 0 }] };
const receiptKids = { value: [{ id: 'F-misc2', displayName: 'misc', childFolderCount: 0 }] };
const kids = (id: string) => `/me/mailFolders/${id}/childFolders?$select=id,displayName,childFolderCount&$top=100`;

describe('resolveId mail-folder', () => {
  it('passes well-known names (normalised) and Exchange IDs through without a request', async () => {
    const { client, paths } = fakeClient({});
    expect(await resolveId(client, 'mail-folder', 'Inbox')).toEqual({ type: 'mail-folder', id: 'inbox', displayName: null });
    expect((await resolveId(client, 'mail-folder', 'Sent Items')).id).toBe('sentitems');
    expect((await resolveId(client, 'mail-folder', 'JunkEmail')).id).toBe('junkemail');
    expect((await resolveId(client, 'mail-folder', 'deleteditems')).id).toBe('deleteditems');
    expect((await resolveId(client, 'mail-folder', FOLDER_ID)).id).toBe(FOLDER_ID);
    expect((await resolveId(client, 'mail-folder', 'AQMkADAwATM0MDAAMS0yMDkALTI2NjAtMDACLTAwCgAuAAAD')).id).toBe('AQMkADAwATM0MDAAMS0yMDkALTI2NjAtMDACLTAwCgAuAAAD');
    expect(paths).toHaveLength(0);
  });

  it('matches a top-level name exactly and case-insensitively, paging with $top=100', async () => {
    const { client, paths } = fakeClient({ '/me/mailFolders': top });
    expect(await resolveId(client, 'mail-folder', 'receipts')).toEqual({ type: 'mail-folder', id: 'F-receipts', displayName: 'Receipts' });
    expect(paths).toEqual([TOP]);
  });

  it('never prefix- or substring-matches (F11)', async () => {
    const { client } = fakeClient({ '/me/mailFolders': { value: [{ id: 'F-r', displayName: 'Receipts', childFolderCount: 0 }] } });
    const err = await caught(resolveId(client, 'mail-folder', 'Rec'));
    expect(err).toMatchObject({ code: 'not_found', exitCode: ExitCode.NOT_FOUND });
    expect(err.message).toContain('Receipts');
  });

  it('walks a path segment by segment through childFolders (F12)', async () => {
    const { client, paths } = fakeClient({ [kids('F-projects')]: projectKids, [kids('inbox')]: inboxKids, '/me/mailFolders': top });
    expect(await resolveId(client, 'mail-folder', 'Inbox/projects/alpha')).toEqual({ type: 'mail-folder', id: 'F-alpha', displayName: 'Alpha' });
    expect(paths).toEqual([kids('inbox'), kids('F-projects')]);
  });

  it('a path whose first segment is custom resolves it at the top level', async () => {
    const { client, paths } = fakeClient({ [kids('F-receipts')]: receiptKids, '/me/mailFolders': top });
    expect((await resolveId(client, 'mail-folder', 'Receipts/Misc')).id).toBe('F-misc2');
    expect(paths).toEqual([TOP, kids('F-receipts')]);
  });

  it('a missing path segment → not_found naming the parent and its children', async () => {
    const { client } = fakeClient({ [kids('inbox')]: inboxKids });
    const err = await caught(resolveId(client, 'mail-folder', 'Inbox/Nope'));
    expect(err).toMatchObject({ code: 'not_found', exitCode: ExitCode.NOT_FOUND });
    expect(err.message).toContain('"Nope" under "Inbox"');
    expect(err.message).toContain('Projects, Misc');
  });

  it('an empty path segment is a usage error', async () => {
    const { client, paths } = fakeClient({});
    expect(await caught(resolveId(client, 'mail-folder', 'Inbox//x'))).toMatchObject({ exitCode: ExitCode.USAGE });
    expect(paths).toHaveLength(0);
  });

  it('a top-level miss falls back to a breadth-first search, skipping leaf folders', async () => {
    const { client, paths } = fakeClient({
      [kids('F-projects')]: projectKids, [kids('F-inbox')]: inboxKids, [kids('F-receipts')]: { value: [] }, '/me/mailFolders': top,
    });
    expect(await resolveId(client, 'mail-folder', 'ALPHA')).toEqual({ type: 'mail-folder', id: 'F-alpha', displayName: 'Alpha' });
    expect(paths).toEqual([TOP, kids('F-inbox'), kids('F-receipts'), kids('F-projects')]);
  });

  it('two breadth-first matches → ambiguous_name listing their paths', async () => {
    const { client } = fakeClient({
      [kids('F-projects')]: projectKids, [kids('F-inbox')]: inboxKids, [kids('F-receipts')]: receiptKids, '/me/mailFolders': top,
    });
    const err = await caught(resolveId(client, 'mail-folder', 'misc'));
    expect(err).toMatchObject({ code: 'ambiguous_name', exitCode: ExitCode.USAGE });
    expect(err.message).toContain('Inbox/Misc (F-misc)');
    expect(err.message).toContain('Receipts/misc (F-misc2)');
  });

  it('the breadth-first search stops at 500 folders', async () => {
    const wide = { value: Array.from({ length: 600 }, (_, i) => ({ id: `W${i}`, displayName: `Folder ${i}`, childFolderCount: 0 })) };
    const { client, paths } = fakeClient({ [kids('F-inbox')]: wide, '/me/mailFolders': { value: [{ id: 'F-inbox', displayName: 'Inbox', childFolderCount: 600 }, { id: 'F-late', displayName: 'Late', childFolderCount: 3 }] } });
    const err = await caught(resolveId(client, 'mail-folder', 'Folder 550'));
    expect(err).toMatchObject({ code: 'not_found', exitCode: ExitCode.NOT_FOUND });
    expect(err.message).toContain('searched the first 500 folders');
    expect(paths).toEqual([TOP, kids('F-inbox')]);
  });

  it('a miss everywhere → not_found listing top-level names and suggesting a path', async () => {
    const { client } = fakeClient({ [kids('F-projects')]: projectKids, [kids('F-inbox')]: inboxKids, [kids('F-receipts')]: receiptKids, '/me/mailFolders': top });
    const err = await caught(resolveId(client, 'mail-folder', 'Recieps'));
    expect(err).toMatchObject({ code: 'not_found', exitCode: ExitCode.NOT_FOUND });
    expect(err.message).toContain('Inbox, Archive, Receipts');
    expect(err.fix).toContain('Inbox/Projects');
  });

  it('resolves against another mailbox when a user is given, keyed separately in the cache', async () => {
    const { client, paths } = fakeClient({
      '/users/shared@contoso.com/mailFolders': { value: [{ id: 'S-r', displayName: 'Receipts', childFolderCount: 0 }] },
      '/me/mailFolders': top,
    });
    expect((await resolveId(client, 'mail-folder', 'Receipts', 'shared@contoso.com')).id).toBe('S-r');
    expect((await resolveId(client, 'mail-folder', 'Receipts')).id).toBe('F-receipts');
    expect(paths[0]).toBe('/users/shared@contoso.com/mailFolders?$select=id,displayName,childFolderCount&$top=100');
  });
});

describe('ms365 resolve', () => {
  it('prints {type, id, displayName}', async () => {
    const r = await runCommand(registerResolveCommands, ['resolve', 'todo-list', 'groceries', '-o', 'json'], {
      routes: [['GET', '/me/todo/lists', () => lists]],
    });
    expect(r.exitCode).toBe(0);
    expect(JSON.parse(r.stdout)).toEqual({ type: 'todo-list', id: 'L-groceries', displayName: 'Groceries' });
  });

  it('resolves a channel with --team given by name', async () => {
    const r = await runCommand(registerResolveCommands, ['resolve', 'channel', 'General', '--team', 'Eng', '-o', 'json'], {
      routes: [
        ['GET', '/me/joinedTeams', () => ({ value: [{ id: TEAM_ID, displayName: 'Engineering' }] })],
        ['GET', `/teams/${TEAM_ID}/channels`, () => ({ value: [{ id: CHANNEL_ID, displayName: 'General' }] })],
      ],
    });
    expect(r.exitCode).toBe(0);
    expect(JSON.parse(r.stdout)).toMatchObject({ type: 'channel', id: CHANNEL_ID });
  });

  it('channel without --team is a usage error', async () => {
    const r = await runCommand(registerResolveCommands, ['resolve', 'channel', 'General']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('rejects an unknown type', async () => {
    const r = await runCommand(registerResolveCommands, ['resolve', 'folder', 'Inbox']);
    expect(r.exitCode).toBe(2);
  });

  it('resolve mail-folder <path> prints the nested folder, and --user scopes the mailbox', async () => {
    const r = await runCommand(registerResolveCommands, ['resolve', 'mail-folder', 'Inbox/Projects', '-o', 'json'], {
      routes: [['GET', '/me/mailFolders/inbox/childFolders', () => inboxKids]],
    });
    expect(r.exitCode).toBe(0);
    expect(JSON.parse(r.stdout)).toEqual({ type: 'mail-folder', id: 'F-projects', displayName: 'Projects' });
    const u = await runCommand(registerResolveCommands, ['resolve', 'mail-folder', 'Receipts', '--user', 'shared@contoso.com', '-o', 'json'], {
      routes: [['GET', '/users/shared@contoso.com/mailFolders', () => top]],
    });
    expect(JSON.parse(u.stdout)).toMatchObject({ id: 'F-receipts' });
  });

  it('resolve mail-folder: a miss exits 5; dry-run passes the name through', async () => {
    const miss = await runCommand(registerResolveCommands, ['resolve', 'mail-folder', 'Arch'], {
      routes: [['GET', /^\/me\/mailFolders\/[^/]+\/childFolders/, () => ({ value: [] })], ['GET', '/me/mailFolders', () => top]],
    });
    expect(miss.exitCode).toBe(5);
    const dry = await runCommand(registerResolveCommands, ['resolve', 'mail-folder', 'Receipts', '-o', 'json'], { dryRun: true });
    expect(dry.exitCode).toBe(0);
    expect(dry.stderr).toContain('[dry-run] GET https://graph.microsoft.com/v1.0/me/mailFolders?');
    expect(dry.stdout).toContain('"id": "Receipts"');
  });

  it('ambiguous name exits 2', async () => {
    const r = await runCommand(registerResolveCommands, ['resolve', 'todo-list', 'work'], {
      routes: [['GET', '/me/todo/lists', () => lists]],
    });
    expect(r.exitCode).toBe(2);
    expect(r.stderr).toContain('ambiguous');
  });
});

describe('commands accept names', () => {
  it('todo tasks <list name>', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'tasks', 'Groceries', '-o', 'json'], {
      routes: [['GET', /^\/me\/todo\/lists$/, () => lists]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls.map(c => c.path)).toEqual(['/me/todo/lists', '/me/todo/lists/L-groceries/tasks']);
  });

  it('todo create with a list ID makes no lookup', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'create', LIST_ID, '--title', 'x', '-o', 'json']);
    expect(r.calls.map(c => `${c.method} ${c.path}`)).toEqual([`POST /me/todo/lists/${LIST_ID}/tasks`]);
  });

  it('teams send-channel-message <team name> <channel name>', async () => {
    const r = await runCommand(registerTeamsCommands, ['teams', 'send-channel-message', 'Eng', 'General', '--body', 'hi'], {
      routes: [
        ['GET', '/me/joinedTeams', () => ({ value: [{ id: TEAM_ID, displayName: 'Engineering' }] })],
        ['GET', `/teams/${TEAM_ID}/channels`, () => ({ value: [{ id: CHANNEL_ID, displayName: 'General' }] })],
      ],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls.at(-1)).toMatchObject({ method: 'POST', path: `/teams/${TEAM_ID}/channels/${CHANNEL_ID}/messages` });
  });

  it('planner create-task --plan <name> --bucket <name>', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'create-task', '--plan', 'Q4', '--bucket', 'To do', '--title', 'x', '-o', 'json'], {
      routes: [
        ['GET', '/me/planner/plans', () => ({ value: [{ id: PLAN_ID, title: 'Q4 Launch' }] })],
        ['GET', `/planner/plans/${PLAN_ID}/buckets`, () => ({ value: [{ id: BUCKET_ID, name: 'To do' }] })],
      ],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls.at(-1)).toMatchObject({ method: 'POST', body: { planId: PLAN_ID, bucketId: BUCKET_ID } });
  });

  it('planner bucket-tasks <bucket name> --plan <plan ID>', async () => {
    const r = await runCommand(registerPlannerCommands, ['planner', 'bucket-tasks', 'To do', '--plan', PLAN_ID, '-o', 'json'], {
      routes: [['GET', `/planner/plans/${PLAN_ID}/buckets`, () => ({ value: [{ id: BUCKET_ID, name: 'To do' }] })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls.at(-1)!.path).toBe(`/planner/buckets/${BUCKET_ID}/tasks`);
  });

  it('input validation fails before any resolution request', async () => {
    const todo = await runCommand(registerTodoCommands, ['todo', 'create', 'Groceries', '--title', 'x', '--due', 'someday']);
    expect(todo.exitCode).toBe(2);
    expect(todo.calls).toHaveLength(0);
    const planner = await runCommand(registerPlannerCommands, ['planner', 'create-task', '--plan', 'Q4', '--title', 'x', '--due', 'someday']);
    expect(planner.exitCode).toBe(2);
    expect(planner.calls).toHaveLength(0);
  });

  it('dry-run previews the resolution GET and falls back to the input', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'tasks', 'Groceries'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain('[dry-run] GET https://graph.microsoft.com/v1.0/me/todo/lists\n');
    expect(r.stderr).toContain('/me/todo/lists/Groceries/tasks');
  });
});
