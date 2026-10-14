import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { runCommand } from '../helpers/cli.js';
import { registerExcelCommands } from '../../src/commands/excel.js';
import { registerTodoCommands } from '../../src/commands/todo.js';
import { registerTeamsCommands } from '../../src/commands/teams.js';
import { registerNotesCommands } from '../../src/commands/notes.js';
import { registerSearchCommands } from '../../src/commands/search.js';

const LIST = 'AAMkAGI2TGuLAAA=' + 'x'.repeat(40);
const TEAM = '00000000-0000-0000-0000-0000000000a1';
const CHANNEL = '19:ch1@thread.tacv2';

let tmp: string;
beforeAll(() => { tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-coverage-')); });
afterAll(() => { fs.rmSync(tmp, { recursive: true, force: true }); });

describe('excel tables (PM-9)', () => {
  const wb = '/drives/D1/items/I1/workbook';

  it('tables lists the workbook tables', async () => {
    const r = await runCommand(registerExcelCommands, ['excel', 'tables', 'D1', 'I1'], {
      routes: [['GET', `${wb}/tables`, () => ({ value: [{ id: '{T1}', name: 'Sales' }] })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].path).toBe(`${wb}/tables`);
    expect(JSON.parse(r.stdout).value[0].name).toBe('Sales');
  });

  it('table-rows encodes the table name and passes --top', async () => {
    const r = await runCommand(registerExcelCommands, ['excel', 'table-rows', 'D1', 'I1', 'Q1 Sales', '--top', '5'], {
      routes: [['GET', `${wb}/tables/`, () => ({ value: [{ index: 0, values: [['a', 1]] }] })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].url).toContain('/workbook/tables/Q1%20Sales/rows?$top=5');
  });

  it('add-rows POSTs {values} from --values JSON', async () => {
    const r = await runCommand(registerExcelCommands, ['excel', 'add-rows', 'D1', 'I1', 'Sales', '--values', '[["a",1],["b",2]]'], {
      routes: [['POST', `${wb}/tables/Sales/rows`, () => ({ index: 3 })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].method).toBe('POST');
    expect(r.calls[0].path).toBe(`${wb}/tables/Sales/rows`);
    expect(r.calls[0].body).toEqual({ values: [['a', 1], ['b', 2]] });
  });

  it.each([['{"a":1}'], ['[1,2]'], ['not json'], ['[]']])('add-rows rejects --values %s with exit 2', async (values) => {
    const r = await runCommand(registerExcelCommands, ['excel', 'add-rows', 'D1', 'I1', 'Sales', '--values', values]);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('add-rows needs exactly one of --values / --values-file', async () => {
    const none = await runCommand(registerExcelCommands, ['excel', 'add-rows', 'D1', 'I1', 'Sales']);
    expect(none.exitCode).toBe(2);
    const file = path.join(tmp, 'rows.json');
    fs.writeFileSync(file, '[["a"]]');
    const both = await runCommand(registerExcelCommands, ['excel', 'add-rows', 'D1', 'I1', 'Sales', '--values', '[["a"]]', '--values-file', file]);
    expect(both.exitCode).toBe(2);
  });

  it('add-rows reads --values-file JSON', async () => {
    const file = path.join(tmp, 'rows2.json');
    fs.writeFileSync(file, '[["x", 9]]');
    const r = await runCommand(registerExcelCommands, ['excel', 'add-rows', 'D1', 'I1', 'Sales', '--values-file', file]);
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toEqual({ values: [['x', 9]] });
  });

  it('add-rows parses --values-file CSV with quoted fields', async () => {
    const file = path.join(tmp, 'rows.csv');
    fs.writeFileSync(file, 'a,b,c\r\n"x, y","say ""hi""",3\n"multi\nline",,\n');
    const r = await runCommand(registerExcelCommands, ['excel', 'add-rows', 'D1', 'I1', 'Sales', '--values-file', file]);
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toEqual({ values: [['a', 'b', 'c'], ['x, y', 'say "hi"', '3'], ['multi\nline', '', '']] });
  });

  it('add-rows: missing --values-file is exit 2', async () => {
    const r = await runCommand(registerExcelCommands, ['excel', 'add-rows', 'D1', 'I1', 'Sales', '--values-file', path.join(tmp, 'nope.csv')]);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});

describe('todo checklist items (PM-9)', () => {
  const base = `/me/todo/lists/${LIST}/tasks/T1/checklistItems`;

  it('checklist lists items', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'checklist', LIST, 'T1'], {
      routes: [['GET', base, () => ({ value: [{ id: 'C1', displayName: 'step', isChecked: false }] })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].path).toBe(base);
    expect(JSON.parse(r.stdout).value[0].id).toBe('C1');
  });

  it('checklist resolves the list by name', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'checklist', 'Groceries', 'T1'], {
      routes: [
        ['GET', '/me/todo/lists/L9/tasks', () => ({ value: [] })],
        ['GET', '/me/todo/lists', () => ({ value: [{ id: 'L9', displayName: 'Groceries' }] })],
      ],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls.map(c => c.path)).toEqual(['/me/todo/lists', '/me/todo/lists/L9/tasks/T1/checklistItems']);
  });

  it('check-add POSTs displayName', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'check-add', LIST, 'T1', '--title', 'buy milk'], {
      routes: [['POST', base, () => ({ id: 'C2', displayName: 'buy milk' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toEqual({ displayName: 'buy milk' });
    expect(JSON.parse(r.stdout).id).toBe('C2');
  });

  it('check-done PATCHes isChecked true, --undo false', async () => {
    const done = await runCommand(registerTodoCommands, ['todo', 'check-done', LIST, 'T1', 'C1']);
    expect(done.calls[0]).toMatchObject({ method: 'PATCH', path: `${base}/C1`, body: { isChecked: true } });
    const undo = await runCommand(registerTodoCommands, ['todo', 'check-done', LIST, 'T1', 'C1', '--undo']);
    expect(undo.calls[0].body).toEqual({ isChecked: false });
  });

  it('check-delete confirms, then DELETEs', async () => {
    const refused = await runCommand(registerTodoCommands, ['todo', 'check-delete', LIST, 'T1', 'C1']);
    expect(refused.exitCode).toBe(7);
    expect(refused.calls).toHaveLength(0);
    const ok = await runCommand(registerTodoCommands, ['todo', 'check-delete', LIST, 'T1', 'C1'], { tty: true, answer: 'y' });
    expect(ok.exitCode).toBe(0);
    expect(ok.calls[0]).toMatchObject({ method: 'DELETE', path: `${base}/C1` });
  });
});

describe('chat create (PM-9)', () => {
  const me = ['GET', '/me', () => ({ id: 'ME' })] as const;
  const member = (id: string) => ({
    '@odata.type': '#microsoft.graph.aadUserConversationMember',
    roles: ['owner'],
    'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${id}')`,
  });

  it('one member → oneOnOne including the signed-in user', async () => {
    const r = await runCommand(registerTeamsCommands, ['chat', 'create', '--members', 'ada@contoso.com'], {
      routes: [[...me], ['POST', '/chats', () => ({ id: '19:abc@unq.gbl.spaces', chatType: 'oneOnOne' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls.map(c => `${c.method} ${c.path}`)).toEqual(['GET /me?$select=id', 'POST /chats']);
    expect(r.calls[1].body).toEqual({ chatType: 'oneOnOne', members: [member('ME'), member('ada@contoso.com')] });
    expect(JSON.parse(r.stdout).id).toBe('19:abc@unq.gbl.spaces');
  });

  it('two members → group with --topic', async () => {
    const r = await runCommand(registerTeamsCommands, ['chat', 'create', '--members', 'a@x.com, b@y.com', '--topic', 'Plan'], {
      routes: [[...me]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[1].body).toEqual({ chatType: 'group', topic: 'Plan', members: [member('ME'), member('a@x.com'), member('b@y.com')] });
  });

  it('--topic on a oneOnOne chat is a usage error', async () => {
    const r = await runCommand(registerTeamsCommands, ['chat', 'create', '--members', 'a@x.com', '--topic', 'Plan']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('invalid email is exit 2 with no request', async () => {
    const r = await runCommand(registerTeamsCommands, ['chat', 'create', '--members', 'a@x.com,bogus']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toContain('bogus');
  });

  it('dry-run previews without a resolved /me id', async () => {
    const r = await runCommand(registerTeamsCommands, ['chat', 'create', '--members', 'a@x.com'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain('[dry-run] POST https://graph.microsoft.com/v1.0/chats');
  });
});

describe('@mentions on chat send / send-channel-message (PM-9)', () => {
  const users = ['GET', /^\/users\//, (c: { path: string }) => {
    const upn = c.path.match(/^\/users\/([^?]+)/)![1];
    return upn.startsWith('ada') ? { id: 'U-ADA', displayName: 'Ada Lovelace' } : { id: 'U-BOB', displayName: 'Bob <B>' };
  }] as const;

  it('chat send appends <at> tags and mentions[] when the text has no @name', async () => {
    const r = await runCommand(registerTeamsCommands, ['chat', 'send', 'C1', '--body', 'see this & that', '--mention', 'ada@contoso.com', '--mention', 'bob@contoso.com'], {
      routes: [[...users]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls.map(c => `${c.method} ${c.path}`)).toEqual([
      'GET /users/ada@contoso.com?$select=id,displayName',
      'GET /users/bob@contoso.com?$select=id,displayName',
      'POST /chats/C1/messages',
    ]);
    const body = r.calls[2].body as any;
    expect(body.body).toEqual({ contentType: 'html', content: 'see this &amp; that <at id="0">Ada Lovelace</at> <at id="1">Bob &lt;B&gt;</at>' });
    expect(body.mentions).toEqual([
      { id: 0, mentionText: 'Ada Lovelace', mentioned: { user: { id: 'U-ADA', displayName: 'Ada Lovelace', userIdentityType: 'aadUser' } } },
      { id: 1, mentionText: 'Bob <B>', mentioned: { user: { id: 'U-BOB', displayName: 'Bob <B>', userIdentityType: 'aadUser' } } },
    ]);
  });

  it('replaces the first @displayName occurrence in the text', async () => {
    const r = await runCommand(registerTeamsCommands, ['chat', 'send', 'C1', '--body', 'hi @Ada Lovelace, and @Ada Lovelace again', '--mention', 'ada@contoso.com'], {
      routes: [[...users]],
    });
    expect((r.calls[1].body as any).body.content).toBe('hi <at id="0">Ada Lovelace</at>, and @Ada Lovelace again');
  });

  it('send-channel-message carries mentions too (with --html body untouched)', async () => {
    const r = await runCommand(registerTeamsCommands, ['teams', 'send-channel-message', TEAM, CHANNEL, '--body', '<b>@Ada Lovelace</b> look', '--html', '--mention', 'ada@contoso.com'], {
      routes: [[...users]],
    });
    expect(r.exitCode).toBe(0);
    const post = r.calls.find(c => c.method === 'POST')!;
    expect(post.path).toBe(`/teams/${TEAM}/channels/${CHANNEL}/messages`);
    expect((post.body as any).body).toEqual({ contentType: 'html', content: '<b><at id="0">Ada Lovelace</at></b> look' });
    expect((post.body as any).mentions).toHaveLength(1);
  });

  it('invalid --mention is exit 2', async () => {
    const r = await runCommand(registerTeamsCommands, ['chat', 'send', 'C1', '--body', 'x', '--mention', 'nope']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('no --mention keeps the plain text body', async () => {
    const r = await runCommand(registerTeamsCommands, ['chat', 'send', 'C1', '--body', 'x']);
    expect(r.calls[0].body).toEqual({ body: { contentType: 'text', content: 'x' } });
  });
});

describe('notes (G-8)', () => {
  it('list GETs /me/notes', async () => {
    const r = await runCommand(registerNotesCommands, ['notes', 'list', '--top', '3'], {
      routes: [['GET', '/me/notes', () => ({ value: [{ id: 'N1' }] })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].path).toBe('/me/notes?$top=3');
    expect(JSON.parse(r.stdout).value[0].id).toBe('N1');
  });

  it('get encodes the id', async () => {
    const r = await runCommand(registerNotesCommands, ['notes', 'get', 'A/B=']);
    expect(r.calls[0].url).toBe('https://graph.microsoft.com/v1.0/me/notes/A%2FB%3D');
  });

  it('create POSTs subject + text body', async () => {
    const r = await runCommand(registerNotesCommands, ['notes', 'create', '--body', 'remember', '--subject', 'Idea'], {
      routes: [['POST', '/me/notes', () => ({ id: 'N2' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toEqual({ subject: 'Idea', body: { contentType: 'text', content: 'remember' } });
    expect(JSON.parse(r.stdout).id).toBe('N2');
  });

  it('update PATCHes only given fields; nothing given is exit 2', async () => {
    const r = await runCommand(registerNotesCommands, ['notes', 'update', 'N1', '--subject', 'New']);
    expect(r.calls[0]).toMatchObject({ method: 'PATCH', path: '/me/notes/N1', body: { subject: 'New' } });
    const none = await runCommand(registerNotesCommands, ['notes', 'update', 'N1']);
    expect(none.exitCode).toBe(2);
    expect(none.calls).toHaveLength(0);
  });

  it('delete confirms', async () => {
    const refused = await runCommand(registerNotesCommands, ['notes', 'delete', 'N1']);
    expect(refused.exitCode).toBe(7);
    expect(refused.calls).toHaveLength(0);
    const ok = await runCommand(registerNotesCommands, ['notes', 'delete', 'N1'], { tty: true, answer: 'y' });
    expect(ok.calls[0]).toMatchObject({ method: 'DELETE', path: '/me/notes/N1' });
  });
});

describe('search query across entity types (PM-9)', () => {
  const hitsFor = (body: any) => {
    const types: string[] = body.requests[0].entityTypes;
    const mail = types.includes('message') || types.includes('event');
    return {
      value: [{
        hitsContainers: [{
          hits: mail
            ? [
                { hitId: 'M1', rank: 1, summary: 'mail hit', resource: { '@odata.type': '#microsoft.graph.message', subject: 'Budget', webLink: 'https://outlook/m1' } },
                { hitId: 'E1', rank: 2, summary: 'event hit', resource: { '@odata.type': '#microsoft.graph.event', subject: 'Review', webLink: 'https://outlook/e1' } },
              ]
            : [{ hitId: 'F1', rank: 1, summary: 'file hit', resource: { '@odata.type': '#microsoft.graph.driveItem', name: 'budget.xlsx', webUrl: 'https://sp/f1' } }],
        }],
      }],
    };
  };

  it('splits incompatible entity types into one request per group and merges hits', async () => {
    const r = await runCommand(registerSearchCommands, ['search', 'query', 'budget', '--entity', 'message,driveItem,event', '--size', '5', '--from', '10'], {
      routes: [['POST', '/search/query', c => hitsFor(c.body)]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(2);
    expect((r.calls[0].body as any).requests).toEqual([{ entityTypes: ['message', 'event'], query: { queryString: 'budget' }, from: 10, size: 5 }]);
    expect((r.calls[1].body as any).requests).toEqual([{ entityTypes: ['driveItem'], query: { queryString: 'budget' }, from: 10, size: 5 }]);
    expect(JSON.parse(r.stdout).value).toEqual([
      { entityType: 'message', id: 'M1', summary: 'mail hit', rank: 1, title: 'Budget', webUrl: 'https://outlook/m1' },
      { entityType: 'event', id: 'E1', summary: 'event hit', rank: 2, title: 'Review', webUrl: 'https://outlook/e1' },
      { entityType: 'driveItem', id: 'F1', summary: 'file hit', rank: 1, title: 'budget.xlsx', webUrl: 'https://sp/f1' },
    ]);
  });

  it('a single compatible group is one request', async () => {
    const r = await runCommand(registerSearchCommands, ['search', 'query', 'x', '--entity', 'driveItem,site,listItem'], {
      routes: [['POST', '/search/query', c => hitsFor(c.body)]],
    });
    expect(r.calls).toHaveLength(1);
    expect((r.calls[0].body as any).requests[0].entityTypes).toEqual(['driveItem', 'site', 'listItem']);
  });

  it('unknown entity type is exit 2', async () => {
    const r = await runCommand(registerSearchCommands, ['search', 'query', 'x', '--entity', 'message,widget']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('dry-run previews every request', async () => {
    const r = await runCommand(registerSearchCommands, ['search', 'query', 'x', '--entity', 'message,site'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stderr.match(/\[dry-run\] POST/g)).toHaveLength(2);
  });
});

describe('notes delete under --dry-run (DR-1)', () => {
  it('prints no "Note deleted." line, only the trailer', async () => {
    const r = await runCommand(registerNotesCommands, ['notes', 'delete', 'N1'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).not.toContain('Note deleted.');
    expect(r.stderr).toContain('request(s) previewed; nothing was sent.');
  });
});
