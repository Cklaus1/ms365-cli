import { describe, it, expect } from 'vitest';
import { registerMailCommands } from '../src/commands/mail.js';
import { registerCalendarCommands } from '../src/commands/calendar.js';
import { registerWatchCommands } from '../src/commands/watch.js';
import { registerTodoCommands } from '../src/commands/todo.js';
import { decodeCursor } from '../src/formatter.js';
import { GraphClient } from '../src/graph-client.js';
import { refResource, withImmutableIds } from '../src/refs.js';
import { runCommand, type RecordedCall } from './helpers/cli.js';
import { leafCommands, sampleArgv } from './helpers/commands.js';
import { fakeJwt } from './helpers/entry.js';

const IDS = 'IdType="ImmutableId"';
const REF_KEYS = ['provider', 'tenantId', 'accountId', 'mailbox', 'resource', 'id', 'version', 'lastModified', 'webLink', 'observedAt'];
const TENANT = '11111111-2222-3333-4444-555555555555';
const OID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
const jwt = fakeJwt({ aud: 'https://graph.microsoft.com', tid: TENANT, oid: OID, upn: 'ceo@contoso.com', exp: Math.floor(Date.now() / 1000) + 3600 });
const me = ['GET', /^\/me\?/, () => ({ id: 'me-id', userPrincipalName: 'ceo@contoso.com', mail: 'ceo@contoso.com' })] as const;
const query = (c: RecordedCall) => new URLSearchParams(c.url.split('?')[1] ?? '');

describe('immutable ids (AH-3, AC 10)', () => {
  it('asks for immutable ids on Outlook mail and calendar paths only, merged into an existing Prefer', () => {
    expect(withImmutableIds('https://graph.microsoft.com/v1.0/me/messages?$top=1', undefined)).toEqual({ Prefer: IDS });
    expect(withImmutableIds('/users/u%40x.com/calendarView/delta', { prefer: 'outlook.timezone="UTC"' })).toEqual({ Prefer: `outlook.timezone="UTC", ${IDS}` });
    expect(withImmutableIds('/me/mailFolders/inbox/messages/delta', { Prefer: IDS })).toEqual({ Prefer: IDS });
    for (const other of ['/me/chats/c1/messages', '/teams/t/channels/c/messages', '/me/drive/root/delta', '/me', '/me/todo/lists', '/me/mailboxSettings']) {
      expect(withImmutableIds(other, undefined), other).toBeUndefined();
    }
  });

  it('every request of every mail and calendar command (batch items included) that reaches Outlook asks for immutable ids', async () => {
    const outlook = /^\/(me|users\/[^/]+)\/(messages|mailFolders|events|calendar|calendars|calendarView|calendarGroups|sendMail)\b/;
    let checked = 0;
    for (const register of [registerMailCommands, registerCalendarCommands]) {
      for (const { chain, cmd } of leafCommands(register)) {
        const r = await runCommand(register, sampleArgv(chain, cmd), { tty: true, answer: 'y', routes: [['*', /.*/, () => ({ value: [] })]] });
        for (const c of r.calls) {
          const items = c.path.startsWith('/$batch') ? ((c.body as { requests: Array<{ url: string; headers?: Record<string, string> }> }).requests) : [{ url: c.path, headers: c.headers }];
          for (const item of items.filter(i => outlook.test(i.url))) {
            checked++;
            expect(item.headers?.Prefer ?? '', `${chain.join(' ')}: ${item.url}`).toContain(IDS);
          }
        }
      }
    }
    expect(checked).toBeGreaterThan(40);
  }, 30_000);
});

describe('--refs _ref (AH-3, AC 10)', () => {
  it('names messages and events, not attachments, folders or Teams messages', () => {
    expect(refResource('/me/messages')).toBe('message');
    expect(refResource('/me/mailFolders/inbox/messages/delta')).toBe('message');
    expect(refResource('/users/x/messages/M1')).toBe('message');
    expect(refResource('/me/calendarView')).toBe('event');
    expect(refResource('/me/events/E1/instances')).toBe('event');
    expect(refResource('/me/calendarView/delta')).toBe('event');
    for (const none of ['/me/messages/M1/attachments', '/me/mailFolders', '/me/calendars', '/me/chats/c/messages']) expect(refResource(none), none).toBeNull();
  });

  it('mail list --select subject --refs: $select widened, every item has every _ref field (opaque token → GET /me)', async () => {
    const r = await runCommand(registerMailCommands, ['mail', 'list', '--select', 'subject', '-o', 'json'], {
      refs: true,
      routes: [me, ['GET', '/me/messages', () => ({ value: [
        { id: 'M1', subject: 'a', '@odata.etag': 'W/"e1"' },
        { id: 'M2', subject: 'b', changeKey: 'ck2', lastModifiedDateTime: '2026-10-17T10:00:00Z', webLink: 'https://outlook.office365.com/owa/?ItemID=M2' },
      ] })]],
    });
    expect(r.exitCode, r.stderr).toBe(0);
    const list = r.calls.find(c => c.path.startsWith('/me/messages'))!;
    expect(query(list).get('$select')!.split(',')).toEqual(expect.arrayContaining(['subject', 'id', 'changeKey', 'lastModifiedDateTime', 'webLink']));
    const [a, b] = JSON.parse(r.stdout).value;
    for (const item of [a, b]) expect(Object.keys(item._ref).sort()).toEqual([...REF_KEYS].sort());
    expect(a._ref).toMatchObject({ provider: 'microsoft-graph', tenantId: null, accountId: 'me-id', mailbox: 'ceo@contoso.com', resource: 'message', id: 'M1', version: 'W/"e1"', webLink: null });
    expect(b._ref).toMatchObject({ id: 'M2', version: 'ck2', lastModified: '2026-10-17T10:00:00Z', webLink: 'https://outlook.office365.com/owa/?ItemID=M2' });
    expect(Number.isNaN(Date.parse(a._ref.observedAt))).toBe(false);
  });

  it('calendar view --refs reads tenant and account from the token (no GET /me)', async () => {
    const r = await runCommand(registerCalendarCommands, ['calendar', 'view', '--start', '2026-10-20', '--end', '2026-10-21', '-o', 'json'], {
      refs: true, token: jwt,
      routes: [['GET', '/me/calendarView', () => ({ value: [{ id: 'E1', changeKey: 'v1' }] })]],
    });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.calls.some(c => c.path.startsWith('/me?'))).toBe(false);
    expect(JSON.parse(r.stdout).value[0]._ref).toMatchObject({ tenantId: TENANT, accountId: OID, mailbox: 'ceo@contoso.com', resource: 'event', id: 'E1', version: 'v1' });
  });

  it('watch --refs puts _ref on each change line', async () => {
    const r = await runCommand(registerWatchCommands, ['watch', 'mail', '--include-existing', '--reset', '--name', 'refs-test'], {
      refs: true, token: jwt,
      routes: [
        ['GET', /^\/me\/mailFolders\/inbox\?/, () => ({ id: 'inbox' })],
        ['GET', /messages\/delta/, () => ({ value: [{ id: 'M1', changeKey: 'c1' }], '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=t' })],
      ],
    });
    expect(r.exitCode, r.stderr).toBe(0);
    const line = JSON.parse(r.stdout.trim().split('\n')[0]);
    expect(line._ref).toMatchObject({ resource: 'message', id: 'M1', version: 'c1', tenantId: TENANT });
    expect(line.item._ref).toBeUndefined();
  });

  it('without --refs nothing changes', async () => {
    const r = await runCommand(registerMailCommands, ['mail', 'list', '--select', 'subject', '-o', 'json'], { routes: [me, ['GET', '/me/messages', () => ({ value: [{ id: 'M1' }] })]] });
    expect(query(r.calls[0]).get('$select')).toBe('subject');
    expect(JSON.parse(r.stdout).value[0]._ref).toBeUndefined();
  });

  it('GraphClient.identity() is read once', async () => {
    let gets = 0;
    const client = new GraphClient({ getToken: async () => { gets++; return jwt; } } as never);
    await Promise.all([client.identity(), client.identity()]);
    expect(gets).toBe(1);
  });
});

describe('completeness in band (AH-3, AC 11)', () => {
  const base = 'https://graph.microsoft.com/v1.0/me/todo/lists/L1/tasks';
  const threePages = [['GET', /tasks/, (c: RecordedCall) => {
    const page = Number(new URL(c.url).searchParams.get('$skiptoken') ?? '1');
    return { value: [{ id: `t${page}` }], ...(page < 3 && { '@odata.nextLink': `${base}?$skiptoken=${page + 1}` }) };
  }]] as never;
  const tasks = (...extra: string[]) => ['todo', 'tasks', 'L1', ...extra];

  it('--all --page-limit 2 of 3 pages: complete false, page_limit, a nextCursor for page 3 (json and ndjson)', async () => {
    const json = await runCommand(registerTodoCommands, tasks('--all', '--page-limit', '2', '-o', 'json'), { routes: threePages });
    const out = JSON.parse(json.stdout);
    expect(out).toMatchObject({ complete: false, incompleteReason: 'page_limit' });
    expect(out.value.map((t: { id: string }) => t.id)).toEqual(['t1', 't2']);
    expect(decodeCursor(out.nextCursor)).toBe(`${base}?$skiptoken=3`);

    const nd = await runCommand(registerTodoCommands, tasks('--all', '--page-limit', '2', '-o', 'ndjson'), { routes: threePages });
    const lines = nd.stdout.trim().split('\n').map(l => JSON.parse(l));
    expect(lines.at(-1)).toEqual({ _meta: { complete: false, incompleteReason: 'page_limit', nextCursor: out.nextCursor } });

    const yaml = await runCommand(registerTodoCommands, tasks('--all', '--page-limit', '2', '-o', 'yaml'), { routes: threePages });
    expect(yaml.stdout).toContain('incompleteReason: page_limit');
  });

  it('a full read says complete: true', async () => {
    const r = await runCommand(registerTodoCommands, tasks('--all', '-o', 'json'), { routes: threePages });
    expect(JSON.parse(r.stdout)).toEqual({ value: [{ id: 't1' }, { id: 't2' }, { id: 't3' }], complete: true });
  });

  it('one page of several: next_cursor (json field, ndjson _meta line)', async () => {
    const json = JSON.parse((await runCommand(registerTodoCommands, tasks('-o', 'json'), { routes: threePages })).stdout);
    expect(json).toMatchObject({ complete: false, incompleteReason: 'next_cursor' });
    const nd = await runCommand(registerTodoCommands, tasks('-o', 'ndjson'), { routes: threePages });
    expect(JSON.parse(nd.stdout.trim().split('\n').at(-1)!)._meta).toEqual({ complete: false, incompleteReason: 'next_cursor', nextCursor: json.nextCursor });
    expect(nd.stderr).not.toContain('nextCursor:');
  });

  it('mail list --search x --all that returns 250 items says search_cap', async () => {
    const value = Array.from({ length: 250 }, (_, i) => ({ id: `m${i}` }));
    const r = await runCommand(registerMailCommands, ['mail', 'list', '--search', 'x', '--all', '-o', 'json'], { routes: [me, ['GET', '/me/messages', () => ({ value })]] });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(JSON.parse(r.stdout)).toMatchObject({ complete: false, incompleteReason: 'search_cap' });
    const few = await runCommand(registerMailCommands, ['mail', 'list', '--search', 'x', '--all', '-o', 'json'], { routes: [me, ['GET', '/me/messages', () => ({ value: value.slice(0, 3) })]] });
    expect(JSON.parse(few.stdout).complete).toBe(true);
  });
});
