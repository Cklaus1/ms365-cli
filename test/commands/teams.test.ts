import { describe, it, expect } from 'vitest';
import { runCli } from '../e2e/run-cli.js';

const ENV = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_TOKEN: 'fake' };
const GRAPH = 'https://graph.microsoft.com/v1.0';

const SENDS: Array<{ args: string[]; path: string }> = [
  { args: ['teams', 'send-channel-message', 'T1', 'C1', '--body', 'hi'], path: '/teams/T1/channels/C1/messages' },
  { args: ['teams', 'reply-channel-message', 'T1', 'C1', 'M1', '--body', 'hi'], path: '/teams/T1/channels/C1/messages/M1/replies' },
  { args: ['chat', 'send', 'CH1', '--body', 'hi'], path: '/chats/CH1/messages' },
  { args: ['chat', 'reply', 'CH1', 'M1', '--body', 'hi'], path: '/chats/CH1/messages/replyWithQuote' },
];

describe('TEAMS-1 / OUT-1 send commands print the created message', () => {
  for (const { args, path } of SENDS) {
    it(`${args.slice(0, 2).join(' ')} exits 0 and prints the created message on stdout`, async () => {
      const r = await runCli([...args, '-o', 'json'], {
        env: ENV,
        routes: [{ match: `${path}$`, method: 'POST', status: 201, body: { id: 'm1', body: { content: 'hi' } } }],
      });
      expect(r.stderr).not.toContain('ReferenceError');
      expect(r.code).toBe(0);
      expect(r.requests).toHaveLength(1);
      expect(r.requests[0].method).toBe('POST');
      expect(r.requests[0].url).toBe(`${GRAPH}${path}`);
      expect(JSON.parse(r.stdout)).toMatchObject({ id: 'm1' });
    });

    it(`${args.slice(0, 2).join(' ')} --dry-run sends nothing and prints the dry-run object`, async () => {
      const r = await runCli(['--dry-run', ...args, '-o', 'json'], { env: ENV, routes: [] });
      expect(r.code).toBe(0);
      expect(r.requests).toHaveLength(0);
      expect(r.stderr).toContain(`[dry-run] POST ${GRAPH}${path}`);
      expect(r.stderr).not.toMatch(/sent/i);
      expect(JSON.parse(r.stdout)).toMatchObject({ dryRun: true, method: 'POST', url: `${GRAPH}${path}` });
    });
  }
});

describe('TEAMS-2 chat reply uses replyWithQuote', () => {
  it('POSTs {messageIds, replyMessage} to /chats/{id}/messages/replyWithQuote', async () => {
    const r = await runCli(['chat', 'reply', '19:abc@thread.v2', 'M 1', '--body', '<b>ok</b>', '--html'], {
      env: ENV,
      routes: [{ match: '/messages/replyWithQuote$', method: 'POST', status: 201, body: { id: 'r1' } }],
    });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(1);
    const req = r.requests[0];
    expect(req.method).toBe('POST');
    expect(req.url).toBe(`${GRAPH}/chats/19%3Aabc%40thread.v2/messages/replyWithQuote`);
    expect(JSON.parse(req.body!)).toEqual({
      messageIds: ['M 1'],
      replyMessage: { body: { contentType: 'html', content: '<b>ok</b>' } },
    });
  });

  it('dry-run previews the replyWithQuote URL and body', async () => {
    const r = await runCli(['--dry-run', 'chat', 'reply', 'CH1', 'M1', '--body', 'hi'], { env: ENV, routes: [] });
    expect(r.code).toBe(0);
    expect(r.stderr).toContain(`[dry-run] POST ${GRAPH}/chats/CH1/messages/replyWithQuote`);
    expect(r.stderr).toContain('"messageIds":["M1"]');
    expect(r.stderr).toContain('"replyMessage":{"body":{"contentType":"text","content":"hi"}}');
  });
});

describe('PAGE-1 chat list', () => {
  it('honours --page-limit with --all and prints one JSON array', async () => {
    const r = await runCli(['chat', 'list', '--all', '--top', '1', '--page-limit', '2', '-o', 'json'], {
      env: ENV,
      routes: [
        { match: 'skiptoken=2', body: { value: [{ id: 'c2' }], '@odata.nextLink': `${GRAPH}/me/chats?$skiptoken=3` } },
        { match: 'skiptoken=3', body: { value: [{ id: 'c3' }] } },
        { match: '/me/chats', body: { value: [{ id: 'c1' }], '@odata.nextLink': `${GRAPH}/me/chats?$skiptoken=2` } },
      ],
    });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(2);
    expect(new URL(r.requests[0].url).searchParams.get('$top')).toBe('1');
    expect(JSON.parse(r.stdout).map((c: { id: string }) => c.id)).toEqual(['c1', 'c2']);
  });

  it('empty result is valid JSON', async () => {
    const r = await runCli(['chat', 'list', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/me/chats', body: { value: [] } }],
    });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout)).toEqual({ value: [] });
  });

  it('rejects a non-integer --page-limit', async () => {
    const r = await runCli(['chat', 'list', '--all', '--page-limit', 'x'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });
});

describe('PAGE-2 teams list/channels drain every page', () => {
  const DRAINS: Array<{ args: string[]; path: string }> = [
    { args: ['teams', 'list'], path: '/me/joinedTeams' },
    { args: ['teams', 'channels', 'T1'], path: '/teams/T1/channels' },
  ];
  for (const { args, path } of DRAINS) {
    it(`${args.slice(0, 2).join(' ')} follows nextLink and prints one combined JSON array`, async () => {
      const r = await runCli([...args, '-o', 'json'], {
        env: ENV,
        routes: [
          { match: 'skiptoken=2', body: { value: [{ id: 'x2' }] } },
          { match: `${path}$`, body: { value: [{ id: 'x1' }], '@odata.nextLink': `${GRAPH}${path}?$skiptoken=2` } },
        ],
      });
      expect(r.code).toBe(0);
      expect(r.requests).toHaveLength(2);
      expect(JSON.parse(r.stdout).map((t: { id: string }) => t.id)).toEqual(['x1', 'x2']);
    });
  }
});

describe('PAGE-2 teams members pagination options', () => {
  const routes = [
    { match: 'skiptoken=2', body: { value: [{ id: 'u2' }] } },
    { match: '/teams/T1/members', body: { value: [{ id: 'u1' }], '@odata.nextLink': `${GRAPH}/teams/T1/members?$skiptoken=2` } },
  ];

  it('--top sets $top and fetches one page', async () => {
    const r = await runCli(['teams', 'members', 'T1', '--top', '1', '-o', 'json'], { env: ENV, routes });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(1);
    expect(new URL(r.requests[0].url).searchParams.get('$top')).toBe('1');
    expect(JSON.parse(r.stdout).value).toEqual([{ id: 'u1' }]);
  });

  it('--all follows nextLink and prints one JSON array', async () => {
    const r = await runCli(['teams', 'members', 'T1', '--all', '-o', 'json'], { env: ENV, routes });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(2);
    expect(JSON.parse(r.stdout).map((u: { id: string }) => u.id)).toEqual(['u1', 'u2']);
  });
});
