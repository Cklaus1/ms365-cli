import { describe, it, expect } from 'vitest';
import { runCommand } from '../helpers/cli.js';
import { registerTeamsCommands } from '../../src/commands/teams.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';

describe('teams/chat send commands report success (M-32)', () => {
  it('chat send exits 0 and confirms on stderr', async () => {
    const r = await runCommand(registerTeamsCommands, ['chat', 'send', 'C1', '--body', 'hi'], {
      routes: [['POST', '/chats/C1/messages', () => ({ id: 'M9' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(1);
    expect(r.stderr).toContain('Message sent.');
    expect(r.stdout).toBe('');
  });

  it('send-channel-message exits 0', async () => {
    const r = await runCommand(registerTeamsCommands, ['teams', 'send-channel-message', '00000000-0000-0000-0000-0000000000a1', '19:ch1@thread.tacv2', '--body', 'hi'], {
      routes: [['POST', '/teams/00000000-0000-0000-0000-0000000000a1/channels/19:ch1@thread.tacv2/messages', () => ({ id: 'M9' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain('Channel message sent.');
  });

  it('reply-channel-message exits 0', async () => {
    const r = await runCommand(registerTeamsCommands, ['teams', 'reply-channel-message', '00000000-0000-0000-0000-0000000000a1', '19:ch1@thread.tacv2', 'M1', '--body', 'hi'], {
      routes: [['POST', '/teams/00000000-0000-0000-0000-0000000000a1/channels/19:ch1@thread.tacv2/messages/M1/replies', () => ({ id: 'R1' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].path).toBe('/teams/00000000-0000-0000-0000-0000000000a1/channels/19:ch1@thread.tacv2/messages/M1/replies');
  });
});

describe('chat reply posts a quoted reply (M-6, D10)', () => {
  const from = { user: { id: 'U1', displayName: 'Ada', userIdentityType: 'aadUser' } };
  const original = { id: 'M1', body: { contentType: 'html', content: '<p>Lunch &amp; <b>learn</b>?</p>' }, from };

  it('GETs the original then POSTs to /chats/{id}/messages with a messageReference attachment', async () => {
    const r = await runCommand(registerTeamsCommands, ['chat', 'reply', 'C1', 'M1', '--body', 'yes <ok>'], {
      routes: [
        ['GET', '/chats/C1/messages/M1', () => original],
        ['POST', '/chats/C1/messages', () => ({ id: 'M2' })],
      ],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls.map((c) => `${c.method} ${c.path}`)).toEqual(['GET /chats/C1/messages/M1', 'POST /chats/C1/messages']);
    const body = r.calls[1].body as any;
    expect(body.body.contentType).toBe('html');
    expect(body.body.content).toBe('<attachment id="M1"></attachment>yes &lt;ok&gt;');
    expect(body.attachments).toHaveLength(1);
    expect(body.attachments[0]).toMatchObject({ id: 'M1', contentType: 'messageReference' });
    expect(JSON.parse(body.attachments[0].content)).toEqual({ messageId: 'M1', messagePreview: 'Lunch & learn?', messageSender: from });
    expect(r.stderr).toContain('Reply sent.');
  });

  it('passes --html bodies through unescaped', async () => {
    const r = await runCommand(registerTeamsCommands, ['chat', 'reply', 'C1', 'M1', '--body', '<b>yes</b>', '--html'], {
      routes: [
        ['GET', '/chats/C1/messages/M1', () => original],
        ['POST', '/chats/C1/messages', () => ({ id: 'M2' })],
      ],
    });
    expect(r.exitCode).toBe(0);
    expect((r.calls[1].body as any).body.content).toBe('<attachment id="M1"></attachment><b>yes</b>');
  });

  it('never calls the nonexistent /replies endpoint', async () => {
    const r = await runCommand(registerTeamsCommands, ['chat', 'reply', 'C1', 'M1', '--body', 'x'], {
      routes: [
        ['GET', '/chats/C1/messages/M1', () => original],
        ['POST', '/chats/C1/messages', () => ({ id: 'M2' })],
      ],
    });
    expect(r.calls.some((c) => c.path.includes('/replies'))).toBe(false);
  });

  it('fails with not-found (exit 5) and posts nothing when the original is missing', async () => {
    const r = await runCommand(registerTeamsCommands, ['chat', 'reply', 'C1', 'M1', '--body', 'x'], {
      routes: [['GET', '/chats/C1/messages/M1', () => new Response('{"error":{"message":"gone"}}', { status: 404 })]],
    });
    expect(r.exitCode).toBe(5);
    expect(r.calls.filter((c) => c.method === 'POST')).toHaveLength(0);
  });
});

describe('chat list --all honours --page-limit (M-26)', () => {
  it('stops after the page limit', async () => {
    let n = 0;
    const r = await runCommand(registerTeamsCommands, ['chat', 'list', '--all', '--page-limit', '2', '--page-delay', '0', '-o', 'json'], {
      routes: [['GET', '/me/chats', () => ({ value: [{ id: `c${++n}` }], '@odata.nextLink': `${GRAPH}/me/chats?$skiptoken=${n}` })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(2);
    expect(JSON.parse(r.stdout).value).toHaveLength(2);
  });
});
