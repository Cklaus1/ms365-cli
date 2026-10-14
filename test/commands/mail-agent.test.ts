import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { runCommand, type Route } from '../helpers/cli.js';
import { registerMailCommands } from '../../src/commands/mail.js';
import { registerTeamsCommands } from '../../src/commands/teams.js';

const run = (argv: string[], opts: Parameters<typeof runCommand>[2] = {}) =>
  runCommand(registerMailCommands, ['mail', ...argv], opts);
const seq = (r: { calls: Array<{ method: string; path: string }> }) =>
  r.calls.map(c => `${c.method} ${c.path.split('?')[0]}`);

// cwd is nested inside `root` so relative paths and any traversal stay inside the cleanup root.
let root: string;
let tmp: string;
let origCwd: string;
beforeEach(() => {
  origCwd = process.cwd();
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'mail-agent-test-'));
  tmp = path.join(root, 'a', 'b');
  fs.mkdirSync(tmp, { recursive: true });
  process.chdir(tmp);
});
afterEach(() => {
  process.chdir(origCwd);
  fs.rmSync(root, { recursive: true, force: true });
});

const MB = 1024 * 1024;
const uploadUrl = 'https://outlook.office.com/api/v2.0/AttachmentSessions(\'s\')?authtoken=t';
const uploadRoutes = (messageId: string): Route[] => [
  ['POST', `/me/messages/${messageId}/attachments/createUploadSession`, () => ({ uploadUrl })],
  ['PUT', /AttachmentSessions/, () => new Response('{}', { status: 200 })],
];

describe('--attach (PM-6): inline for small files', () => {
  it('mail send embeds a repeatable --attach as fileAttachments in sendMail', async () => {
    fs.writeFileSync('a.pdf', 'pdf!');
    fs.writeFileSync('notes.txt', 'hi');
    const r = await run(['send', '--to', 'x@contoso.com', '--subject', 's', '--body', 'b', '--attach', 'a.pdf', '--attach', 'notes.txt']);
    expect(r.exitCode).toBe(0);
    expect(seq(r)).toEqual(['POST /me/sendMail']);
    const msg = (r.calls[0].body as any).message;
    expect(msg.attachments).toEqual([
      { '@odata.type': '#microsoft.graph.fileAttachment', name: 'a.pdf', contentType: 'application/pdf', contentBytes: Buffer.from('pdf!').toString('base64') },
      { '@odata.type': '#microsoft.graph.fileAttachment', name: 'notes.txt', contentType: 'text/plain', contentBytes: Buffer.from('hi').toString('base64') },
    ]);
  });

  it('mail draft includes attachments in the created message', async () => {
    fs.writeFileSync('a.txt', 'x');
    const r = await run(['draft', '--to', 'x@contoso.com', '--subject', 's', '--body', 'b', '--attach', 'a.txt', '-o', 'json'], {
      routes: [['POST', '/me/messages', () => ({ id: 'D1' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(seq(r)).toEqual(['POST /me/messages']);
    expect((r.calls[0].body as any).attachments[0]).toMatchObject({ name: 'a.txt', contentBytes: 'eA==' });
  });

  it('mail reply sends attachments in the message property', async () => {
    fs.writeFileSync('a.txt', 'x');
    const r = await run(['reply', 'M1', '--body', 'thanks', '--attach', 'a.txt']);
    expect(r.exitCode).toBe(0);
    expect(seq(r)).toEqual(['POST /me/messages/M1/reply']);
    expect(r.calls[0].body).toMatchObject({ comment: 'thanks', message: { attachments: [{ name: 'a.txt' }] } });
  });

  it('mail forward sends attachments in the message property', async () => {
    fs.writeFileSync('a.txt', 'x');
    const r = await run(['forward', 'M1', '--to', 'y@contoso.com', '--attach', 'a.txt']);
    expect(r.exitCode).toBe(0);
    expect(seq(r)).toEqual(['POST /me/messages/M1/forward']);
    expect(r.calls[0].body).toMatchObject({ toRecipients: [{ emailAddress: { address: 'y@contoso.com' } }], message: { attachments: [{ name: 'a.txt' }] } });
  });

  it('without --attach the request bodies are unchanged', async () => {
    const r = await run(['reply', 'M1', '--body', 'thanks']);
    expect(r.calls[0].body).toEqual({ comment: 'thanks' });
  });
});

describe('--attach (PM-6): draft + upload session above 3 MB', () => {
  it('mail send with a >3 MB file creates a draft, uploads, then sends the draft', async () => {
    fs.writeFileSync('big.bin', Buffer.alloc(3 * MB + 1, 1));
    fs.writeFileSync('small.txt', 'abc');
    const r = await run(['send', '--to', 'x@contoso.com', '--subject', 's', '--body', 'b', '--attach', 'big.bin', '--attach', 'small.txt'], {
      routes: [...uploadRoutes('D1'), ['POST', '/me/messages/D1/attachments', () => ({ id: 'A1' })], ['POST', '/me/messages/D1/send', () => new Response(null, { status: 202 })], ['POST', '/me/messages', () => ({ id: 'D1' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(seq(r)).toEqual([
      'POST /me/messages',
      'POST /me/messages/D1/attachments/createUploadSession',
      `PUT ${uploadUrl.split('?')[0]}`,
      `PUT ${uploadUrl.split('?')[0]}`,
      'POST /me/messages/D1/attachments',
      'POST /me/messages/D1/send',
    ]);
    const draft = r.calls[0].body as any;
    expect(draft.subject).toBe('s');
    expect(draft.attachments).toBeUndefined();
    expect((r.calls[1].body as any).AttachmentItem).toMatchObject({ attachmentType: 'file', name: 'big.bin', size: 3 * MB + 1, contentType: 'application/octet-stream' });
    expect(r.calls[4].body).toMatchObject({ name: 'small.txt', contentBytes: Buffer.from('abc').toString('base64') });
    expect(r.calls[2].headers.Authorization).toBeUndefined();
  });

  it('small files whose total exceeds 3 MB also go through a draft', async () => {
    fs.writeFileSync('a.bin', Buffer.alloc(2 * MB, 1));
    fs.writeFileSync('b.bin', Buffer.alloc(2 * MB, 1));
    const r = await run(['send', '--to', 'x@contoso.com', '--subject', 's', '--body', 'b', '--attach', 'a.bin', '--attach', 'b.bin'], {
      routes: [['POST', '/me/messages/D1/', () => ({})], ['POST', '/me/messages', () => ({ id: 'D1' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(seq(r)).toEqual(['POST /me/messages', 'POST /me/messages/D1/attachments', 'POST /me/messages/D1/attachments', 'POST /me/messages/D1/send']);
  });

  it('mail draft with a large file uploads but does not send', async () => {
    fs.writeFileSync('big.bin', Buffer.alloc(3 * MB + 1, 1));
    const r = await run(['draft', '--to', 'x@contoso.com', '--subject', 's', '--body', 'b', '--attach', 'big.bin', '-o', 'json'], {
      routes: [...uploadRoutes('D1'), ['POST', '/me/messages', () => ({ id: 'D1' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(seq(r)).not.toContain('POST /me/messages/D1/send');
    expect(seq(r)[1]).toBe('POST /me/messages/D1/attachments/createUploadSession');
    expect(JSON.parse(r.stdout)).toMatchObject({ id: 'D1' });
  });

  it('mail reply --reply-all with a large file uses createReplyAll then send', async () => {
    fs.writeFileSync('big.bin', Buffer.alloc(3 * MB + 1, 1));
    const r = await run(['reply', 'M1', '--body', 'see attached', '--reply-all', '--attach', 'big.bin'], {
      routes: [...uploadRoutes('D2'), ['POST', '/me/messages/M1/createReplyAll', () => ({ id: 'D2' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(seq(r)[0]).toBe('POST /me/messages/M1/createReplyAll');
    expect(r.calls[0].body).toEqual({ comment: 'see attached' });
    expect(seq(r).at(-1)).toBe('POST /me/messages/D2/send');
  });

  it('mail forward with a large file uses createForward then send', async () => {
    fs.writeFileSync('big.bin', Buffer.alloc(3 * MB + 1, 1));
    const r = await run(['forward', 'M1', '--to', 'y@contoso.com', '--comment', 'fyi', '--attach', 'big.bin'], {
      routes: [...uploadRoutes('D3'), ['POST', '/me/messages/M1/createForward', () => ({ id: 'D3' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(seq(r)[0]).toBe('POST /me/messages/M1/createForward');
    expect(r.calls[0].body).toMatchObject({ comment: 'fyi', toRecipients: [{ emailAddress: { address: 'y@contoso.com' } }] });
    expect(seq(r).at(-1)).toBe('POST /me/messages/D3/send');
  });
});

describe('--attach (PM-6): validation', () => {
  for (const argv of [
    ['send', '--to', 'x@contoso.com', '--subject', 's', '--body', 'b'],
    ['draft', '--to', 'x@contoso.com', '--subject', 's', '--body', 'b'],
    ['reply', 'M1', '--body', 'b'],
    ['forward', 'M1', '--to', 'x@contoso.com'],
  ]) {
    it(`mail ${argv[0]}: a missing file exits 2 before any request`, async () => {
      fs.writeFileSync('ok.txt', 'x');
      const r = await run([...argv, '--attach', 'ok.txt', '--attach', 'nope.pdf']);
      expect(r.exitCode).toBe(2);
      expect(r.calls).toHaveLength(0);
      expect(r.stderr).toContain('nope.pdf');
    });
  }

  it('a directory is rejected as a usage error', async () => {
    fs.mkdirSync('dir');
    const r = await run(['send', '--to', 'x@contoso.com', '--subject', 's', '--body', 'b', '--attach', 'dir']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});

const me: Route = ['GET', '/me?', () => ({ mail: 'amy@contoso.com', userPrincipalName: 'amy@contoso.com' })];
const mailMsg = (id: string, address: string) => ({
  id, subject: 's', bodyPreview: 'Hello there friend',
  from: { emailAddress: { address } },
  body: { contentType: 'html', content: '<p>Hello <b>there</b> friend</p>' },
});

describe('mail get: --body, --unique, --max-chars, untrusted marking (§6)', () => {
  it('default keeps the HTML body and marks the message untrusted with sender.external', async () => {
    const r = await run(['get', 'M1', '-o', 'json'], { routes: [me, ['GET', '/me/messages/M1', () => mailMsg('M1', 'bob@fabrikam.com')]] });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].headers.Prefer).toBeUndefined();
    const out = JSON.parse(r.stdout);
    expect(out.body.content).toBe('<p>Hello <b>there</b> friend</p>');
    expect(out.untrusted).toBe(true);
    expect(out.sender).toEqual({ external: true });
    expect(r.calls.filter(c => c.path.startsWith('/me?'))).toHaveLength(1);
    expect(r.calls.find(c => c.path.startsWith('/me?'))!.path).toBe('/me?$select=mail,userPrincipalName');
  });

  it('internal senders are external:false', async () => {
    const r = await run(['get', 'M1', '-o', 'json'], { routes: [me, ['GET', '/me/messages/M1', () => mailMsg('M1', 'joe@contoso.com')]] });
    expect(JSON.parse(r.stdout).sender).toEqual({ external: false });
  });

  it('--body text sends the outlook.body-content-type Prefer header', async () => {
    const r = await run(['get', 'M1', '--body', 'text', '-o', 'json'], {
      routes: [me, ['GET', '/me/messages/M1', () => ({ ...mailMsg('M1', 'a@contoso.com'), body: { contentType: 'text', content: 'Hello there friend' } })]],
    });
    expect(r.calls.find(c => c.path.startsWith('/me/messages/M1'))!.headers.Prefer).toBe('outlook.body-content-type="text"');
    expect(JSON.parse(r.stdout).body).toEqual({ contentType: 'text', content: 'Hello there friend' });
  });

  it('--body markdown converts the HTML body locally', async () => {
    const r = await run(['get', 'M1', '--body', 'markdown', '-o', 'json'], { routes: [me, ['GET', '/me/messages/M1', () => mailMsg('M1', 'a@contoso.com')]] });
    expect(r.calls.find(c => c.path.startsWith('/me/messages/M1'))!.headers.Prefer).toBeUndefined();
    expect(JSON.parse(r.stdout).body).toEqual({ contentType: 'markdown', content: 'Hello **there** friend' });
  });

  it('--body none drops the body', async () => {
    const r = await run(['get', 'M1', '--body', 'none', '-o', 'json'], { routes: [me, ['GET', '/me/messages/M1', () => mailMsg('M1', 'a@contoso.com')]] });
    expect(JSON.parse(r.stdout).body).toBeUndefined();
  });

  it('rejects an unknown --body mode with exit 2', async () => {
    const r = await run(['get', 'M1', '--body', 'pdf']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('rejects a non-positive --max-chars with exit 2', async () => {
    const r = await run(['get', 'M1', '--max-chars', '0']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('--unique selects uniqueBody instead of body', async () => {
    const r = await run(['get', 'M1', '--unique', '-o', 'json'], {
      routes: [me, ['GET', '/me/messages/M1', () => ({ id: 'M1', uniqueBody: { contentType: 'html', content: '<p>new</p>' } })]],
    });
    const select = new URL(r.calls.find(c => c.path.startsWith('/me/messages/M1'))!.url).searchParams.get('$select')!.split(',');
    expect(select).toContain('uniqueBody');
    expect(select).not.toContain('body');
    expect(JSON.parse(r.stdout).uniqueBody.content).toBe('<p>new</p>');
  });

  it('--unique with --select swaps body for uniqueBody', async () => {
    const r = await run(['get', 'M1', '--unique', '--select', 'subject,body', '-o', 'json'], { routes: [me] });
    expect(new URL(r.calls.find(c => c.path.startsWith('/me/messages/M1'))!.url).searchParams.get('$select')).toBe('subject,uniqueBody');
  });

  it('--max-chars and --wrap-untrusted shape the body', async () => {
    const r = await run(['get', 'M1', '--body', 'markdown', '--max-chars', '5', '--wrap-untrusted', '-o', 'json'], {
      routes: [me, ['GET', '/me/messages/M1', () => mailMsg('M1', 'a@contoso.com')]],
    });
    const out = JSON.parse(r.stdout);
    expect(out.body.content).toBe('<<<UNTRUSTED_CONTENT\nHello…[truncated 17 chars]\nUNTRUSTED_CONTENT>>>');
    expect(out.bodyPreview).toBe('<<<UNTRUSTED_CONTENT\nHello…[truncated 13 chars]\nUNTRUSTED_CONTENT>>>');
  });

  it('still prints the message (without external) when /me fails', async () => {
    const r = await run(['get', 'M1', '-o', 'json'], {
      routes: [['GET', '/me?', () => new Response(JSON.stringify({ error: { code: 'Forbidden', message: 'no' } }), { status: 403 })], ['GET', '/me/messages/M1', () => mailMsg('M1', 'a@fabrikam.com')]],
    });
    expect(r.exitCode).toBe(0);
    const out = JSON.parse(r.stdout);
    expect(out.untrusted).toBe(true);
    expect(out.sender).toBeUndefined();
  });

  it('does not look up /me under --dry-run', async () => {
    const r = await run(['get', 'M1', '--body', 'text'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).not.toContain('/me?');
    expect(r.stderr).toContain('[dry-run] GET https://graph.microsoft.com/v1.0/me/messages/M1');
  });
});

describe('mail list / triage: bodies and marking (§6)', () => {
  it('marks every listed message and truncates bodies; /me is fetched once', async () => {
    const r = await run(['list', '--max-chars', '5', '-o', 'json'], {
      routes: [me, ['GET', '/me/messages', () => ({ value: [mailMsg('1', 'a@contoso.com'), mailMsg('2', 'b@fabrikam.com')] })]],
    });
    expect(r.exitCode).toBe(0);
    const out = JSON.parse(r.stdout);
    expect(out.value.map((m: any) => m.sender.external)).toEqual([false, true]);
    expect(out.value.every((m: any) => m.untrusted === true)).toBe(true);
    expect(out.value[0].body.content).toBe('<p>He…[truncated 27 chars]');
    expect(r.calls.filter(c => c.path.startsWith('/me?'))).toHaveLength(1);
  });

  it('--all streams marked items across pages with one /me lookup', async () => {
    const next = 'https://graph.microsoft.com/v1.0/me/messages?$skiptoken=2';
    const r = await run(['list', '--all', '--body', 'text', '-o', 'ndjson'], {
      routes: [
        me,
        ['GET', '/me/messages?$skiptoken=2', () => ({ value: [mailMsg('2', 'b@fabrikam.com')] })],
        ['GET', '/me/messages', () => ({ value: [mailMsg('1', 'a@contoso.com')], '@odata.nextLink': next })],
      ],
    });
    expect(r.exitCode).toBe(0);
    const lines = r.stdout.trim().split('\n').map(l => JSON.parse(l));
    expect(lines.map(l => [l.id, l.untrusted, l.sender.external])).toEqual([['1', true, false], ['2', true, true]]);
    expect(lines[0].body.content).toBe('Hello there friend');
    const pageCalls = r.calls.filter(c => c.path.startsWith('/me/messages'));
    expect(pageCalls.every(c => c.headers.Prefer === 'outlook.body-content-type="text"')).toBe(true);
    expect(r.calls.filter(c => c.path.startsWith('/me?'))).toHaveLength(1);
  });

  it('an empty page needs no /me lookup', async () => {
    const r = await run(['list', '-o', 'json'], { routes: [me, ['GET', '/me/messages', () => ({ value: [] })]] });
    expect(r.calls.filter(c => c.path.startsWith('/me?'))).toHaveLength(0);
  });

  it('triage -o json marks messages and truncates bodyPreview', async () => {
    const r = await run(['triage', '--max-chars', '5', '-o', 'json'], {
      routes: [me, ['GET', '/me/messages', () => ({ value: [{ id: '1', bodyPreview: 'Hello there', from: { emailAddress: { address: 'x@fabrikam.com' } } }] })]],
    });
    const out = JSON.parse(r.stdout);
    expect(out.value[0]).toMatchObject({ untrusted: true, sender: { external: true }, bodyPreview: 'Hello…[truncated 6 chars]' });
  });
});

const runTeams = (argv: string[], opts: Parameters<typeof runCommand>[2] = {}) => runCommand(registerTeamsCommands, argv, opts);
const org: Route = ['GET', '/organization', () => ({ value: [{ id: 'T1' }] })];
const teamsMsg = (id: string, tenantId: string) => ({
  id, from: { user: { id: 'u', displayName: 'U', tenantId } },
  body: { contentType: 'html', content: '<div>Hi <b>all</b><br>see <a href="https://x.test">this</a></div>' },
});

describe('Teams messages: --body conversion and untrusted marking (§6)', () => {
  it('chat messages --body text converts HTML locally and flags other tenants as external', async () => {
    const r = await runTeams(['chat', 'messages', 'C1', '--body', 'text', '-o', 'json'], {
      routes: [org, ['GET', '/chats/C1/messages', () => ({ value: [teamsMsg('1', 'T1'), teamsMsg('2', 'T9')] })]],
    });
    expect(r.exitCode).toBe(0);
    const out = JSON.parse(r.stdout);
    expect(out.value[0].body).toEqual({ contentType: 'text', content: 'Hi all\nsee this (https://x.test)' });
    expect(out.value.map((m: any) => [m.untrusted, m.sender.external])).toEqual([[true, false], [true, true]]);
    expect(r.calls.find(c => c.path.startsWith('/organization'))!.path).toBe('/organization?$select=id');
    expect(r.calls.find(c => c.path.startsWith('/chats'))!.headers.Prefer).toBeUndefined();
  });

  it('channel-messages --body markdown --wrap-untrusted', async () => {
    const r = await runTeams(['teams', 'channel-messages', 'T', 'CH', '--body', 'markdown', '--wrap-untrusted', '-o', 'json'], {
      routes: [org, ['GET', '/teams/T/channels/CH/messages', () => ({ value: [teamsMsg('1', 'T1')] })]],
    });
    expect(JSON.parse(r.stdout).value[0].body.content)
      .toBe('<<<UNTRUSTED_CONTENT\nHi **all**\nsee [this](https://x.test)\nUNTRUSTED_CONTENT>>>');
  });

  it('omits external when the signed-in tenant is unknown', async () => {
    const r = await runTeams(['chat', 'messages', 'C1', '-o', 'json'], {
      routes: [['GET', '/organization', () => new Response('{}', { status: 403 })], ['GET', '/chats/C1/messages', () => ({ value: [teamsMsg('1', 'T1')] })]],
    });
    const m = JSON.parse(r.stdout).value[0];
    expect(m.untrusted).toBe(true);
    expect(m.sender).toBeUndefined();
    expect(m.body.content).toContain('<b>all</b>');
  });

  it('--body none and --max-chars apply to chat messages', async () => {
    const none = await runTeams(['chat', 'messages', 'C1', '--body', 'none', '-o', 'json'], {
      routes: [org, ['GET', '/chats/C1/messages', () => ({ value: [teamsMsg('1', 'T1')] })]],
    });
    expect(JSON.parse(none.stdout).value[0].body).toBeUndefined();
    const cut = await runTeams(['chat', 'messages', 'C1', '--body', 'text', '--max-chars', '2', '-o', 'json'], {
      routes: [org, ['GET', '/chats/C1/messages', () => ({ value: [teamsMsg('1', 'T1')] })]],
    });
    expect(JSON.parse(cut.stdout).value[0].body.content).toBe('Hi…[truncated 30 chars]');
  });
});
