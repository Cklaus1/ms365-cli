import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { runCommand, type Route } from '../helpers/cli.js';
import { registerMailCommands } from '../../src/commands/mail.js';
import { setErrorFormat } from '../../src/errors.js';

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
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'mail-followup-test-'));
  tmp = path.join(root, 'a', 'b');
  fs.mkdirSync(tmp, { recursive: true });
  process.chdir(tmp);
});
afterEach(() => {
  process.chdir(origCwd);
  fs.rmSync(root, { recursive: true, force: true });
  setErrorFormat('text');
});

const MB = 1024 * 1024;
const uploadUrl = 'https://outlook.office.com/api/v2.0/AttachmentSessions(\'s\')?authtoken=t';
const json = (data: unknown, status = 200) => new Response(JSON.stringify(data), { status, headers: { 'content-type': 'application/json' } });
const graphError = (status: number) => json({ error: { code: 'X', message: `failed ${status}` } }, status);
let keyN = 0;
const newKey = () => `followup-${process.pid}-${Date.now()}-${keyN++}`;

describe('RT-2: mail attachments use the shared upload session (R-1)', () => {
  it('add-attachment --file audits the upload as one PUT on the attachments endpoint', async () => {
    fs.writeFileSync('big.bin', Buffer.alloc(3 * MB + 1, 1));
    const cfg = path.join(root, 'cfg');
    const r = await run(['add-attachment', 'M', '--file', 'big.bin', '--content-type', 'application/octet-stream', '-o', 'json'], {
      env: { MS365_CLI_CONFIG_DIR: cfg, MS365_CLI_AUDIT: '1' },
      routes: [
        ['POST', '/me/messages/M/attachments/createUploadSession', () => ({ uploadUrl })],
        ['PUT', /AttachmentSessions/, () => new Response('{}', { status: 200 })],
      ],
    });
    expect(r.exitCode).toBe(0);
    const lines = fs.readFileSync(path.join(cfg, 'audit.log'), 'utf8').trim().split('\n').map(l => JSON.parse(l));
    expect(lines.filter(l => l.method === 'PUT')).toEqual([expect.objectContaining({ endpoint: '/me/messages/M/attachments', ok: true })]);
    expect(JSON.parse(r.stdout)).toMatchObject({ name: 'big.bin', size: 3 * MB + 1 });
  });

  it('a 503 on a chunk is retried and the upload resumes from nextExpectedRanges', async () => {
    fs.writeFileSync('big.bin', Buffer.alloc(3 * MB + 1, 1));
    let puts = 0;
    const r = await run(['add-attachment', 'M', '--file', 'big.bin', '--content-type', 'application/octet-stream', '-o', 'json'], {
      routes: [
        ['POST', '/me/messages/M/attachments/createUploadSession', () => ({ uploadUrl, nextExpectedRanges: ['0-'] })],
        ['GET', /AttachmentSessions/, () => graphError(404)],
        ['PUT', /AttachmentSessions/, () => {
          puts++;
          if (puts === 1) return json({ nextExpectedRanges: [`${3 * MB}-`] });
          if (puts === 2) return graphError(503);
          return new Response(null, { status: 201 });
        }],
      ],
    });
    expect(r.exitCode).toBe(0);
    const ranges = r.calls.filter(c => c.method === 'PUT').map(c => c.headers['Content-Range']);
    expect(ranges).toEqual([`bytes 0-${3 * MB - 1}/${3 * MB + 1}`, `bytes ${3 * MB}-${3 * MB}/${3 * MB + 1}`, `bytes ${3 * MB}-${3 * MB}/${3 * MB + 1}`]);
  });

  it('--dry-run traces the chunk plan on stderr, never stdout', async () => {
    fs.writeFileSync('big.bin', Buffer.alloc(7 * MB, 1));
    const r = await run(['add-attachment', 'M', '--file', 'big.bin', '--content-type', 'application/octet-stream', '-o', 'json'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stdout).not.toContain('[dry-run]');
    expect(r.stderr).toContain('[dry-run] Then PUT 3 chunk(s) to the session uploadUrl');
    expect(JSON.parse(r.stdout)).toMatchObject({ dryRun: true });
  });
});

const quoted = '<html><head></head><body dir="ltr"><p>quoted thread</p></body></html>';
const draftRoutes = (create: string, id: string): Route[] => [
  ['POST', `/me/messages/${id}/send`, () => new Response(null, { status: 202 })],
  ['GET', `/me/messages/${id}`, () => ({ id, body: { contentType: 'html', content: quoted } })],
  ['PATCH', `/me/messages/${id}`, () => ({ id })],
  ['POST', create, () => ({ id })],
];

describe('RT-6: reply/forward --html go through a draft (R-2, R-3, F18)', () => {
  it('reply --html: createReply → GET → PATCH (HTML before the quoted body) → send', async () => {
    const r = await run(['reply', 'M1', '--body', '<b>Approved</b>', '--html'], { routes: draftRoutes('/me/messages/M1/createReply', 'D1') });
    expect(r.exitCode).toBe(0);
    expect(seq(r)).toEqual(['POST /me/messages/M1/createReply', 'GET /me/messages/D1', 'PATCH /me/messages/D1', 'POST /me/messages/D1/send']);
    expect(r.calls[0].body ?? {}).not.toHaveProperty('comment');
    expect(r.calls[2].body).toEqual({ body: { contentType: 'HTML', content: '<html><head></head><body dir="ltr"><b>Approved</b><p>quoted thread</p></body></html>' } });
  });

  it('reply --reply-all --html uses createReplyAll', async () => {
    const r = await run(['reply', 'M1', '--body', '<i>x</i>', '--html', '--reply-all'], { routes: draftRoutes('/me/messages/M1/createReplyAll', 'D2') });
    expect(r.exitCode).toBe(0);
    expect(seq(r)[0]).toBe('POST /me/messages/M1/createReplyAll');
    expect(seq(r).at(-1)).toBe('POST /me/messages/D2/send');
  });

  it('forward --html: createForward with recipients → PATCH → send', async () => {
    const r = await run(['forward', 'M1', '--to', 'y@contoso.com', '--comment', '<b>fyi</b>', '--html'], { routes: draftRoutes('/me/messages/M1/createForward', 'D3') });
    expect(r.exitCode).toBe(0);
    expect(seq(r)).toEqual(['POST /me/messages/M1/createForward', 'GET /me/messages/D3', 'PATCH /me/messages/D3', 'POST /me/messages/D3/send']);
    expect(r.calls[0].body).toEqual({ toRecipients: [{ emailAddress: { address: 'y@contoso.com' } }] });
    expect((r.calls[2].body as any).body.content).toContain('<body dir="ltr"><b>fyi</b><p>quoted thread</p>');
  });

  it('--html with attachments adds them to the same draft before sending', async () => {
    fs.writeFileSync('a.txt', 'x');
    const r = await run(['reply', 'M1', '--body', '<b>x</b>', '--html', '--attach', 'a.txt'], {
      routes: [['POST', '/me/messages/D1/attachments', () => ({ id: 'A1' })], ...draftRoutes('/me/messages/M1/createReply', 'D1')],
    });
    expect(r.exitCode).toBe(0);
    expect(seq(r)).toEqual(['POST /me/messages/M1/createReply', 'GET /me/messages/D1', 'PATCH /me/messages/D1', 'POST /me/messages/D1/attachments', 'POST /me/messages/D1/send']);
  });

  it('reply --html and reply with the same key conflict', async () => {
    const key = newKey();
    const first = await run(['reply', 'M1', '--body', 'b', '--idempotency-key', key]);
    expect(first.exitCode).toBe(0);
    const second = await run(['reply', 'M1', '--body', 'b', '--html', '--idempotency-key', key], { routes: draftRoutes('/me/messages/M1/createReply', 'D1') });
    expect(second.exitCode).toBe(2);
    expect(second.stderr).toMatch(/different request/);
    expect(second.calls).toHaveLength(0);
  });

  it('forward --html and forward with the same key conflict', async () => {
    const key = newKey();
    const first = await run(['forward', 'M1', '--to', 'y@contoso.com', '--comment', 'c', '--idempotency-key', key]);
    expect(first.exitCode).toBe(0);
    const second = await run(['forward', 'M1', '--to', 'y@contoso.com', '--comment', 'c', '--html', '--idempotency-key', key]);
    expect(second.exitCode).toBe(2);
    expect(second.calls).toHaveLength(0);
  });
});

describe('RT-8: a failed attachment upload deletes the unsent draft (R-8)', () => {
  it('a failing chunk upload triggers one DELETE of the draft and no send', async () => {
    fs.writeFileSync('big.bin', Buffer.alloc(3 * MB + 1, 1));
    const r = await run(['send', '--to', 'x@contoso.com', '--subject', 's', '--body', 'b', '--attach', 'big.bin'], {
      routes: [
        ['POST', '/me/messages/D1/attachments/createUploadSession', () => ({ uploadUrl })],
        ['PUT', /AttachmentSessions/, () => graphError(400)],
        ['DELETE', /AttachmentSessions/, () => new Response(null, { status: 204 })],
        ['DELETE', '/me/messages/D1', () => new Response(null, { status: 204 })],
        ['POST', '/me/messages', () => ({ id: 'D1' })],
      ],
    });
    expect(r.exitCode).toBe(1);
    expect(seq(r).filter(s => s === 'DELETE /me/messages/D1')).toHaveLength(1);
    expect(seq(r)).not.toContain('POST /me/messages/D1/send');
    expect(r.stderr).toMatch(/draft was deleted/i);
  });

  it('a failed inline attachment POST on the draft path also deletes the draft', async () => {
    fs.writeFileSync('a.bin', Buffer.alloc(2 * MB, 1));
    fs.writeFileSync('b.bin', Buffer.alloc(2 * MB, 1));
    const r = await run(['send', '--to', 'x@contoso.com', '--subject', 's', '--body', 'b', '--attach', 'a.bin', '--attach', 'b.bin'], {
      routes: [
        ['POST', '/me/messages/D1/attachments', () => graphError(400)],
        ['DELETE', '/me/messages/D1', () => new Response(null, { status: 204 })],
        ['POST', '/me/messages', () => ({ id: 'D1' })],
      ],
    });
    expect(r.exitCode).toBe(1);
    expect(seq(r)).toEqual(['POST /me/messages', 'POST /me/messages/D1/attachments', 'DELETE /me/messages/D1']);
  });

  it('a failure at /send never deletes the draft', async () => {
    fs.writeFileSync('a.bin', Buffer.alloc(2 * MB, 1));
    fs.writeFileSync('b.bin', Buffer.alloc(2 * MB, 1));
    const r = await run(['send', '--to', 'x@contoso.com', '--subject', 's', '--body', 'b', '--attach', 'a.bin', '--attach', 'b.bin'], {
      routes: [
        ['POST', '/me/messages/D1/send', () => graphError(400)],
        ['POST', '/me/messages/D1/attachments', () => ({ id: 'A' })],
        ['POST', '/me/messages', () => ({ id: 'D1' })],
      ],
    });
    expect(r.exitCode).toBe(1);
    expect(r.calls.some(c => c.method === 'DELETE')).toBe(false);
  });
});

const folders: Route = ['GET', '/me/mailFolders?', () => ({ value: [
  { id: 'INBOX_ID', displayName: 'Inbox', childFolderCount: 1 },
  { id: 'F9', displayName: 'Receipts', childFolderCount: 0 },
] })];
const inboxChildren: Route = ['GET', '/me/mailFolders/inbox/childFolders', () => ({ value: [{ id: 'P1', displayName: 'Projects', childFolderCount: 0 }] })];

describe('QP-8: every folder option resolves through resolve.ts (Q-10..Q-15)', () => {
  it('move --folder <name> resolves to the folder ID', async () => {
    const r = await run(['move', 'M1', '--folder', 'receipts'], { routes: [folders] });
    expect(r.exitCode).toBe(0);
    expect(r.calls.at(-1)).toMatchObject({ method: 'POST', path: '/me/messages/M1/move', body: { destinationId: 'F9' } });
  });

  it('move --folder "Sent Items" passes the well-known name with no lookup', async () => {
    const r = await run(['move', 'M1', '--folder', 'Sent Items']);
    expect(seq(r)).toEqual(['POST /me/messages/M1/move']);
    expect(r.calls[0].body).toEqual({ destinationId: 'sentitems' });
  });

  it('an unknown folder exits 5 and sends nothing', async () => {
    for (const argv of [['move', 'M1', '--folder', 'Nope'], ['batch-move', 'M1,M2', '--folder', 'Nope'], ['folder', 'Nope']]) {
      const r = await run(argv, { routes: [folders] });
      expect(r.exitCode, argv.join(' ')).toBe(5);
      expect(r.stderr, argv.join(' ')).toContain('Receipts');
      expect(r.calls.every(c => c.method === 'GET'), argv.join(' ')).toBe(true);
    }
  });

  it('folder <name> GETs the resolved folder', async () => {
    const r = await run(['folder', 'Receipts', '-o', 'json'], { routes: [['GET', '/me/mailFolders/F9', () => ({ id: 'F9' })], folders] });
    expect(r.exitCode).toBe(0);
    expect(seq(r).at(-1)).toBe('GET /me/mailFolders/F9');
  });

  it('list --folder takes a path to a nested folder', async () => {
    const r = await run(['list', '--folder', 'Inbox/Projects', '-o', 'json'], { routes: [inboxChildren, folders] });
    expect(r.exitCode).toBe(0);
    expect(seq(r)).toEqual(['GET /me/mailFolders/inbox/childFolders', 'GET /me/mailFolders/P1/messages']);
  });

  it('triage, create-rule --move-to, create-folder --parent and folders --parent resolve names', async () => {
    const triage = await run(['triage', '--folder', 'Receipts', '-o', 'json'], { routes: [folders] });
    expect(seq(triage)).toContain('GET /me/mailFolders/F9/messages');
    const rule = await run(['create-rule', '--name', 'r', '--from', 'a@b.com', '--move-to', 'Receipts', '-o', 'json'], { routes: [folders] });
    expect(rule.calls.at(-1)!.body).toMatchObject({ actions: { moveToFolder: 'F9' } });
    const create = await run(['create-folder', '--name', 'n', '--parent', 'Receipts', '-o', 'json'], { routes: [folders] });
    expect(seq(create).at(-1)).toBe('POST /me/mailFolders/F9/childFolders');
    const list = await run(['folders', '--parent', 'Receipts', '-o', 'json'], { routes: [folders] });
    expect(seq(list).at(-1)).toBe('GET /me/mailFolders/F9/childFolders');
  });

  it('shared-list --folder resolves against the other mailbox', async () => {
    const r = await run(['shared-list', 'u@contoso.com', '--folder', 'Receipts', '-o', 'json'], {
      routes: [['GET', '/users/u@contoso.com/mailFolders?', () => ({ value: [{ id: 'S9', displayName: 'Receipts' }] })]],
    });
    expect(r.exitCode).toBe(0);
    expect(seq(r)).toEqual(['GET /users/u@contoso.com/mailFolders', 'GET /users/u@contoso.com/mailFolders/S9/messages']);
  });

  it('batch-move resolves the folder once for every message', async () => {
    const r = await run(['batch-move', 'M1,M2', '--folder', 'Receipts'], {
      routes: [folders, ['POST', '/$batch', (c) => ({ responses: (c.body as { requests: Array<{ id: string }> }).requests.map(q => ({ id: q.id, status: 201, body: {} })) })]],
    });
    expect(r.exitCode).toBe(0);
    const batch = r.calls.find(c => c.path === '/$batch')!.body as { requests: Array<{ body: unknown }> };
    expect(batch.requests.map(q => q.body)).toEqual([{ destinationId: 'F9' }, { destinationId: 'F9' }]);
  });

  it('drafts lists the well-known Drafts folder with no lookup', async () => {
    const r = await run(['drafts', '-o', 'json'], { routes: [['GET', '/me/mailFolders/drafts/messages', () => ({ value: [] })]] });
    expect(r.exitCode).toBe(0);
    expect(seq(r)).toEqual(['GET /me/mailFolders/drafts/messages']);
    expect(r.calls[0].path).toContain('$top=10');
  });
});

describe('QP-5: rules, attachments and drafts page like every list (Q-8)', () => {
  const twoPages = (endpoint: string): Route[] => {
    const next = `https://graph.microsoft.com/v1.0${endpoint}?$skiptoken=2`;
    return [
      ['GET', `${endpoint}?$skiptoken=2`, () => ({ value: [{ id: '2' }] })],
      ['GET', endpoint, () => ({ value: [{ id: '1' }], '@odata.nextLink': next })],
    ];
  };
  for (const [argv, endpoint] of [
    [['rules'], '/me/mailFolders/inbox/messageRules'],
    [['attachments', 'M1'], '/me/messages/M1/attachments'],
    [['drafts'], '/me/mailFolders/drafts/messages'],
  ] as Array<[string[], string]>) {
    it(`mail ${argv[0]}: nextCursor without flags, both pages with --all`, async () => {
      const one = await run([...argv, '-o', 'json'], { routes: twoPages(endpoint) });
      expect(one.exitCode).toBe(0);
      const page = JSON.parse(one.stdout);
      expect(page.value.map((v: { id: string }) => v.id)).toEqual(['1']);
      expect(page.nextCursor).toBeTruthy();
      const all = await run([...argv, '--all', '-o', 'json'], { routes: twoPages(endpoint) });
      expect(JSON.parse(all.stdout).value.map((v: { id: string }) => v.id)).toEqual(['1', '2']);
    });
  }
});

describe('DR-T4/DR-T5/DR-T6: dry-run prints no invented outcomes (DR-1, DR-12, DR-13)', () => {
  const STATUS = /sent|deleted|moved|Marked|flagged|Flag |Categories set|Draft created|Draft sent|forwarded|Saved to|No messages found/;
  it('write commands print the trailer instead of a confirmation', async () => {
    const cases: string[][] = [
      ['send', '--to', 'a@b.com', '--subject', 's', '--body', 'b'],
      ['move', 'M', '--folder', 'inbox'],
      ['draft', '--to', 'a@b.com', '--subject', 's', '--body', 'b'],
      ['send-draft', 'M'],
      ['reply', 'M', '--body', 'b'],
      ['reply', 'M', '--body', '<b>b</b>', '--html'],
      ['forward', 'M', '--to', 'a@b.com'],
      ['mark-read', 'M'], ['mark-unread', 'M'], ['flag', 'M'], ['unflag', 'M'], ['complete', 'M'],
      ['categorize', 'M', '--categories', 'Red'],
      ['shared-send', 'u@b.com', '--to', 'a@b.com', '--subject', 's', '--body', 'b'],
      ['batch-read', 'M1,M2'], ['batch-move', 'M1,M2', '--folder', 'inbox'],
      ['delete', 'M'], ['delete-attachment', 'M', 'A'],
      ['triage'],
    ];
    for (const argv of cases) {
      const r = await run(argv, { dryRun: true });
      expect(r.exitCode, argv.join(' ')).toBe(0);
      const status = r.stderr.split('\n').filter(l => !l.startsWith('[dry-run]'));
      expect(status.join('\n'), argv.join(' ')).not.toMatch(STATUS);
      expect(r.stderr.trim().split('\n').at(-1), argv.join(' ')).toMatch(/^\[dry-run\] \d+ request\(s\) previewed; nothing was sent\.$/);
    }
  });

  it('triage -o json under dry-run prints {"dryRun":true}', async () => {
    const r = await run(['triage', '-o', 'json'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(JSON.parse(r.stdout)).toEqual({ dryRun: true });
    expect(r.stderr).toContain('[dry-run] GET');
  });

  it('reply --html under dry-run previews the PATCH with a <from GET> marker', async () => {
    const r = await run(['reply', 'M', '--body', '<b>b</b>', '--html'], { dryRun: true });
    expect(r.stderr).toContain('[dry-run] PATCH https://graph.microsoft.com/v1.0/me/messages/DRAFT_ID');
    expect(r.stderr).toContain('<from GET>');
  });

  it('download-attachment under dry-run exits 0 and writes no file', async () => {
    const r = await run(['download-attachment', 'M', 'A'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(fs.readdirSync(tmp)).toEqual([]);
    expect(JSON.parse(r.stdout)).toMatchObject({ dryRun: true });
  });

  it('an attachment with no bytes fails with code unsupported_attachment', async () => {
    setErrorFormat('json');
    const r = await run(['download-attachment', 'M', 'A'], { routes: [['GET', '/me/messages/M/attachments/A', () => ({ name: 'x', '@odata.type': '#microsoft.graph.itemAttachment' })]] });
    expect(r.exitCode).toBe(1);
    expect(JSON.parse(r.stderr.trim().split('\n').at(-1)!).error.code).toBe('unsupported_attachment');
  });
});

describe('VT7: enumerated options use choices (V-19)', () => {
  it('create-rule --importance rejects an unknown level (exit 2, nothing sent)', async () => {
    const r = await run(['create-rule', '--name', 'x', '--from', 'a@b.com', '--importance', 'urgent']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toContain('high');
  });

  it('update-rule --enabled accepts only true/false', async () => {
    const r = await run(['update-rule', 'R1', '--enabled', 'yes']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});

describe('RT-7: shared-send --idempotency-key (R-4)', () => {
  const argv = (key: string, subject = 's') => ['shared-send', 'u@contoso.com', '--to', 'a@contoso.com', '--subject', subject, '--body', 'b', '--idempotency-key', key];
  it('a retry with the same key and request replays without a POST', async () => {
    const key = newKey();
    const first = await run(argv(key));
    expect(first.exitCode).toBe(0);
    expect(seq(first)).toEqual(['POST /users/u@contoso.com/sendMail']);
    const second = await run(argv(key));
    expect(second.exitCode).toBe(0);
    expect(second.calls).toHaveLength(0);
    expect(JSON.parse(second.stdout)).toHaveProperty('sentAt');
  });

  it('a different request with the same key fails idempotency_conflict', async () => {
    const key = newKey();
    await run(argv(key));
    setErrorFormat('json');
    const r = await run(argv(key, 'other'));
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
    expect(JSON.parse(r.stderr.trim().split('\n').at(-1)!).error.code).toBe('idempotency_conflict');
  });
});
