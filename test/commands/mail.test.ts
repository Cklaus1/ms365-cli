import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { runCommand } from '../helpers/cli.js';
import { registerMailCommands } from '../../src/commands/mail.js';

const run = (argv: string[], opts: Parameters<typeof runCommand>[2] = {}) =>
  runCommand(registerMailCommands, ['mail', ...argv], opts);

// cwd is nested two levels inside `root` so a traversal regression still lands inside the cleanup root.
let root: string;
let tmp: string;
let origCwd: string;
beforeEach(() => {
  origCwd = process.cwd();
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'mail-test-'));
  tmp = path.join(root, 'a', 'b');
  fs.mkdirSync(tmp, { recursive: true });
  process.chdir(tmp);
});
afterEach(() => {
  process.chdir(origCwd);
  fs.rmSync(root, { recursive: true, force: true });
});

describe('mail download-attachment (M-1, D12)', () => {
  const attachment = (name: string) => ['GET', '/me/messages/M/attachments/A', () => ({
    name, contentBytes: Buffer.from('hello').toString('base64'),
  })] as [string, string, () => unknown];

  it('strips path components from a sender-controlled name', async () => {
    const r = await run(['download-attachment', 'M', 'A'], { routes: [attachment('../../evil.sh')] });
    expect(r.exitCode).toBe(0);
    expect(fs.readFileSync(path.join(tmp, 'evil.sh'), 'utf8')).toBe('hello');
    expect(fs.existsSync(path.join(root, 'evil.sh'))).toBe(false);
  });

  it('refuses to overwrite an existing file without --force', async () => {
    fs.writeFileSync(path.join(tmp, 'report.pdf'), 'original');
    const r = await run(['download-attachment', 'M', 'A'], { routes: [attachment('report.pdf')] });
    expect(r.exitCode).toBe(1);
    expect(r.stderr).toMatch(/overwrite/i);
    expect(fs.readFileSync(path.join(tmp, 'report.pdf'), 'utf8')).toBe('original');
  });

  it('overwrites with --force', async () => {
    fs.writeFileSync(path.join(tmp, 'report.pdf'), 'original');
    const r = await run(['download-attachment', 'M', 'A', '--force'], { routes: [attachment('report.pdf')] });
    expect(r.exitCode).toBe(0);
    expect(fs.readFileSync(path.join(tmp, 'report.pdf'), 'utf8')).toBe('hello');
    expect(r.stdout).toBe('');
  });
});

describe('mail triage (M-3)', () => {
  it('default filter leads with receivedDateTime so $orderby is allowed', async () => {
    const r = await run(['triage', '-o', 'json'], { routes: [['GET', '/me/messages', () => ({ value: [] })]] });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].path).toContain("$filter=receivedDateTime ge 1900-01-01T00:00:00Z and (isRead eq false)");
    expect(r.calls[0].path).toContain('$orderby=receivedDateTime desc');
    expect(JSON.parse(r.stdout)).toEqual({ value: [] });
  });

  it('--include-read sends no $filter', async () => {
    const r = await run(['triage', '--include-read', '-o', 'json'], { routes: [['GET', '/me/messages', () => ({ value: [] })]] });
    expect(r.calls[0].path).not.toContain('$filter');
  });
});

describe('mail create-rule (M-11, D11)', () => {
  it('refuses a rule with no conditions (exit 2, nothing sent)', async () => {
    const r = await run(['create-rule', '--name', 'x', '--delete']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toContain('--all-messages');
  });

  it('--all-messages allows an unconditioned rule', async () => {
    const r = await run(['create-rule', '--name', 'x', '--mark-read', '--all-messages', '-o', 'json']);
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toMatchObject({ conditions: {}, actions: { markAsRead: true }, isEnabled: true });
  });

  it('--disabled creates the rule disabled; --enabled no longer exists (D29)', async () => {
    const r1 = await run(['create-rule', '--name', 'x', '--from', 'a@b.com', '--mark-read', '--disabled', '-o', 'json']);
    expect(r1.calls[0].body).toMatchObject({ isEnabled: false, conditions: { senderContains: ['a@b.com'] } });
    const r2 = await run(['create-rule', '--name', 'x', '--from', 'a@b.com', '--mark-read', '--enabled', '-o', 'json']);
    expect(r2.exitCode).toBe(2);
    expect(r2.calls).toHaveLength(0);
    const r3 = await run(['create-rule', '--name', 'x', '--from', 'a@b.com', '--mark-read', '-o', 'json']);
    expect(r3.calls[0].body).toMatchObject({ isEnabled: true });
  });
});

describe('mail add-attachment (M-17)', () => {
  it('--file under 3MB posts inline base64 with the file name', async () => {
    fs.writeFileSync(path.join(tmp, 'small.txt'), 'abc');
    const r = await run(['add-attachment', 'M', '--file', 'small.txt', '--content-type', 'text/plain', '-o', 'json']);
    expect(r.exitCode).toBe(0);
    expect(r.calls[0]).toMatchObject({ method: 'POST', path: '/me/messages/M/attachments' });
    expect(r.calls[0].body).toMatchObject({ name: 'small.txt', contentBytes: Buffer.from('abc').toString('base64') });
  });

  it('--file over 3MB uses an upload session with chunked PUTs (no bearer token)', async () => {
    const size = 7 * 1024 * 1024;
    fs.writeFileSync(path.join(tmp, 'big.bin'), Buffer.alloc(size, 1));
    const uploadUrl = 'https://outlook.office.com/api/v2.0/Users(\'x\')/Messages(\'M\')/AttachmentSessions(\'s\')?authtoken=t';
    let puts = 0;
    const r = await run(['add-attachment', 'M', '--file', 'big.bin', '--content-type', 'application/octet-stream', '-o', 'json'], {
      routes: [
        ['POST', '/me/messages/M/attachments/createUploadSession', () => ({ uploadUrl, nextExpectedRanges: ['0-'] })],
        ['PUT', /AttachmentSessions/, () => {
          puts++;
          return puts < 3 ? new Response(JSON.stringify({ nextExpectedRanges: [`${puts * 3 * 1024 * 1024}-`] }), { status: 200 })
            : new Response(null, { status: 201 });
        }],
      ],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toEqual({ AttachmentItem: { attachmentType: 'file', name: 'big.bin', size, contentType: 'application/octet-stream' } });
    const putCalls = r.calls.filter(c => c.method === 'PUT');
    expect(putCalls).toHaveLength(3);
    expect(putCalls[0].headers['Content-Range']).toBe(`bytes 0-${3 * 1024 * 1024 - 1}/${size}`);
    expect(putCalls[2].headers['Content-Range']).toBe(`bytes ${6 * 1024 * 1024}-${size - 1}/${size}`);
    expect(putCalls.every(c => !c.headers.Authorization)).toBe(true);
  });

  it('requires exactly one of --content or --file', async () => {
    const r = await run(['add-attachment', 'M', '--name', 'a', '--content-type', 'text/plain']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('still accepts small base64 --content', async () => {
    const r = await run(['add-attachment', 'M', '--name', 'a.txt', '--content-type', 'text/plain', '--content', 'YWJj', '-o', 'json']);
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toMatchObject({ name: 'a.txt', contentBytes: 'YWJj' });
  });
});

describe('mail flag --due (M-19)', () => {
  it('sends startDateTime and dueDateTime in the resolved time zone', async () => {
    const r = await run(['flag', 'M', '--due', '2099-03-01', '--timezone', 'America/New_York']);
    expect(r.exitCode).toBe(0);
    const flag = (r.calls[0].body as { flag: Record<string, { dateTime: string; timeZone: string }> }).flag;
    expect(flag.dueDateTime).toEqual({ dateTime: '2099-03-01T00:00:00', timeZone: 'America/New_York' });
    expect(flag.startDateTime.timeZone).toBe('America/New_York');
    expect(flag.startDateTime.dateTime < flag.dueDateTime.dateTime).toBe(true);
  });
});

describe('mail list (M-24, M-30, M-31)', () => {
  it('escapes quotes in --search with KQL quoting', async () => {
    const r = await run(['list', '--search', 'say "hi"', '-o', 'json']);
    expect(r.calls[0].path).toContain('$search="say \\"hi\\""');
  });

  it('rejects --search with --filter (exit 2, nothing sent)', async () => {
    const r = await run(['list', '--search', 'foo', '--filter', 'isRead eq false']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('warns about the 250-result cap with --search --all', async () => {
    const r = await run(['list', '--search', 'foo', '--all', '-o', 'json']);
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain('250');
  });

  it('looks up non-ASCII folder names intact', async () => {
    const r = await run(['list', '--folder', "L'Été", '-o', 'json'], {
      routes: [['GET', '/me/mailFolders?', () => ({ value: [{ id: 'F1', displayName: "L'Été" }] })]],
    });
    expect(r.calls[0].path).toContain("$filter=displayName eq 'L''Été'");
    expect(r.calls[1].path.startsWith('/me/mailFolders/F1/messages')).toBe(true);
  });
});

describe('status lines go to stderr (C-19)', () => {
  it('draft / categorize / batch-read / batch-move / shared-send keep stdout clean', async () => {
    const cases: string[][] = [
      ['draft', '--to', 'a@b.com', '--subject', 's', '--body', 'b', '-o', 'table'],
      ['categorize', 'M', '--categories', 'Red'],
      ['batch-read', 'M1,M2'],
      ['batch-move', 'M1,M2', '--folder', 'inbox'],
      ['shared-send', 'u@b.com', '--to', 'a@b.com', '--subject', 's', '--body', 'b'],
    ];
    for (const argv of cases) {
      const r = await run(argv, { routes: [
        ['POST', '/$batch', (c) => ({ responses: (c.body as { requests: Array<{ id: string }> }).requests.map(q => ({ id: q.id, status: 200, body: {} })) })],
        ['POST', '/me/messages', () => ({ id: 'D1' })],
      ] });
      expect(r.exitCode, argv.join(' ')).toBe(0);
      expect(r.stdout, argv.join(' ')).toBe('');
      expect(r.stderr.length, argv.join(' ')).toBeGreaterThan(0);
    }
  });
});

describe('mail update-draft (G-1)', () => {
  it('refuses to PATCH a non-draft message', async () => {
    const r = await run(['update-draft', 'M', '--subject', 'x'], {
      routes: [['GET', '/me/messages/M', () => ({ id: 'M', isDraft: false })]],
    });
    expect(r.exitCode).toBe(1);
    expect(r.stderr).toMatch(/not a draft/i);
    expect(r.calls.some(c => c.method === 'PATCH')).toBe(false);
  });

  it('PATCHes a draft', async () => {
    const r = await run(['update-draft', 'M', '--subject', 'x', '-o', 'json'], {
      routes: [['GET', '/me/messages/M', () => ({ id: 'M', isDraft: true })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].path).toContain('$select=isDraft');
    expect(r.calls[1]).toMatchObject({ method: 'PATCH', path: '/me/messages/M', body: { subject: 'x' } });
  });
});
