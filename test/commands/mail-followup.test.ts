import { describe, it, expect } from 'vitest';
import { runCli, type FakeRoute, type RunResult } from '../e2e/run-cli.js';

const ENV = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_TOKEN: 'fake' };
const GRAPH = 'https://graph.microsoft.com/v1.0';
const FOLDER_ID = 'AAMkADfolder' + 'x'.repeat(40);

function mail(args: string[], routes: FakeRoute[], env: Record<string, string> = {}): Promise<RunResult> {
  return runCli(['mail', ...args], { env: { ...ENV, ...env }, routes });
}

function bodyOf(r: RunResult, method: string): Record<string, any> {
  const req = r.requests.find((q) => q.method === method);
  expect(req, `${method} request`).toBeDefined();
  return JSON.parse(req!.body ?? 'null');
}

function param(url: string, name: string): string | null {
  return new URL(url).searchParams.get(name);
}

describe('DRY-1 mail success lines under --dry-run', () => {
  it.each([
    [['mail', 'send', '--to', 'a@b.cd', '--subject', 's', '--body', 'b'], 'POST', '/me/sendMail', 'Email sent'],
    [['mail', 'reply', 'msg-1', '--body', 'hi'], 'POST', '/me/messages/msg-1/reply', 'Reply sent'],
    [['mail', 'flag', 'msg-1', '--due', '2099-01-01'], 'PATCH', '/me/messages/msg-1', 'Message flagged'],
    [['mail', 'batch-move', 'm1,m2', '--folder', 'archive'], 'POST', '/me/messages/m2/move', 'Moved'],
    [['mail', 'batch-read', 'm1,m2'], 'PATCH', '/me/messages/m2', 'Marked'],
  ])('%j previews and prints no success line', async (args, method, path, success) => {
    const r = await runCli(['--dry-run', ...args], { env: ENV, routes: [] });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toContain(`[dry-run] ${method} ${GRAPH}${path}`);
    expect(r.stderr).not.toContain(success);
  });

  it('real runs still confirm on stderr', async () => {
    const r = await mail(['batch-move', 'm1', '--folder', 'archive'], [{ match: '/move', method: 'POST', status: 201, body: {} }]);
    expect(r.code).toBe(0);
    expect(r.stderr).toContain('Moved 1 message(s).');
    const s = await mail(['send', '--to', 'a@b.cd', '--subject', 's', '--body', 'b'], [{ match: '/sendMail', method: 'POST', status: 202, body: '' }]);
    expect(s.stderr).toContain('Email sent successfully.');
  });
});

describe('MAIL-9 attachments $select', () => {
  it('selects metadata only (no contentBytes)', async () => {
    const r = await mail(['attachments', 'msg-1', '-o', 'json'], [{ match: '/attachments', body: { value: [] } }]);
    expect(r.code).toBe(0);
    expect(param(r.requests[0].url, '$select')).toBe('id,name,contentType,size,isInline,lastModifiedDateTime');
  });
});

describe('MAIL-10 folder options resolve names', () => {
  const existing: FakeRoute = { match: '/messageRules', method: 'GET', body: { value: [] } };
  const created: FakeRoute = { match: '/messageRules', method: 'POST', status: 201, body: { id: 'r1' } };

  it('create-rule --move-to Archive sends the real folder ID', async () => {
    const r = await mail(['create-rule', '--name', 'A', '--from', 'a@b.cd', '--move-to', 'Archive', '--sequence', '1', '-o', 'json'], [
      { match: '/me/mailFolders/archive', method: 'GET', body: { id: FOLDER_ID } },
      created,
    ]);
    expect(r.code).toBe(0);
    const get = r.requests.find((q) => q.url.includes('/mailFolders/archive'))!;
    expect(param(get.url, '$select')).toBe('id');
    expect(bodyOf(r, 'POST').actions).toEqual({ moveToFolder: FOLDER_ID });
  });

  it('create-rule --move-to a display name resolves it; ambiguous names fail', async () => {
    const r = await mail(['create-rule', '--name', 'A', '--from', 'a@b.cd', '--move-to', 'Receipts', '-o', 'json'], [
      { match: '/me/mailFolders\\?', body: { value: [{ id: 'rcpt-1', displayName: 'Receipts' }] } },
      existing, created,
    ]);
    expect(r.code).toBe(0);
    expect(bodyOf(r, 'POST').actions).toEqual({ moveToFolder: 'rcpt-1' });

    const amb = await mail(['create-rule', '--name', 'A', '--from', 'a@b.cd', '--move-to', 'Dup', '-o', 'json'], [
      { match: '/me/mailFolders/p1/childFolders', body: { value: [{ id: 'd1', displayName: 'Dup', childFolderCount: 0 }] } },
      { match: '/me/mailFolders/p2/childFolders', body: { value: [{ id: 'd2', displayName: 'dup', childFolderCount: 0 }] } },
      { match: '/me/mailFolders\\?.*(%24|\\$)filter', body: { value: [] } },
      { match: '/me/mailFolders\\?', body: { value: [
        { id: 'p1', displayName: 'P1', childFolderCount: 1 },
        { id: 'p2', displayName: 'P2', childFolderCount: 1 },
      ] } },
      existing, created,
    ]);
    expect(amb.code).toBe(1);
    expect(amb.stderr).toMatch(/ambiguous/i);
    expect(amb.requests.filter((q) => q.method === 'POST')).toHaveLength(0);
  });

  it('create-rule --move-to keeps the name under --dry-run', async () => {
    const r = await runCli(['--dry-run', 'mail', 'create-rule', '--name', 'A', '--from', 'a@b.cd', '--move-to', 'Archive', '--sequence', '1'], { env: ENV, routes: [] });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toContain('"moveToFolder":"archive"');
  });

  it('folders --parent Inbox lists inbox child folders', async () => {
    const r = await mail(['folders', '--parent', 'Inbox', '-o', 'json'], [{ match: '/childFolders', body: { value: [] } }]);
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(1);
    expect(r.requests[0].url).toContain('/me/mailFolders/inbox/childFolders');
  });

  it('create-folder --parent resolves a display name', async () => {
    const r = await mail(['create-folder', '--name', 'Sub', '--parent', 'Projects', '-o', 'json'], [
      { match: '/me/mailFolders/proj-1/childFolders', method: 'POST', status: 201, body: { id: 'sub-1' } },
      { match: '/me/mailFolders\\?', body: { value: [{ id: 'proj-1', displayName: 'Projects' }] } },
    ]);
    expect(r.code).toBe(0);
    expect(r.requests.at(-1)!.url).toContain('/me/mailFolders/proj-1/childFolders');
    expect(bodyOf(r, 'POST')).toEqual({ displayName: 'Sub' });
  });
});

describe('TZ-1 mail flag dates', () => {
  const patch: FakeRoute = { match: '/me/messages/msg-1', method: 'PATCH', body: {} };

  it('uses the local zone when Graph documents it', async () => {
    const r = await mail(['flag', 'msg-1', '--due', '2026-10-07', '--start', '2026-10-01'], [patch], { TZ: 'Asia/Tokyo' });
    expect(r.code).toBe(0);
    expect(bodyOf(r, 'PATCH').flag).toEqual({
      flagStatus: 'flagged',
      startDateTime: { dateTime: '2026-10-01T00:00:00', timeZone: 'Asia/Tokyo' },
      dueDateTime: { dateTime: '2026-10-07T00:00:00', timeZone: 'Asia/Tokyo' },
    });
  });

  it('falls back to UTC midnight when the zone is not documented', async () => {
    const r = await mail(['flag', 'msg-1', '--due', '2026-10-07'], [patch], { TZ: 'Europe/Amsterdam' });
    expect(r.code).toBe(0);
    expect(bodyOf(r, 'PATCH').flag.dueDateTime).toEqual({ dateTime: '2026-10-07T00:00:00', timeZone: 'UTC' });
  });

  it('defaults the start to the local today, not the UTC date', async () => {
    // One of these zones always has a different calendar date than UTC right now.
    const utcHour = new Date().getUTCHours();
    const zone = utcHour >= 10 ? 'Pacific/Kiritimati' : 'Pacific/Pago_Pago';
    const today = new Intl.DateTimeFormat('en-CA', { timeZone: zone }).format(new Date());
    expect(today).not.toBe(new Date().toISOString().slice(0, 10));
    const r = await mail(['flag', 'msg-1', '--due', '2099-01-01'], [patch], { TZ: zone });
    expect(r.code).toBe(0);
    expect(bodyOf(r, 'PATCH').flag.startDateTime.dateTime).toBe(`${today}T00:00:00`);
  });
});

describe('CHOICE-1 create-rule --importance', () => {
  it('rejects a value outside low/normal/high before any request', async () => {
    const r = await mail(['create-rule', '--name', 'A', '--from', 'a@b.cd', '--importance', 'urgent'], []);
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('Allowed choices');
    expect(r.requests).toHaveLength(0);
  });

  it('accepts high', async () => {
    const r = await mail(['create-rule', '--name', 'A', '--from', 'a@b.cd', '--importance', 'high', '--sequence', '1', '-o', 'json'], [
      { match: '/messageRules', method: 'POST', status: 201, body: { id: 'r1' } },
    ]);
    expect(r.code).toBe(0);
    expect(bodyOf(r, 'POST').actions).toEqual({ markImportance: 'high' });
  });
});
