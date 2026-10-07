import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { runCli, tempDir, type FakeRoute, type RunResult } from '../e2e/run-cli.js';

const ENV = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_TOKEN: 'fake' };

function mail(args: string[], routes: FakeRoute[], extra: { cwd?: string; env?: Record<string, string> } = {}): Promise<RunResult> {
  return runCli(['mail', ...args], { ...extra, env: { ...ENV, ...extra.env }, routes });
}

function param(url: string, name: string): string | null {
  return new URL(url).searchParams.get(name);
}

function bodyOf(r: RunResult, method: string): Record<string, any> {
  const req = r.requests.find((q) => q.method === method);
  expect(req, `${method} request`).toBeDefined();
  return JSON.parse(req!.body ?? 'null');
}

const BYTES = Buffer.from([0x25, 0x50, 0x44, 0x46, 0x00, 0xff, 0x0d, 0x0a]);

function attachmentRoutes(name: string): FakeRoute[] {
  return [
    { match: '/attachments/att-1/(\\$|%24)value', method: 'GET', headers: { 'content-type': 'application/pdf' }, bodyBase64: BYTES.toString('base64') },
    { match: '/attachments/att-1', method: 'GET', body: { '@odata.type': '#microsoft.graph.fileAttachment', id: 'att-1', name, size: BYTES.length } },
  ];
}

describe('MAIL-1 download-attachment', () => {
  it('keeps a ../../x name inside the out dir, fetches $value, and refuses to overwrite', async () => {
    const root = tempDir();
    const out = path.join(root, 'a', 'b');
    fs.mkdirSync(out, { recursive: true });
    const r = await mail(['download-attachment', 'msg-1', 'att-1', '--out', out], attachmentRoutes('../../x'));
    expect(r.code).toBe(0);
    expect(fs.readFileSync(path.join(out, 'x')).equals(BYTES)).toBe(true);
    expect(fs.existsSync(path.join(root, 'x'))).toBe(false);
    expect(r.requests.some((q) => /\/attachments\/att-1\/(\$|%24)value$/.test(q.url))).toBe(true);
    expect(r.stdout).toBe('');

    fs.writeFileSync(path.join(out, 'x'), 'keep me');
    const again = await mail(['download-attachment', 'msg-1', 'att-1', '--out', out], attachmentRoutes('../../x'));
    expect(again.code).toBe(1);
    expect(again.stderr).toContain('File exists');
    expect(fs.readFileSync(path.join(out, 'x'), 'utf8')).toBe('keep me');

    const forced = await mail(['download-attachment', 'msg-1', 'att-1', '--out', out, '--force'], attachmentRoutes('../../x'));
    expect(forced.code).toBe(0);
    expect(fs.readFileSync(path.join(out, 'x')).equals(BYTES)).toBe(true);
  });

  it('defaults to the basename in the current directory', async () => {
    const root = tempDir();
    const cwd = path.join(root, 'work');
    fs.mkdirSync(cwd);
    const r = await mail(['download-attachment', 'msg-1', 'att-1'], attachmentRoutes('..\\..\\evil.pdf'), { cwd });
    expect(r.code).toBe(0);
    expect(fs.readdirSync(cwd)).toEqual(['evil.pdf']);
    expect(fs.readdirSync(root)).toEqual(['work']);
  });

  it('rejects an attachment name that is only dots', async () => {
    const cwd = tempDir();
    const r = await mail(['download-attachment', 'msg-1', 'att-1'], attachmentRoutes('..'), { cwd });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('--out');
    expect(fs.readdirSync(cwd)).toEqual([]);
  });

  it('explains that reference attachments have no content', async () => {
    const r = await mail(['download-attachment', 'msg-1', 'att-1'], [
      { match: '/attachments/att-1', body: { '@odata.type': '#microsoft.graph.referenceAttachment', id: 'att-1', name: 'link' } },
    ]);
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('reference attachment');
    expect(r.requests).toHaveLength(1);
  });
});

describe('MAIL-2 create-rule / update-rule', () => {
  const existing: FakeRoute = { match: '/messageRules', method: 'GET', body: { value: [{ id: 'r1', sequence: 2 }, { id: 'r2', sequence: 5 }] } };
  const created: FakeRoute = { match: '/messageRules', method: 'POST', status: 201, body: { id: 'r3' } };

  it('sends sequence = max existing + 1 and exact sender addresses', async () => {
    const folderId = 'AAMkADf1' + 'x'.repeat(40);
    const r = await mail(['create-rule', '--name', 'Bob', '--from', 'bob@x.com', '--move-to', folderId, '-o', 'json'], [existing, created]);
    expect(r.code).toBe(0);
    const body = bodyOf(r, 'POST');
    expect(body.sequence).toBe(6);
    expect(body.conditions).toEqual({ fromAddresses: [{ emailAddress: { address: 'bob@x.com' } }] });
    expect(body.actions).toEqual({ moveToFolder: folderId });
    expect(body.isEnabled).toBe(true);
    expect(JSON.parse(r.stdout).id).toBe('r3');
  });

  it('honours --sequence and --no-enabled', async () => {
    const r = await mail(['create-rule', '--name', 'S', '--subject-contains', 'a,b', '--mark-read', '--sequence', '3', '--no-enabled', '-o', 'json'], [existing, created]);
    expect(r.code).toBe(0);
    const body = bodyOf(r, 'POST');
    expect(body.sequence).toBe(3);
    expect(body.isEnabled).toBe(false);
    expect(body.conditions).toEqual({ subjectContains: ['a', 'b'] });
  });

  it('requires a confirmation for a rule without conditions (refused when non-interactive)', async () => {
    const r = await mail(['create-rule', '--name', 'All', '--delete'], [existing, created]);
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('Nothing changed');
    expect(r.requests.filter((q) => q.method === 'POST')).toHaveLength(0);
  });

  it('requires at least one action', async () => {
    const r = await mail(['create-rule', '--name', 'None', '--from', 'bob@x.com'], [existing, created]);
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('action');
    expect(r.requests.filter((q) => q.method === 'POST')).toHaveLength(0);
  });

  it('update-rule --enabled yes enables, garbage is rejected', async () => {
    const patch: FakeRoute = { match: '/messageRules/r1', method: 'PATCH', body: { id: 'r1' } };
    const yes = await mail(['update-rule', 'r1', '--enabled', 'yes', '-o', 'json'], [patch]);
    expect(yes.code).toBe(0);
    expect(bodyOf(yes, 'PATCH')).toEqual({ isEnabled: true });
    const off = await mail(['update-rule', 'r1', '--enabled', 'false', '-o', 'json'], [patch]);
    expect(bodyOf(off, 'PATCH')).toEqual({ isEnabled: false });
    const bad = await mail(['update-rule', 'r1', '--enabled', 'maybe'], [patch]);
    expect(bad.code).toBe(1);
    expect(bad.requests).toHaveLength(0);
  });
});

describe('MAIL-3 folder resolution', () => {
  const tree: FakeRoute[] = [
    // top-level exact filter finds nothing
    { match: '/me/mailFolders\\?.*%24filter|/me/mailFolders\\?.*\\$filter', body: { value: [] } },
    { match: '/me/mailFolders/top-2/childFolders', body: { value: [{ id: 'deep-1', displayName: 'Q&A', childFolderCount: 0 }] } },
    { match: '/me/mailFolders\\?', body: { value: [
      { id: 'top-1', displayName: 'QA', childFolderCount: 0 },
      { id: 'top-2', displayName: 'Projects', childFolderCount: 1 },
    ] } },
    { match: '/move', method: 'POST', status: 201, body: { id: 'moved' } },
  ];

  it('finds a nested folder by exact name (Q&A is not QA) and moves into it', async () => {
    const r = await mail(['move', 'msg-1', '--folder', 'Q&A'], tree);
    expect(r.code).toBe(0);
    const filterReq = r.requests.find((q) => param(q.url, '$filter'));
    expect(param(filterReq!.url, '$filter')).toBe("displayName eq 'Q&A'");
    expect(bodyOf(r, 'POST')).toEqual({ destinationId: 'deep-1' });
    expect(r.stdout).toBe('');
  });

  it('escapes quotes with odataQuote', async () => {
    const r = await mail(['folder', "O'Brien", '-o', 'json'], [
      { match: '/me/mailFolders\\?', body: { value: [{ id: 'ob', displayName: "O'Brien" }] } },
      { match: '/me/mailFolders/ob', body: { id: 'ob', displayName: "O'Brien" } },
    ]);
    expect(r.code).toBe(0);
    expect(param(r.requests[0].url, '$filter')).toBe("displayName eq 'O''Brien'");
    expect(JSON.parse(r.stdout).id).toBe('ob');
  });

  it('fails with "not found" instead of using the name as an ID', async () => {
    const r = await mail(['batch-move', 'm1', '--folder', 'Nope'], tree);
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('Folder "Nope" not found');
    expect(r.requests.filter((q) => q.method === 'POST')).toHaveLength(0);
  });

  it('triage --folder resolves names; well-known names pass through', async () => {
    const r = await mail(['triage', '--folder', 'Q&A', '-o', 'json'], [...tree.slice(0, 2), { match: '/deep-1/messages', body: { value: [] } }, ...tree.slice(2)]);
    expect(r.code).toBe(0);
    expect(r.requests.at(-1)!.url).toContain('/me/mailFolders/deep-1/messages');
    const wk = await mail(['list', '--folder', 'Sent Items', '-o', 'json'], [{ match: '/sentitems/messages', body: { value: [] } }]);
    expect(wk.code).toBe(0);
    expect(wk.requests).toHaveLength(1);
  });
});

describe('MAIL-4 flag --due', () => {
  it('sends startDateTime together with dueDateTime (UTC zone)', async () => {
    const r = await mail(['flag', 'msg-1', '--due', '2026-12-01', '--start', '2026-11-20'], [{ match: '/me/messages/msg-1', method: 'PATCH', body: {} }], { env: { TZ: 'UTC' } });
    expect(r.code).toBe(0);
    expect(bodyOf(r, 'PATCH').flag).toEqual({
      flagStatus: 'flagged',
      startDateTime: { dateTime: '2026-11-20T00:00:00', timeZone: 'UTC' },
      dueDateTime: { dateTime: '2026-12-01T00:00:00', timeZone: 'UTC' },
    });
  });

  it('defaults the start to today, or to the due date when that is earlier', async () => {
    const r = await mail(['flag', 'msg-1', '--due', '2000-01-02'], [{ match: '/me/messages/msg-1', method: 'PATCH', body: {} }], { env: { TZ: 'UTC' } });
    expect(r.code).toBe(0);
    expect(bodyOf(r, 'PATCH').flag.startDateTime).toEqual({ dateTime: '2000-01-02T00:00:00', timeZone: 'UTC' });
  });

  it.each([
    ['--due', '2026-02-30'],
    ['--due', '2026-12-01', '--start', '2026-11'],
  ])('rejects a date that is not a real YYYY-MM-DD day, before any request: %s %s', async (...args) => {
    const r = await mail(['flag', 'msg-1', ...args], []);
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toContain('Use YYYY-MM-DD format');
  });
});

describe('MAIL-5/6 triage', () => {
  const msgs = [{ id: 'A'.repeat(40), subject: 's', from: { emailAddress: { name: 'n', address: 'n@x.io' } }, receivedDateTime: '2026-10-01T10:00:00Z' }];

  it('reads the inbox with receivedDateTime first in $filter', async () => {
    const r = await mail(['triage', '-o', 'json'], [{ match: '/me/mailFolders/inbox/messages', body: { value: msgs } }]);
    expect(r.code).toBe(0);
    const url = r.requests[0].url;
    expect(param(url, '$filter')).toBe('receivedDateTime ge 1900-01-01T00:00:00Z and (isRead eq false)');
    expect(param(url, '$orderby')).toBe('receivedDateTime desc');
    expect(param(url, '$top')).toBe('20');
    expect(JSON.parse(r.stdout).value).toHaveLength(1);
  });

  it('--include-read keeps a user --filter', async () => {
    const r = await mail(['triage', '--include-read', '--filter', "importance eq 'high'", '-o', 'json'], [{ match: '/messages', body: { value: [] } }]);
    expect(param(r.requests[0].url, '$filter')).toBe("receivedDateTime ge 1900-01-01T00:00:00Z and (importance eq 'high')");
    const all = await mail(['triage', '--include-read', '-o', 'json'], [{ match: '/messages', body: { value: [] } }]);
    expect(param(all.requests[0].url, '$filter')).toBe('receivedDateTime ge 1900-01-01T00:00:00Z');
  });

  it('empty result still parses as JSON; csv carries full IDs', async () => {
    const empty = await mail(['triage', '-o', 'json'], [{ match: '/messages', body: { value: [] } }]);
    expect(empty.code).toBe(0);
    expect(JSON.parse(empty.stdout)).toEqual({ value: [] });
    const csv = await mail(['triage', '-o', 'csv'], [{ match: '/messages', body: { value: msgs } }]);
    expect(csv.stdout).toContain('A'.repeat(40));
    const yaml = await mail(['triage', '-o', 'yaml'], [{ match: '/messages', body: { value: msgs } }]);
    expect(yaml.stdout).toContain('A'.repeat(40));
    expect(yaml.stdout).not.toContain('[A]=attachment');
  });
});

describe('MAIL-7', () => {
  it('reply --html sends an HTML message body instead of a comment', async () => {
    const r = await mail(['reply', 'msg-1', '--body', '<b>hi</b>', '--html'], [{ match: '/reply', method: 'POST', status: 202, body: '' }]);
    expect(r.code).toBe(0);
    expect(bodyOf(r, 'POST')).toEqual({ message: { body: { contentType: 'HTML', content: '<b>hi</b>' } } });
    const plain = await mail(['reply', 'msg-1', '--body', 'hi'], [{ match: '/reply', method: 'POST', status: 202, body: '' }]);
    expect(bodyOf(plain, 'POST')).toEqual({ comment: 'hi' });
  });

  it('delete-attachment needs confirmation', async () => {
    const r = await mail(['delete-attachment', 'msg-1', 'att-1'], [{ match: '/attachments/', method: 'DELETE', status: 204 }]);
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });

  it('read -o json includes the attachment list', async () => {
    const r = await mail(['read', 'msg-1', '-o', 'json'], [
      { match: '/msg-1/attachments', body: { value: [{ id: 'att-1', name: 'a.pdf', contentType: 'application/pdf', size: 3 }] } },
      { match: '/msg-1', body: { '@odata.context': 'x', id: 'msg-1', subject: 's', hasAttachments: true } },
    ]);
    expect(r.code).toBe(0);
    const out = JSON.parse(r.stdout);
    expect(out.attachments).toEqual([{ id: 'att-1', name: 'a.pdf', contentType: 'application/pdf', size: 3 }]);
    expect(out['@odata.context']).toBeUndefined();
  });
});

describe('MAIL-8 $search quoting', () => {
  it('escapes double quotes and backslashes in the search phrase', async () => {
    const r = await mail(['list', '--search', 'say "hi" \\o/', '-o', 'json'], [{ match: '/me/messages', body: { value: [] } }]);
    expect(r.code).toBe(0);
    expect(param(r.requests[0].url, '$search')).toBe('"say \\"hi\\" \\\\o/"');
  });
});

describe('CORE-5 batch commands', () => {
  const longId = 'AAMkAD' + 'x'.repeat(60);

  it('batch-read exits 1 when one item fails, naming the full ID and the error', async () => {
    const r = await mail(['batch-read', `ok-1,${longId}`], [
      { match: `/me/messages/${longId}`, method: 'PATCH', status: 404, body: { error: { code: 'ErrorItemNotFound', message: 'The specified object was not found in the store.' } } },
      { match: '/me/messages/ok-1', method: 'PATCH', body: {} },
    ]);
    expect(r.code).toBe(1);
    expect(r.stderr).toContain(longId);
    expect(r.stderr).toContain('not found in the store');
    expect(r.stderr).toContain('Marked 1 message(s) as read');
    expect(r.stdout).toBe('');
  });

  it('batch-move exits 1 on a failure and 0 when all succeed', async () => {
    const routes: FakeRoute[] = [
      { match: `/${longId}/move`, method: 'POST', status: 400, body: { error: { code: 'ErrorInvalidIdMalformed', message: 'Id is malformed.' } } },
      { match: '/move', method: 'POST', status: 201, body: {} },
    ];
    const r = await mail(['batch-move', `ok-1,${longId}`, '--folder', 'archive'], routes);
    expect(r.code).toBe(1);
    expect(r.stderr).toContain(`Failed: ${longId}: Graph API error (400): Id is malformed.`);
    const ok = await mail(['batch-move', 'ok-1,ok-2', '--folder', 'archive'], routes);
    expect(ok.code).toBe(0);
    expect(JSON.parse(ok.requests[0].body!)).toEqual({ destinationId: 'archive' });
  });
});

describe('CORE-4 / PAGE-1', () => {
  it('categorize and shared-send keep stdout empty', async () => {
    const c = await mail(['categorize', 'msg-1', '--categories', 'Red'], [{ match: '/msg-1', method: 'PATCH', body: {} }]);
    expect(c.code).toBe(0);
    expect(c.stdout).toBe('');
    expect(c.stderr).toContain('Categories set: Red');
    const s = await mail(['shared-send', 'box@x.io', '--to', 'a@b.cd', '--subject', 's', '--body', 'b'], [{ match: '/sendMail', method: 'POST', status: 202, body: '' }]);
    expect(s.code).toBe(0);
    expect(s.stdout).toBe('');
  });

  it('draft -o json prints only JSON', async () => {
    const r = await mail(['draft', '--to', 'a@b.cd', '--subject', 's', '--body', 'b', '-o', 'json'], [{ match: '/me/messages', method: 'POST', status: 201, body: { id: 'd1' } }]);
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout).id).toBe('d1');
  });

  it('list --all honours --page-limit', async () => {
    const r = await mail(['list', '--all', '--page-limit', '1', '-o', 'json'], [
      { match: '/me/messages', body: { value: [{ id: '1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=1' } },
    ]);
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(1);
    expect(JSON.parse(r.stdout)).toEqual([{ id: '1' }]);
  });

  it('folders, rules, attachments, drafts and shared-list accept --top', async () => {
    const routes: FakeRoute[] = [{ match: '.', body: { value: [] } }];
    for (const args of [['folders'], ['rules'], ['attachments', 'm1'], ['drafts'], ['shared-list', 'box@x.io']]) {
      const r = await mail([...args, '--top', '7', '-o', 'json'], routes);
      expect(r.code, args.join(' ')).toBe(0);
      expect(param(r.requests.at(-1)!.url, '$top'), args.join(' ')).toBe('7');
      expect(JSON.parse(r.stdout)).toEqual({ value: [] });
    }
  });
});
