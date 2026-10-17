import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import type { Command } from 'commander';
import { runCommand, type RunOptions, type Route } from './helpers/cli.js';
import { fakeJwt } from './helpers/entry.js';
import { registerMailCommands } from '../src/commands/mail.js';
import { registerCalendarCommands } from '../src/commands/calendar.js';
import { registerDescribeCommand } from '../src/commands/describe.js';
import type { GraphClient } from '../src/graph-client.js';
import { setErrorFormat } from '../src/errors.js';
import { listEntries } from '../src/idempotency.js';
import type { Plan } from '../src/plan.js';

const TOKEN = fakeJwt({ tid: '11111111-2222-3333-4444-555555555555', oid: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee', aud: 'https://graph.microsoft.com' });
const IMID = '<m1@contoso.com>';
const register = (p: Command, g: () => GraphClient) => {
  registerMailCommands(p, g);
  registerCalendarCommands(p, g);
  registerDescribeCommand(p);
};
const run = (argv: string[], opts: RunOptions = {}) => runCommand(register, argv, opts);
const lastError = (stderr: string) => JSON.parse(stderr.trim().split('\n').pop()!).error as Record<string, unknown>;
const send = (...extra: string[]) => ['mail', 'send', '--to', 'cfo@contoso.com', '--subject', 'Q3', '--body', 'Numbers.', ...extra];
const DRAFT: Route = ['POST', '/me/messages', () => ({ id: 'D1', internetMessageId: IMID })];
const LOST: Route = ['POST', /\/send$/, () => { throw new TypeError('fetch failed'); }];
/** Folder searches: `hit` answers with one message, the others with none. */
const folders = (hit?: string): Route[] => ['sentitems', 'outbox', 'drafts'].map(f => [
  'GET', `/me/mailFolders/${f}/messages`,
  () => ({ value: f === hit ? [{ id: 'S1', sentDateTime: '2026-10-17T09:00:00Z', webLink: 'https://outlook.example/S1', internetMessageId: IMID }] : [] }),
]);

let cfg: string;
beforeEach(() => {
  setErrorFormat('json');
  cfg = fs.mkdtempSync(path.join(os.tmpdir(), 'send-outcome-'));
  vi.stubEnv('MS365_CLI_CONFIG_DIR', cfg);
});
afterEach(() => {
  setErrorFormat('text');
  vi.unstubAllEnvs();
  fs.rmSync(cfg, { recursive: true, force: true });
});

describe('honest send (AC21)', () => {
  it('-o json reports ACCEPTED with the draft correlation and delivery unknown', async () => {
    const r = await run([...send('--idempotency-key', 'k1'), '-o', 'json'], { routes: [DRAFT] });
    expect(r.exitCode).toBe(0);
    expect(r.calls.map(c => `${c.method} ${c.path}`)).toEqual(['POST /me/messages', 'POST /me/messages/D1/send']);
    const out = JSON.parse(r.stdout);
    expect(out).toEqual({
      outcome: 'ACCEPTED', acceptedAt: expect.any(String), delivery: 'unknown',
      correlation: { draftId: 'D1', internetMessageId: IMID, idempotencyKey: 'k1' },
    });
    expect(listEntries()[0]).toMatchObject({ key: 'k1', status: 'done', result: { correlation: { internetMessageId: IMID } } });
  });

  it('text says accepted by Exchange, never sent successfully', async () => {
    const r = await run([...send(), '-o', 'table'], { routes: [DRAFT] });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toMatch(/Accepted by Exchange for delivery \(internetMessageId <m1@contoso\.com>\); delivery is not confirmed/);
    expect(r.stderr + r.stdout).not.toMatch(/sent successfully/i);
  });

  it('a lost /send response is exit 1 outcome_unknown with maybeProcessed and the correlation; the draft is kept', async () => {
    const r = await run(send('--idempotency-key', 'k2'), { routes: [LOST, DRAFT] });
    expect(r.exitCode).toBe(1);
    expect(lastError(r.stderr)).toMatchObject({
      code: 'outcome_unknown', exitCode: 1, maybeProcessed: true,
      correlation: { draftId: 'D1', internetMessageId: IMID, idempotencyKey: 'k2' },
    });
    expect(lastError(r.stderr).fix).toContain(`mail reconcile --internet-message-id "${IMID}"`);
    expect(r.calls.some(c => c.method === 'DELETE')).toBe(false);
    expect(listEntries()[0]).toMatchObject({ key: 'k2', status: 'unknown', result: { correlation: { internetMessageId: IMID } } });
  });
});

describe('mail reconcile (AC22)', () => {
  const lose = () => run(send('--idempotency-key', 'k3'), { routes: [LOST, DRAFT] });

  it('finds the Sent Items match, returns SENT with evidence and moves the entry to done', async () => {
    await lose();
    const r = await run(['mail', 'reconcile', '--idempotency-key', 'k3', '-o', 'json'], { routes: folders('sentitems') });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].path).toContain(`internetMessageId eq '${IMID}'`);
    expect(JSON.parse(r.stdout)).toEqual({
      state: 'SENT', checkedAt: expect.any(String),
      evidence: { id: 'S1', folder: 'sentitems', sentDateTime: '2026-10-17T09:00:00Z', webLink: 'https://outlook.example/S1', internetMessageId: IMID },
    });
    expect(listEntries()[0]).toMatchObject({ key: 'k3', status: 'done', result: { reconciled: { id: 'S1', folder: 'sentitems' } } });
  });

  it('no match is NOT_FOUND_YET (never "not sent") and the entry stays unknown', async () => {
    await lose();
    const r = await run(['mail', 'reconcile', '--idempotency-key', 'k3', '-o', 'json'], { routes: folders() });
    expect(r.exitCode).toBe(0);
    expect(r.calls.map(c => c.path.split('?')[0])).toEqual(['/me/mailFolders/sentitems/messages', '/me/mailFolders/outbox/messages', '/me/mailFolders/drafts/messages']);
    expect(JSON.parse(r.stdout)).toMatchObject({ state: 'NOT_FOUND_YET', evidence: null });
    expect(listEntries()[0]).toMatchObject({ key: 'k3', status: 'unknown' });
  });

  it('a Drafts match is STILL_DRAFT; an Outbox match is NOT_FOUND_YET with evidence; --internet-message-id needs no ledger', async () => {
    const draft = await run(['mail', 'reconcile', '--internet-message-id', IMID, '-o', 'json'], { routes: folders('drafts') });
    expect(JSON.parse(draft.stdout)).toMatchObject({ state: 'STILL_DRAFT', evidence: { folder: 'drafts' } });
    const outbox = await run(['mail', 'reconcile', '--internet-message-id', IMID, '-o', 'json'], { routes: folders('outbox') });
    expect(JSON.parse(outbox.stdout)).toMatchObject({ state: 'NOT_FOUND_YET', evidence: { folder: 'outbox' } });
  });

  it('--internet-message-id settles the ledger entry that recorded it', async () => {
    await lose();
    await run(['mail', 'reconcile', '--internet-message-id', IMID, '-o', 'json'], { routes: folders('sentitems') });
    expect(listEntries()[0]).toMatchObject({ key: 'k3', status: 'done' });
  });

  it('refuses an unknown key, a missing handle, and both handles at once (exit 2, nothing sent)', async () => {
    for (const argv of [['--idempotency-key', 'nope'], [], ['--idempotency-key', 'k', '--internet-message-id', IMID]]) {
      const r = await run(['mail', 'reconcile', ...argv]);
      expect(r.exitCode).toBe(2);
      expect(r.calls).toEqual([]);
    }
  });
});

describe('plans for send, reply and forward (AH-5)', () => {
  let dir: string;
  beforeEach(() => { dir = fs.mkdtempSync(path.join(os.tmpdir(), 'send-plan-')); });
  afterEach(() => fs.rmSync(dir, { recursive: true, force: true }));
  const planOf = (file: string) => JSON.parse(fs.readFileSync(file, 'utf8')) as Plan;

  it('mail send plans the draft and its /send, bound by placeholder', async () => {
    const file = path.join(dir, 'p.json');
    const r = await run(send(), { token: TOKEN, plan: { file, command: 'mail send' } });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toEqual([]);
    const plan = planOf(file);
    expect(plan.requests.map(q => `${q.method} ${q.path}`)).toEqual(['POST /me/messages', 'POST /me/messages/%7B%7B0.id%7D%7D/send']);
    expect(plan.payload).toMatchObject({ subject: 'Q3', toRecipients: [{ emailAddress: { address: 'cfo@contoso.com' } }], attachments: [] });
  });

  it('mail reply binds the original message (id and version) without If-Match; --html cannot be planned', async () => {
    const file = path.join(dir, 'r.json');
    const routes: Route[] = [['GET', '/me/messages/M1', () => ({ id: 'M1', '@odata.etag': 'W/"ck7"' })]];
    const r = await run(['mail', 'reply', 'M1', '--body', 'Thanks.'], { token: TOKEN, plan: { file, command: 'mail reply' }, routes });
    expect(r.exitCode).toBe(0);
    const plan = planOf(file);
    expect(plan.target).toEqual({ id: 'M1', version: 'W/"ck7"' });
    expect(plan.requests).toHaveLength(1);
    expect(plan.requests[0]).toMatchObject({ method: 'POST', path: '/me/messages/M1/reply' });
    expect(plan.requests[0].headers['If-Match']).toBeUndefined();

    const html = await run(['mail', 'reply', 'M1', '--body', '<p>x</p>', '--html'], { token: TOKEN, plan: { file: path.join(dir, 'h.json'), command: 'mail reply' }, routes });
    expect(html.exitCode).toBe(2);
    expect(lastError(html.stderr).message).toMatch(/cannot be planned/);
  });
});

describe('describe (AH-6(5))', () => {
  it('names calendar create\'s provider idempotency and which commands take --plan', async () => {
    const r = await run(['describe', '--json']);
    const commands = JSON.parse(r.stdout).commands as Array<{ path: string; plan: boolean; providerIdempotency: string | null }>;
    const by = (p: string) => commands.find(c => c.path === p)!;
    expect(by('calendar create').providerIdempotency).toMatch(/transactionId/);
    expect(by('mail send').providerIdempotency).toBeNull();
    expect([by('mail send').plan, by('mail reply').plan, by('mail list').plan]).toEqual([true, true, false]);
  });
});
