import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import type { Command } from 'commander';
import { runCommand, jsonResponse, type RunOptions } from './helpers/cli.js';
import { entryHome, fakeJwt, runEntry } from './helpers/entry.js';
import { registerMailCommands } from '../src/commands/mail.js';
import { registerCalendarCommands } from '../src/commands/calendar.js';
import { registerApplyCommand } from '../src/commands/apply.js';
import type { GraphClient } from '../src/graph-client.js';
import { setErrorFormat } from '../src/errors.js';
import { beginPlan, planDigest, type Plan } from '../src/plan.js';
import { saveProfile } from '../src/profiles.js';

const TENANT = '11111111-2222-3333-4444-555555555555';
const OID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
const TOKEN = fakeJwt({ tid: TENANT, oid: OID, upn: 'ceo@contoso.com', aud: 'https://graph.microsoft.com' });

const register = (p: Command, g: () => GraphClient) => {
  registerMailCommands(p, g);
  registerCalendarCommands(p, g);
  registerApplyCommand(p, g);
};
const run = (argv: string[], opts: RunOptions = {}) => runCommand(register, argv, { token: TOKEN, ...opts });
const lastError = (stderr: string) => JSON.parse(stderr.trim().split('\n').pop()!) as { error: { code: string } };

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'plan-test-'));
  setErrorFormat('json');
});
afterEach(() => {
  setErrorFormat('text');
  vi.useRealTimers();
  fs.rmSync(dir, { recursive: true, force: true });
});

const readPlanFile = (file: string): Plan => JSON.parse(fs.readFileSync(file, 'utf8')) as Plan;

/** `mail draft` under --plan; returns the plan file and its digest. */
async function planDraft(extra: string[] = [], opts: RunOptions = {}) {
  const file = path.join(dir, 'draft.plan.json');
  const r = await run(['mail', 'draft', '--to', 'cfo@contoso.com', '--subject', 'Q3', '--body', 'Numbers attached.', ...extra],
    { plan: { file, command: 'mail draft' }, ...opts });
  return { r, file, plan: fs.existsSync(file) ? readPlanFile(file) : undefined };
}

const DRAFT_ROUTE: [string, string, () => unknown] = ['GET', '/me/messages/D1', () => ({ isDraft: true, '@odata.etag': 'W/"ck1"' })];

describe('--plan (AC18)', () => {
  it('writes a 0600 plan with the exact request and sends no write', async () => {
    const { r, file, plan } = await planDraft(['--bcc', 'audit@contoso.com']);
    expect(r.exitCode).toBe(0);
    expect(r.calls.filter(c => c.method !== 'GET')).toEqual([]);
    expect(fs.statSync(file).mode & 0o777).toBe(0o600);
    expect(plan!.digest).toBe(planDigest(plan!));
    expect(JSON.parse(r.stdout)).toMatchObject({ plan: file, command: 'mail draft', digest: plan!.digest, requests: 1 });
    expect(plan).toMatchObject({ v: 1, command: 'mail draft', identity: { tenantId: TENANT, accountId: OID } });
    expect(plan!.requests).toHaveLength(1);
    const req = plan!.requests[0];
    expect(req).toMatchObject({ method: 'POST', path: '/me/messages' });
    expect(req.headers.Prefer).toMatch(/IdType="ImmutableId"/);
    expect(JSON.parse(req.body!)).toMatchObject({ bccRecipients: [{ emailAddress: { address: 'audit@contoso.com' } }] });
    expect(Date.parse(plan!.expiresAt) - Date.parse(plan!.createdAt)).toBe(15 * 60_000);
  });

  it('apply sends the planned body byte for byte, with its headers', async () => {
    const { plan, file } = await planDraft();
    const r = await run(['apply', file, '--expect-digest', plan!.digest], { routes: [['POST', '/me/messages', () => ({ id: 'D9' })]] });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(1);
    expect(r.calls[0].rawBody).toBe(plan!.requests[0].body);
    expect(r.calls[0].headers.Prefer).toBe(plan!.requests[0].headers.Prefer);
    expect(JSON.parse(r.stdout)).toMatchObject({ outcome: 'APPLIED', command: 'mail draft', digest: plan!.digest, responses: [{ method: 'POST', status: 200, id: 'D9' }] });
  });

  it('records the draft version for update-draft and send-draft as If-Match', async () => {
    const upd = path.join(dir, 'upd.json');
    const r1 = await run(['mail', 'update-draft', 'D1', '--subject', 'v2'], { plan: { file: upd, command: 'mail update-draft' }, routes: [DRAFT_ROUTE] });
    expect(r1.exitCode).toBe(0);
    const p1 = readPlanFile(upd);
    expect(p1.target).toEqual({ id: 'D1', version: 'W/"ck1"' });
    expect(p1.requests[0]).toMatchObject({ method: 'PATCH', path: '/me/messages/D1', headers: { 'If-Match': 'W/"ck1"' } });

    const send = path.join(dir, 'send.json');
    const r2 = await run(['mail', 'send-draft', 'D1', '--idempotency-key', 'k1'], { plan: { file: send, command: 'mail send-draft' }, routes: [DRAFT_ROUTE] });
    expect(r2.exitCode).toBe(0);
    expect(r2.stderr).not.toMatch(/Draft sent/);
    const p2 = readPlanFile(send);
    expect(p2.payload).toEqual({ messageId: 'D1' });
    expect(p2.requests[0]).toMatchObject({ method: 'POST', path: '/me/messages/D1/send', headers: { 'If-Match': 'W/"ck1"' } });
  });

  it('refuses a command that cannot be planned, and an upload session', async () => {
    expect(() => beginPlan({ file: path.join(dir, 'x'), command: 'mail send', ttlMinutes: 15, profile: null, client: () => { throw new Error('unused'); } }))
      .toThrow(/not supported by "mail send"/);
    const big = path.join(dir, 'big.bin');
    fs.writeFileSync(big, Buffer.alloc(3 * 1024 * 1024 + 1));
    const { r } = await planDraft(['--attach', big], { routes: [['POST', '/me/messages', () => ({ id: 'D1' })]] });
    expect(r.exitCode).toBe(2);
    expect(r.stderr).toMatch(/upload session cannot be planned/);
    expect(r.calls.filter(c => c.method !== 'GET')).toEqual([]);
  });
});

describe('apply refusals (AC19)', () => {
  it('refuses a tampered plan or another digest, sending nothing', async () => {
    const { plan, file } = await planDraft();
    const r1 = await run(['apply', file, '--expect-digest', 'f'.repeat(64)]);
    expect(r1.exitCode).toBe(2);
    expect(lastError(r1.stderr).error.code).toBe('plan_digest_mismatch');

    fs.writeFileSync(file, JSON.stringify({ ...plan, requests: [{ ...plan!.requests[0], body: '{"subject":"other"}' }] }));
    const r2 = await run(['apply', file, '--expect-digest', plan!.digest]);
    expect(r2.exitCode).toBe(2);
    expect(lastError(r2.stderr).error.code).toBe('plan_digest_mismatch');

    fs.writeFileSync(file, '{"v":1}');
    const r3 = await run(['apply', file, '--expect-digest', plan!.digest]);
    expect(r3.exitCode).toBe(2);
    expect(lastError(r3.stderr).error.code).toBe('plan_invalid');
    expect([...r1.calls, ...r2.calls, ...r3.calls]).toEqual([]);
  });

  it('refuses when an attachment changed with the same name and size', async () => {
    const att = path.join(dir, 'report.txt');
    fs.writeFileSync(att, 'aaaa');
    const { plan, file } = await planDraft(['--attach', att]);
    expect(plan!.attachments).toEqual([expect.objectContaining({ name: 'report.txt', size: 4, path: att })]);
    fs.writeFileSync(att, 'bbbb');
    const r = await run(['apply', file, '--expect-digest', plan!.digest]);
    expect(r.exitCode).toBe(2);
    expect(lastError(r.stderr).error.code).toBe('plan_attachment_changed');
    expect(r.calls).toEqual([]);
  });

  it('refuses another identity, an expired plan, and a command outside the profile ceiling', async () => {
    const { plan, file } = await planDraft();
    const other = fakeJwt({ tid: TENANT, oid: '99999999-8888-7777-6666-555555555555', aud: 'https://graph.microsoft.com' });
    const r1 = await run(['apply', file, '--expect-digest', plan!.digest], { token: other });
    expect(r1.exitCode).toBe(3);
    expect(lastError(r1.stderr).error.code).toBe('token_identity_mismatch');

    saveProfile('readonly-host', { type: 'host', tenantId: TENANT, scopes: ['Mail.Read'] });
    const r2 = await run(['apply', file, '--expect-digest', plan!.digest], { env: { MS365_CLI_PROFILE: 'readonly-host' } });
    expect(r2.exitCode).toBe(4);
    expect(lastError(r2.stderr).error.code).toBe('scope_ceiling');

    vi.useFakeTimers({ toFake: ['Date'], now: Date.parse(plan!.expiresAt) + 1 });
    const r3 = await run(['apply', file, '--expect-digest', plan!.digest]);
    expect(r3.exitCode).toBe(2);
    expect(lastError(r3.stderr).error.code).toBe('plan_expired');
    expect([...r1.calls, ...r2.calls, ...r3.calls].filter(c => c.method !== 'GET')).toEqual([]);
  });

  it('maps 412 to exit 4 version_changed without a retry', async () => {
    const file = path.join(dir, 'upd.json');
    await run(['mail', 'update-draft', 'D1', '--subject', 'v2'], { plan: { file, command: 'mail update-draft' }, routes: [DRAFT_ROUTE] });
    const plan = readPlanFile(file);
    const r = await run(['apply', file, '--expect-digest', plan.digest], {
      routes: [['PATCH', '/me/messages/D1', () => jsonResponse({ error: { code: 'ErrorIrresolvableConflict', message: 'changed' } }, 412)]],
    });
    expect(r.exitCode).toBe(4);
    expect(lastError(r.stderr).error.code).toBe('version_changed');
    expect(r.calls).toHaveLength(1);
    expect(r.calls[0].headers['If-Match']).toBe('W/"ck1"');
  });

  it('applies a plan once: the digest is the idempotency key', async () => {
    const { plan, file } = await planDraft();
    const routes: RunOptions['routes'] = [['POST', '/me/messages', () => ({ id: 'D9' })]];
    const r1 = await run(['apply', file, '--expect-digest', plan!.digest], { routes });
    const r2 = await run(['apply', file, '--expect-digest', plan!.digest], { routes });
    expect(r2.exitCode).toBe(0);
    expect(r2.calls).toEqual([]);
    expect(r2.stderr).toMatch(/already applied/);
    expect(JSON.parse(r2.stdout)).toEqual(JSON.parse(r1.stdout));
  });
});

describe('If-Match, --bcc and --body-file (AC20)', () => {
  it('update-draft sends an explicit --if-match, or the version read with auto', async () => {
    const r1 = await run(['mail', 'update-draft', 'D1', '--subject', 'x', '--if-match', 'W/"given"'], { routes: [DRAFT_ROUTE] });
    expect(r1.calls.find(c => c.method === 'PATCH')!.headers['If-Match']).toBe('W/"given"');
    const r2 = await run(['mail', 'update-draft', 'D1', '--subject', 'x', '--if-match', 'auto'], { routes: [DRAFT_ROUTE] });
    expect(r2.calls.find(c => c.method === 'PATCH')!.headers['If-Match']).toBe('W/"ck1"');
    const r3 = await run(['mail', 'update-draft', 'D1', '--subject', 'x'], { routes: [DRAFT_ROUTE] });
    expect(r3.calls.find(c => c.method === 'PATCH')!.headers['If-Match']).toBeUndefined();
  });

  it('calendar update/delete send If-Match; a 412 is version_changed', async () => {
    const routes: RunOptions['routes'] = [
      ['GET', '/me/events/E1', () => ({ changeKey: 'ek1' })],
      ['PATCH', '/me/events/E1', () => jsonResponse({ error: { code: 'ErrorIrresolvableConflict', message: 'changed' } }, 412)],
    ];
    const r1 = await run(['calendar', 'update', 'E1', '--subject', 'x', '--if-match', 'auto'], { routes });
    expect(r1.calls.find(c => c.method === 'PATCH')!.headers['If-Match']).toBe('W/"ek1"');
    expect(r1.exitCode).toBe(4);
    expect(lastError(r1.stderr).error.code).toBe('version_changed');
    const r2 = await run(['calendar', 'delete', 'E1', '--if-match', 'W/"ek2"'], { tty: true, answer: 'y' });
    expect(r2.calls.find(c => c.method === 'DELETE')?.headers['If-Match']).toBe('W/"ek2"');
  });

  it('--body-file sends the file bytes exactly; with --body it is a usage error', async () => {
    const body = path.join(dir, 'body.txt');
    fs.writeFileSync(body, 'Line 1\r\n  "quoted" ünïcode\n');
    const r1 = await run(['mail', 'draft', '--to', 'cfo@contoso.com', '--subject', 's', '--body-file', body], { routes: [['POST', '/me/messages', () => ({ id: 'D1' })]] });
    expect(r1.exitCode).toBe(0);
    expect((r1.calls[0].body as { body: { content: string } }).body.content).toBe('Line 1\r\n  "quoted" ünïcode\n');
    const r2 = await run(['mail', 'draft', '--to', 'cfo@contoso.com', '--subject', 's', '--body', 'x', '--body-file', body]);
    expect(r2.exitCode).toBe(2);
    const r3 = await run(['mail', 'draft', '--to', 'cfo@contoso.com', '--subject', 's']);
    expect(r3.exitCode).toBe(2);
    expect(r3.stderr).toMatch(/--body/);
    const r4 = await run(['mail', 'draft', '--to', 'cfo@contoso.com', '--subject', 's', '--body-file', '-'], { env: { MS365_CLI_MCP: '1' } });
    expect(r4.exitCode).toBe(2);
    expect([...r2.calls, ...r3.calls, ...r4.calls]).toEqual([]);
  });
});

describe('--plan through the CLI entry (spawned)', () => {
  it('refuses --dry-run, plans with no write, and apply sends the one request', () => {
    const { home, configDir } = entryHome();
    const token = fakeJwt({ tid: TENANT, oid: OID, scp: 'Mail.ReadWrite Mail.Send User.Read', aud: 'https://graph.microsoft.com', exp: Math.floor(Date.now() / 1000) + 3600 });
    const env = { MS365_CLI_TOKEN: token };
    const file = path.join(home, 'p.json');
    const draft = ['mail', 'draft', '--to', 'cfo@contoso.com', '--subject', 'Q3', '--body', 'b'];
    const dry = runEntry(['--plan', file, '--dry-run', ...draft], { home, configDir, env });
    expect(dry.status, dry.stderr).toBe(2);
    expect(dry.error?.code).toBe('usage');

    const planned = runEntry(['--plan', file, '--plan-ttl', '5', ...draft], { home, configDir, env });
    expect(planned.status, planned.stderr).toBe(0);
    expect(planned.fetches).toEqual([]);
    const plan = readPlanFile(file);
    expect(Date.parse(plan.expiresAt) - Date.parse(plan.createdAt)).toBe(5 * 60_000);

    const applied = runEntry(['apply', file, '--expect-digest', plan.digest], {
      home, configDir, env, routes: [{ method: 'POST', match: '/me/messages', status: 201, body: { id: 'D1' } }],
    });
    expect(applied.status, applied.stderr).toBe(0);
    expect(applied.fetches).toEqual(['POST https://graph.microsoft.com/v1.0/me/messages']);
    expect(JSON.parse(applied.stdout)).toMatchObject({ outcome: 'APPLIED', responses: [{ status: 201, id: 'D1' }] });
  }, 60_000);
});
