import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import type { Command } from 'commander';
import { runCommand, type RunOptions } from './helpers/cli.js';
import { fakeJwt } from './helpers/entry.js';
import { registerCalendarCommands } from '../src/commands/calendar.js';
import { registerApplyCommand } from '../src/commands/apply.js';
import type { GraphClient } from '../src/graph-client.js';
import { setErrorFormat } from '../src/errors.js';
import { planDigest, type Plan } from '../src/plan.js';
import { saveProfile } from '../src/profiles.js';

// AC23 (AH-7): calendar writes read the event first, refuse a series master without --series, never take --series in a
// host plan, record scope and notifies in plans, page occurrences, and let only the organizer cancel.

const TOKEN = fakeJwt({ tid: '11111111-2222-3333-4444-555555555555', oid: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee', aud: 'https://graph.microsoft.com' });
const register = (p: Command, g: () => GraphClient) => {
  registerCalendarCommands(p, g);
  registerApplyCommand(p, g);
};
const run = (argv: string[], opts: RunOptions = {}) => runCommand(register, argv, { token: TOKEN, ...opts });
const lastError = (stderr: string) => JSON.parse(stderr.trim().split('\n').pop()!) as { error: { code: string; message: string } };

const ATTENDEES = [
  { emailAddress: { address: 'cfo@contoso.com' }, type: 'required' },
  { emailAddress: { address: 'board@contoso.com' }, type: 'optional' },
];
type Route = [string, RegExp, () => unknown];
const event = (fields: Record<string, unknown>): Route =>
  ['GET', /\/me\/events\/E1\?/, () => ({ '@odata.etag': 'W/"ek1"', attendees: ATTENDEES, ...fields })];
const MASTER = event({ type: 'seriesMaster', isOrganizer: true });
const OCCURRENCE = event({ type: 'occurrence', seriesMasterId: 'S1', isOrganizer: true });
const ATTENDING = event({ type: 'singleInstance', isOrganizer: false });

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cal-plan-test-'));
  setErrorFormat('json');
});
afterEach(() => {
  setErrorFormat('text');
  fs.rmSync(dir, { recursive: true, force: true });
});

const planFile = (name: string) => path.join(dir, `${name}.plan.json`);
const readPlan = (file: string) => JSON.parse(fs.readFileSync(file, 'utf8')) as Plan;
const writes = (r: { calls: { method: string }[] }) => r.calls.filter(c => c.method !== 'GET');

describe('series masters (AC23)', () => {
  for (const argv of [['calendar', 'delete', 'E1'], ['calendar', 'update', 'E1', '--subject', 'X'], ['calendar', 'cancel', 'E1']]) {
    it(`${argv[1]} on a series master without --series exits 2 and sends nothing`, async () => {
      const r = await run(argv, { routes: [MASTER], tty: true, answer: 'y' });
      expect(r.exitCode).toBe(2);
      expect(lastError(r.stderr).error.code).toBe('series_refused');
      expect(lastError(r.stderr).error.message).toContain('calendar instances E1');
      expect(writes(r)).toEqual([]);
    });
  }

  it('delete --series on a terminal deletes the series, with If-Match from the read', async () => {
    const r = await run(['calendar', 'delete', 'E1', '--series', '--if-match', 'auto'], { routes: [MASTER, ['DELETE', /\/me\/events\/E1$/, () => ({})]], tty: true, answer: 'y' });
    expect(r.exitCode).toBe(0);
    expect(writes(r)).toHaveLength(1);
    expect(r.calls[1].headers['If-Match']).toBe('W/"ek1"');
  });

  it('a host plan with --series is refused before any read', async () => {
    const file = planFile('series');
    const r = await run(['calendar', 'delete', 'E1', '--series'], { routes: [MASTER], plan: { file, command: 'calendar delete', host: true } });
    expect(r.exitCode).toBe(2);
    expect(lastError(r.stderr).error.code).toBe('series_refused');
    expect(r.calls).toEqual([]);
    expect(fs.existsSync(file)).toBe(false);
  });
});

describe('plans record scope and notifies (AC23)', () => {
  it('update: an occurrence the organizer changes notifies every attendee; bound to the version', async () => {
    const file = planFile('update');
    const r = await run(['calendar', 'update', 'E1', '--subject', 'Moved'], { routes: [OCCURRENCE], plan: { file, command: 'calendar update', host: true } });
    expect(r.exitCode).toBe(0);
    expect(writes(r)).toEqual([]);
    const plan = readPlan(file);
    expect(plan.target).toEqual({ id: 'E1', version: 'W/"ek1"' });
    expect(plan.payload).toEqual({
      eventId: 'E1', patch: { subject: 'Moved' }, scope: 'occurrence',
      notifies: [{ address: 'cfo@contoso.com', type: 'required' }, { address: 'board@contoso.com', type: 'optional' }],
    });
    expect(plan.requests[0]).toMatchObject({ method: 'PATCH', headers: { 'If-Match': 'W/"ek1"' } });
  });

  it('delete: an attendee\'s own copy notifies nobody; planning asks nothing', async () => {
    const file = planFile('delete');
    const r = await run(['calendar', 'delete', 'E1'], { routes: [ATTENDING], plan: { file, command: 'calendar delete', host: true } });
    expect(r.exitCode).toBe(0);
    expect(readPlan(file).payload).toEqual({ eventId: 'E1', scope: 'single', notifies: [] });
  });

  it('cancel: records the comment and the attendees it reaches', async () => {
    const file = planFile('cancel');
    const r = await run(['calendar', 'cancel', 'E1', '--comment', 'Board moved'], { routes: [OCCURRENCE], plan: { file, command: 'calendar cancel', host: true } });
    expect(r.exitCode).toBe(0);
    const plan = readPlan(file);
    expect(plan.payload).toMatchObject({ eventId: 'E1', comment: 'Board moved', scope: 'occurrence' });
    expect(plan.requests[0]).toMatchObject({ method: 'POST', body: JSON.stringify({ comment: 'Board moved' }) });
    expect(plan.requests[0].path).toMatch(/\/me\/events\/E1\/cancel$/);
  });

  it('create: invitees are who it notifies; the transactionId travels in the plan', async () => {
    const file = planFile('create');
    const r = await run(['calendar', 'create', '--subject', 'S', '--start', '2026-10-20T09:00', '--end', '2026-10-20T09:30', '--timezone', 'UTC',
      '--attendees', 'cfo@contoso.com', '--idempotency-key', 'k-1'], { plan: { file, command: 'calendar create', host: true } });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toEqual([]);
    const plan = readPlan(file);
    expect(plan.payload).toMatchObject({ scope: 'single', notifies: [{ address: 'cfo@contoso.com', type: 'required' }], transactionId: 'k-1' });
    expect(JSON.parse(String(plan.requests[0].body))).toMatchObject({ transactionId: 'k-1' });
  });

  it('apply of a delete plan without a terminal is refused (exit 7) and sends nothing; on a terminal it asks', async () => {
    const file = planFile('apply-delete');
    await run(['calendar', 'delete', 'E1'], { routes: [OCCURRENCE], plan: { file, command: 'calendar delete', host: true } });
    const { digest } = readPlan(file);
    const routes: Route[] = [['DELETE', /\/me\/events\/E1$/, () => ({})]];
    const refused = await run(['apply', file, '--expect-digest', digest], { routes });
    expect(refused.exitCode).toBe(7);
    expect(refused.calls).toEqual([]);
    const applied = await run(['apply', file, '--expect-digest', digest], { routes, tty: true, answer: 'y' });
    expect(applied.exitCode).toBe(0);
    expect(writes(applied)).toHaveLength(1);
    expect(writes(applied)[0]).toMatchObject({ method: 'DELETE' });
    expect(applied.calls.at(-1)!.headers['If-Match']).toBe('W/"ek1"');
  });
});

describe('apply without a terminal under a host profile (OQ-8 b)', () => {
  const OTHER = fakeJwt({ tid: '11111111-2222-3333-4444-555555555555', oid: '99999999-8888-7777-6666-555555555555', aud: 'https://graph.microsoft.com' });
  const HOST = { MS365_CLI_PROFILE: 'chief-host' };
  const DELETE_ROUTE: Route = ['DELETE', /\/me\/events\/E1$/, () => ({})];
  beforeEach(() => {
    saveProfile('chief-host', { type: 'host', tenantId: '11111111-2222-3333-4444-555555555555', scopes: ['Calendars.ReadWrite'] });
    saveProfile('chief-read', { type: 'host', tenantId: '11111111-2222-3333-4444-555555555555', scopes: ['Calendars.Read'] });
    saveProfile('ceo-user', { type: 'user', scopes: ['Calendars.ReadWrite'] });
  });
  afterEach(() => { vi.useRealTimers(); });

  /** A calendar delete (or cancel) plan made in host mode; returns its file and digest. */
  async function hostPlan(name: string, argv = ['calendar', 'delete', 'E1']) {
    const file = planFile(name);
    const made = await run(argv, { routes: [OCCURRENCE], plan: { file, command: `calendar ${argv[1]}`, host: true } });
    expect(made.exitCode).toBe(0);
    return { file, plan: readPlan(file) };
  }
  /** Rewrite the plan with other requests and/or label and a matching digest (what a hand-written plan can do). */
  function forge(file: string, change: Partial<Plan>): string {
    const { digest: _old, ...rest } = { ...readPlan(file), ...change };
    const digest = planDigest(rest);
    fs.writeFileSync(file, JSON.stringify({ ...rest, digest }));
    return digest;
  }

  it('applies an approved delete: host profile, the approved digest, unexpired, same identity, inside the ceiling', async () => {
    const { file, plan } = await hostPlan('host-delete');
    const r = await run(['apply', file, '--expect-digest', plan.digest], { routes: [DELETE_ROUTE], env: HOST });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toEqual([expect.objectContaining({ method: 'DELETE', headers: expect.objectContaining({ 'If-Match': 'W/"ek1"' }) })]);
    expect(JSON.parse(r.stdout)).toMatchObject({ outcome: 'APPLIED', command: 'calendar delete', digest: plan.digest });
  });

  it('applies an approved cancel the same way', async () => {
    const { file, plan } = await hostPlan('host-cancel', ['calendar', 'cancel', 'E1', '--comment', 'Board moved to Friday']);
    const r = await run(['apply', file, '--expect-digest', plan.digest], { routes: [['POST', /\/me\/events\/E1\/cancel$/, () => ({})]], env: HOST });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toEqual([expect.objectContaining({ method: 'POST', body: { comment: 'Board moved to Friday' } })]);
  });

  it('without a host profile it stays exit 7, sends nothing, and the same plan then applies under the host profile', async () => {
    const { file, plan } = await hostPlan('no-host');
    const none = await run(['apply', file, '--expect-digest', plan.digest], { routes: [DELETE_ROUTE] });
    const user = await run(['apply', file, '--expect-digest', plan.digest], { routes: [DELETE_ROUTE], env: { MS365_CLI_PROFILE: 'ceo-user' } });
    for (const r of [none, user]) {
      expect(r.exitCode).toBe(7);
      expect(lastError(r.stderr).error.code).toBe('confirmation_refused');
      expect(r.calls).toEqual([]);
    }
    expect(lastError(user.stderr).error.message).toContain('type: host');
    const applied = await run(['apply', file, '--expect-digest', plan.digest], { routes: [DELETE_ROUTE], env: HOST });
    expect(applied.exitCode).toBe(0);
    expect(applied.calls).toHaveLength(1);
  });

  it('refuses a missing or wrong --expect-digest, an expired plan, another identity and a ceiling without the write, sending nothing', async () => {
    const { file, plan } = await hostPlan('refusals');
    const missing = await run(['apply', file], { routes: [DELETE_ROUTE], env: HOST });
    expect(missing.exitCode).toBe(2);
    const wrong = await run(['apply', file, '--expect-digest', 'f'.repeat(64)], { routes: [DELETE_ROUTE], env: HOST });
    expect(wrong.exitCode).toBe(2);
    expect(lastError(wrong.stderr).error.code).toBe('plan_digest_mismatch');
    const other = await run(['apply', file, '--expect-digest', plan.digest], { routes: [DELETE_ROUTE], env: HOST, token: OTHER });
    expect(other.exitCode).toBe(3);
    expect(lastError(other.stderr).error.code).toBe('token_identity_mismatch');
    const ceiling = await run(['apply', file, '--expect-digest', plan.digest], { routes: [DELETE_ROUTE], env: { MS365_CLI_PROFILE: 'chief-read' } });
    expect(ceiling.exitCode).toBe(4);
    expect(lastError(ceiling.stderr).error.code).toBe('scope_ceiling');
    vi.useFakeTimers({ toFake: ['Date'], now: Date.parse(plan.expiresAt) + 1 });
    const expired = await run(['apply', file, '--expect-digest', plan.digest], { routes: [DELETE_ROUTE], env: HOST });
    expect(expired.exitCode).toBe(2);
    expect(lastError(expired.stderr).error.code).toBe('plan_expired');
    expect([...missing.calls, ...wrong.calls, ...other.calls, ...ceiling.calls, ...expired.calls]).toEqual([]);
  });

  it('refuses a mislabelled plan (E27) and one that names no identity: exit 7, nothing sent', async () => {
    const { file, plan } = await hostPlan('mislabel');
    const del = plan.requests[0];
    const relabelled = forge(file, { command: 'calendar update' });
    const r1 = await run(['apply', file, '--expect-digest', relabelled], { routes: [DELETE_ROUTE], env: HOST });
    const mixed = forge(file, { command: 'calendar delete', requests: [{ ...del, method: 'PATCH', body: '{"subject":"x"}' }, del] });
    const r2 = await run(['apply', file, '--expect-digest', mixed], { routes: [DELETE_ROUTE], env: HOST });
    const anonymous = forge(file, { requests: [del], identity: { tenantId: null, accountId: null, profile: null } });
    const r3 = await run(['apply', file, '--expect-digest', anonymous], { routes: [DELETE_ROUTE], env: HOST });
    for (const r of [r1, r2, r3]) {
      expect(r.exitCode).toBe(7);
      expect(lastError(r.stderr).error.code).toBe('confirmation_refused');
      expect(r.calls).toEqual([]);
    }
  });

  it('a terminal still asks under the host profile; --dry-run previews without asking', async () => {
    const { file, plan } = await hostPlan('tty');
    const declined = await run(['apply', file, '--expect-digest', plan.digest], { routes: [DELETE_ROUTE], env: HOST, tty: true, answer: 'n' });
    expect(declined.exitCode).toBe(7);
    expect(declined.calls).toEqual([]);
    const preview = await run(['apply', file, '--expect-digest', plan.digest], { routes: [DELETE_ROUTE], env: HOST, dryRun: true });
    expect(preview.exitCode).toBe(0);
    expect(preview.calls).toEqual([]);
    expect(preview.stderr).toContain('would ask');
  });
});

describe('calendar cancel (AC23)', () => {
  it('a non-organizer exits 2 with a decline hint', async () => {
    const r = await run(['calendar', 'cancel', 'E1'], { routes: [ATTENDING], tty: true, answer: 'y' });
    expect(r.exitCode).toBe(2);
    expect(lastError(r.stderr).error).toMatchObject({ code: 'not_organizer' });
    expect(lastError(r.stderr).error.message).toContain('calendar decline E1');
    expect(writes(r)).toEqual([]);
  });

  it('the organizer, on a terminal, POSTs /cancel with the comment', async () => {
    const r = await run(['calendar', 'cancel', 'E1', '--comment', 'Sorry'], { routes: [OCCURRENCE, ['POST', /\/cancel$/, () => ({})]], tty: true, answer: 'y' });
    expect(r.exitCode).toBe(0);
    expect(writes(r)).toEqual([expect.objectContaining({ method: 'POST', body: { comment: 'Sorry' } })]);
    expect(r.stderr).toContain('Event cancelled');
  });
});

describe('calendar instances (AC23)', () => {
  it('pages the occurrences of a series with --all', async () => {
    let page = 0;
    const r = await run(['calendar', 'instances', 'S1', '--start', '2026-10-01T00:00', '--end', '2026-11-01T00:00', '--timezone', 'UTC', '--all', '-o', 'json'], {
      routes: [['GET', /\/me\/events\/S1\/instances/, () => (page++ === 0
        ? { value: [{ id: 'O1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/events/S1/instances?$skiptoken=p2' }
        : { value: [{ id: 'O2' }] })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(2);
    expect(r.calls[0].url).toContain('startDateTime=2026-10-01T00%3A00%3A00.000Z');
    expect(JSON.stringify(JSON.parse(r.stdout))).toContain('O2');
  });
});
