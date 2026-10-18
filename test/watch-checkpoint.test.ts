import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { runCommand, jsonResponse, type RecordedCall } from './helpers/cli.js';
import { registerWatchCommands, watchDeps } from '../src/commands/watch.js';
import { setErrorFormat } from '../src/errors.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const MAIL = '/me/mailFolders/inbox/messages/delta';
const EVENTS = '/me/calendarView/delta';
const NOW = '2026-10-14T17:00:00.000Z';

let dir: string;
let env: Record<string, string>;
const lines = (stdout: string) => stdout.trim().split('\n').filter(Boolean).map(l => JSON.parse(l));
const query = (call: RecordedCall) => new URLSearchParams(call.url.split('?')[1] ?? '');
const run = (argv: string[], routes: Parameters<typeof runCommand>[2]['routes'] = []) =>
  runCommand(registerWatchCommands, ['watch', ...argv], { env, routes });
const blob = (cp: Record<string, unknown>) => Buffer.from(JSON.stringify(cp)).toString('base64url');
const decode = (b: string) => JSON.parse(Buffer.from(b, 'base64url').toString('utf8'));
const mailCheckpoint = (extra: Record<string, unknown> = {}) => blob({
  v: 1, resource: 'mail', endpoint: MAIL, deltaLink: `${GRAPH}${MAIL}?$deltatoken=OLD`, syncedAt: '2026-10-14T16:00:00.000Z', coverageStart: '2026-10-01T00:00:00.000Z', folder: 'inbox', ...extra,
});
const errorCode = (stderr: string) => JSON.parse(stderr.trim().split('\n').at(-1)!).error.code;

const baseline = [['GET', MAIL, () => ({ value: [{ id: 'M1', subject: 'a', changeKey: 'c1' }], '@odata.deltaLink': `${GRAPH}${MAIL}?$deltatoken=D1` })]] as never;
const changes = [['GET', MAIL, () => ({ value: [
  { id: 'M1', subject: 'a', changeKey: 'c2', lastModifiedDateTime: '2026-10-14T16:30:00Z', createdDateTime: '2026-10-01T00:00:00Z' },
  { id: 'M3', '@removed': { reason: 'changed' } },
], '@odata.deltaLink': `${GRAPH}${MAIL}?$deltatoken=D2` })]] as never;

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-cp-'));
  env = { MS365_CLI_CONFIG_DIR: dir, MS365_CLI_TIMEZONE: 'UTC' };
  vi.useFakeTimers({ toFake: ['Date'] });
  vi.setSystemTime(new Date(NOW));
});
afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
  setErrorFormat('text');
  fs.rmSync(dir, { recursive: true, force: true });
});

describe('watch --checkpoint (AH-4(1), AC 12)', () => {
  it('new: baseline, no cursor file, one checkpoint line with folder and coverage', async () => {
    const r = await run(['mail', '--checkpoint', 'new'], baseline);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(fs.existsSync(path.join(dir, 'watch'))).toBe(false);
    const out = lines(r.stdout);
    expect(out).toHaveLength(1);
    expect(out[0]).toMatchObject({ type: 'checkpoint', resource: 'mail', folder: 'inbox', coverageStart: NOW, coverageEnd: NOW, syncedAt: NOW });
    expect(decode(out[0].checkpoint)).toEqual({ v: 1, resource: 'mail', endpoint: MAIL, deltaLink: `${GRAPH}${MAIL}?$deltatoken=D1`, syncedAt: NOW, coverageStart: NOW, folder: 'inbox' });
  });

  it('an empty value is a baseline too; --include-existing prints backfill lines before the checkpoint', async () => {
    const r = await run(['mail', '--checkpoint', '', '--include-existing'], baseline);
    expect(lines(r.stdout).map(l => l.type)).toEqual(['backfill', 'checkpoint']);
  });

  it('resumes from @file and from @- (stdin), keeping coverageStart and writing no file', async () => {
    const file = path.join(dir, 'cp.txt');
    fs.writeFileSync(file, `${mailCheckpoint()}\n`);
    for (const value of [`@${file}`, '@-']) {
      vi.spyOn(watchDeps, 'readStdin').mockReturnValue(mailCheckpoint());
      const r = await run(['mail', '--checkpoint', value], changes);
      expect(r.exitCode, r.stderr).toBe(0);
      expect(r.calls.map(c => c.url)).toEqual([`${GRAPH}${MAIL}?$deltatoken=OLD`]);
      const out = lines(r.stdout);
      expect(out.map(l => l.type)).toEqual(['updated', 'deleted', 'checkpoint']);
      expect(out[2]).toMatchObject({ coverageStart: '2026-10-01T00:00:00.000Z', coverageEnd: NOW });
      expect(decode(out[2].checkpoint).deltaLink).toBe(`${GRAPH}${MAIL}?$deltatoken=D2`);
    }
    expect(fs.existsSync(path.join(dir, 'watch'))).toBe(false);
  });

  it.each([
    ['another resource', ['events', '--checkpoint', mailCheckpoint()]],
    ['another folder', ['mail', '--folder', 'archive', '--checkpoint', mailCheckpoint()]],
    ['another endpoint', ['mail', '--checkpoint', mailCheckpoint({ endpoint: '/me/mailFolders/archive/messages/delta' })]],
    ['a deltaLink to another resource', ['mail', '--checkpoint', mailCheckpoint({ deltaLink: `${GRAPH}/me/drive/root/delta?token=x` })]],
    ['a deltaLink off Graph', ['mail', '--checkpoint', mailCheckpoint({ deltaLink: 'https://evil.example.com/v1.0/me/mailFolders/inbox/messages/delta' })]],
    ['a deltaLink for another API version', ['mail', '--checkpoint', mailCheckpoint({ deltaLink: `https://graph.microsoft.com/beta${MAIL}?$deltatoken=x` })]],
    ['another window', ['events', '--start', '2026-10-20T00:00Z', '--checkpoint', blob({
      v: 1, resource: 'events', endpoint: EVENTS, deltaLink: `${GRAPH}${EVENTS}?$deltatoken=E`, syncedAt: NOW, coverageStart: NOW,
      window: { start: '2026-10-14T00:00:00.000Z', end: '2026-11-13T00:00:00.000Z' },
    })]],
  ])('%s → exit 2 checkpoint_mismatch, zero requests', async (_label, argv) => {
    setErrorFormat('json');
    const r = await run(argv);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
    expect(errorCode(r.stderr)).toBe('checkpoint_mismatch');
  });

  it("Graph's mailFolders('…') deltaLink form still matches, and still asks for immutable ids", async () => {
    const link = `${GRAPH}/me/mailFolders('AAMkFOLDER')/messages/delta?$deltatoken=x`;
    const r = await run(['mail', '--checkpoint', mailCheckpoint({ deltaLink: link })], [['GET', /messages\/delta/, () => ({ value: [], '@odata.deltaLink': link })]] as never);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.calls.map(c => c.url)).toEqual([link]);
    expect(r.calls[0].headers.Prefer).toMatch(/IdType="ImmutableId"/);
  });

  it.each([[['--name', 'x']], [['--reset']]])('--checkpoint with %j exits 2 before any request', async (extra) => {
    const r = await run(['mail', '--checkpoint', 'new', ...extra]);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('a value that is not a checkpoint exits 2', async () => {
    const r = await run(['mail', '--checkpoint', 'garbage']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});

describe('expired delta token (AH-4(3), AC 13)', () => {
  const expired = [['GET', MAIL, (c: RecordedCall) => c.url.includes('deltatoken=OLD')
    ? jsonResponse({ error: { code: 'syncStateNotFound', message: 'gone' } }, 410)
    : { value: [{ id: 'M1', changeKey: 'c1' }], '@odata.deltaLink': `${GRAPH}${MAIL}?$deltatoken=FRESH` }]] as never;

  it('without --strict: gap, then backfill, then checkpoint — all on stdout', async () => {
    const r = await run(['mail', '--checkpoint', mailCheckpoint()], expired);
    expect(r.exitCode, r.stderr).toBe(0);
    const out = lines(r.stdout);
    expect(out.map(l => l.type)).toEqual(['gap', 'backfill', 'checkpoint']);
    expect(out[0]).toEqual({ type: 'gap', resource: 'mail', reason: 'sync_state_expired', from: '2026-10-14T16:00:00.000Z', to: NOW });
    expect(out[1]).toMatchObject({ id: 'M1', version: 'c1' });
    expect(out[2].coverageStart).toBe(NOW);
    expect(decode(out[2].checkpoint).deltaLink).toBe(`${GRAPH}${MAIL}?$deltatoken=FRESH`);
  });

  it('with --strict: exit 1 delta_expired, nothing on stdout', async () => {
    setErrorFormat('json');
    const r = await run(['mail', '--strict', '--checkpoint', mailCheckpoint()], expired);
    expect(r.exitCode).toBe(1);
    expect(errorCode(r.stderr)).toBe('delta_expired');
    expect(r.stdout).toBe('');
  });
});

describe('watch mail --received-since and page size', () => {
  const SINCE = '2026-07-16T17:00:00.000Z'; // NOW - 90 days
  const paged = [['GET', MAIL, (c: RecordedCall) => c.url.includes('skiptoken=P2')
    ? { value: [], '@odata.deltaLink': `${GRAPH}${MAIL}?$deltatoken=D1` }
    : { value: [{ id: 'M1', changeKey: 'c1' }], '@odata.nextLink': `${GRAPH}${MAIL}?$skiptoken=P2` }]] as never;

  it('a baseline filters on receivedDateTime, asks for big pages on every page, and the checkpoint keeps the bound', async () => {
    const r = await run(['mail', '--checkpoint', 'new', '--received-since', '-90d'], paged);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(query(r.calls[0]).get('$filter')).toBe(`receivedDateTime ge ${SINCE}`);
    expect(r.calls.map(c => c.headers.Prefer)).toEqual(Array(2).fill('odata.maxpagesize=200, IdType="ImmutableId"'));
    const out = lines(r.stdout);
    expect(out[0]).toMatchObject({ type: 'checkpoint', receivedSince: SINCE, coverageStart: NOW });
    expect(decode(out[0].checkpoint).receivedSince).toBe(SINCE);
  });

  it('a resume keeps the bound without asking again; an expired cursor re-syncs with the same filter', async () => {
    const resumed = await run(['mail', '--checkpoint', mailCheckpoint({ receivedSince: SINCE })], changes);
    expect(resumed.calls.map(c => c.url)).toEqual([`${GRAPH}${MAIL}?$deltatoken=OLD`]);
    expect(lines(resumed.stdout).at(-1).receivedSince).toBe(SINCE);
    const expired = [['GET', MAIL, (c: RecordedCall) => c.url.includes('deltatoken=OLD')
      ? jsonResponse({ error: { code: 'syncStateNotFound', message: 'gone' } }, 410)
      : { value: [], '@odata.deltaLink': `${GRAPH}${MAIL}?$deltatoken=FRESH` }]] as never;
    const resynced = await run(['mail', '--checkpoint', mailCheckpoint({ receivedSince: SINCE })], expired);
    expect(resynced.exitCode, resynced.stderr).toBe(0);
    expect(query(resynced.calls[1]).get('$filter')).toBe(`receivedDateTime ge ${SINCE}`);
    expect(decode(lines(resynced.stdout).at(-1).checkpoint).receivedSince).toBe(SINCE);
  });

  it.each([
    ['a checkpoint with another bound', ['mail', '--received-since', '-30d', '--checkpoint', mailCheckpoint({ receivedSince: SINCE })], 'checkpoint_mismatch'],
    ['a whole-folder checkpoint', ['mail', '--received-since', '-90d', '--checkpoint', mailCheckpoint()], 'checkpoint_mismatch'],
    ['a time in the future', ['mail', '--received-since', '+1d', '--checkpoint', 'new'], 'usage'],
    ['watch events', ['events', '--received-since', '-90d', '--checkpoint', 'new'], 'usage'],
  ])('%s → exit 2 %s, zero requests', async (_label, argv, code) => {
    setErrorFormat('json');
    const r = await run(argv as string[]);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
    expect(errorCode(r.stderr)).toBe(code);
  });

  it('a named cursor keeps its bound; asking for another one needs --reset', async () => {
    expect((await run(['mail', '--name', 'n', '--received-since', '-90d'], baseline)).exitCode).toBe(0);
    const kept = await run(['mail', '--name', 'n'], changes);
    expect(kept.exitCode, kept.stderr).toBe(0);
    expect(JSON.parse(fs.readFileSync(path.join(dir, 'watch', 'n.json'), 'utf8')).receivedSince).toBe(SINCE);
    setErrorFormat('json');
    const other = await run(['mail', '--name', 'n', '--received-since', '-30d']);
    expect(other.exitCode).toBe(2);
    expect(other.calls).toHaveLength(0);
  });
});

describe('change lines (AH-4(2), AH-4(6), AC 14)', () => {
  it('mail lines carry version, lastModified, untrusted; a delete is a tombstone with removedReason', async () => {
    const r = await run(['mail', '--checkpoint', mailCheckpoint()], changes);
    const [updated, deleted] = lines(r.stdout);
    expect(updated).toMatchObject({ type: 'updated', id: 'M1', version: 'c2', lastModified: '2026-10-14T16:30:00Z', untrusted: true });
    expect(deleted).toMatchObject({ type: 'deleted', id: 'M3', version: null, lastModified: null, untrusted: true, removedReason: 'changed' });
  });

  it('a delete without a reason says deleted; the etag is the version when changeKey is absent', async () => {
    const r = await run(['events', '--checkpoint', 'new', '--include-existing'], [['GET', EVENTS, () => ({ value: [{ id: 'E1', '@odata.etag': 'W/"e"' }], '@odata.deltaLink': `${GRAPH}${EVENTS}?$deltatoken=E` })]] as never);
    expect(lines(r.stdout)[0]).toMatchObject({ type: 'backfill', resource: 'events', version: 'W/"e"' });
    expect(lines(r.stdout)[0].untrusted).toBeUndefined();
    const d = await run(['events', '--checkpoint', lines(r.stdout)[1].checkpoint], [['GET', EVENTS, () => ({ value: [{ id: 'E1', '@removed': {} }], '@odata.deltaLink': `${GRAPH}${EVENTS}?$deltatoken=E2` })]] as never);
    expect(lines(d.stdout)[0]).toMatchObject({ type: 'deleted', removedReason: 'deleted' });
  });

  it('watch mail --select id,subject asks for and keeps only those fields plus the dedupe fields', async () => {
    const r = await run(['mail', '--select', 'id,subject', '--checkpoint', 'new', '--include-existing'], [['GET', MAIL, () => ({ value: [
      { id: 'M1', subject: 'a', changeKey: 'c', lastModifiedDateTime: 't', bodyPreview: 'secret', '@odata.etag': 'W/"x"' },
    ], '@odata.deltaLink': `${GRAPH}${MAIL}?$deltatoken=D1` })]] as never);
    expect(query(r.calls[0]).get('$select')).toBe('id,subject,changeKey,lastModifiedDateTime');
    expect(lines(r.stdout)[0].item).toEqual({ id: 'M1', subject: 'a', changeKey: 'c', lastModifiedDateTime: 't' });
  });

  it('watch events --select projects locally (calendarView/delta may ignore $select †)', async () => {
    const r = await run(['events', '--select', 'subject', '--checkpoint', 'new', '--include-existing'], [['GET', EVENTS, () => ({ value: [
      { id: 'E1', subject: 's', body: { content: 'secret' }, changeKey: 'c' },
    ], '@odata.deltaLink': `${GRAPH}${EVENTS}?$deltatoken=E` })]] as never);
    expect(query(r.calls[0]).get('$select')).toBeNull();
    expect(lines(r.stdout)[0].item).toEqual({ subject: 's', id: 'E1', changeKey: 'c' });
  });

  it('--select on files or contacts exits 2', async () => {
    const r = await run(['files', '--select', 'id']);
    expect(r.exitCode).toBe(2);
  });
});

describe('events window roll (AH-4(5), AC 30)', () => {
  const eventsCheckpoint = (window: { start: string; end: string }) => blob({
    v: 1, resource: 'events', endpoint: EVENTS, deltaLink: `${GRAPH}${EVENTS}?$deltatoken=OLD`, syncedAt: '2026-10-14T16:00:00.000Z', coverageStart: window.start, window,
  });
  const routes = [['GET', EVENTS, () => ({ value: [{ id: 'E9', changeKey: 'k' }], '@odata.deltaLink': `${GRAPH}${EVENTS}?$deltatoken=NEW` })]] as never;

  it('less than half left: gap window_rolled, backfill from the new window, checkpoint with the new window', async () => {
    const old = { start: '2026-09-24T17:00:00.000Z', end: '2026-10-24T17:00:00.000Z' }; // 10 of 30 days left
    const r = await run(['events', '--checkpoint', eventsCheckpoint(old)], routes);
    expect(r.exitCode, r.stderr).toBe(0);
    const rolled = { start: NOW, end: '2026-11-13T17:00:00.000Z' };
    expect(r.calls).toHaveLength(1);
    expect(query(r.calls[0]).get('startDateTime')).toBe(rolled.start);
    expect(query(r.calls[0]).get('endDateTime')).toBe(rolled.end);
    const out = lines(r.stdout);
    expect(out.map(l => l.type)).toEqual(['gap', 'backfill', 'checkpoint']);
    expect(out[0]).toEqual({ type: 'gap', resource: 'events', reason: 'window_rolled', from: old.end, to: rolled.end, window: rolled });
    expect(out[2].window).toEqual(rolled);
    expect(decode(out[2].checkpoint).window).toEqual(rolled);
  });

  it('more than half left: resumes the deltaLink and keeps the window', async () => {
    const window = { start: '2026-10-10T17:00:00.000Z', end: '2026-11-09T17:00:00.000Z' };
    const r = await run(['events', '--checkpoint', eventsCheckpoint(window)], routes);
    expect(r.calls.map(c => c.url)).toEqual([`${GRAPH}${EVENTS}?$deltatoken=OLD`]);
    expect(lines(r.stdout).at(-1).window).toEqual(window);
  });

  it('a failed roll under --follow rolls again on the next pass instead of re-baselining silently', async () => {
    vi.spyOn(watchDeps, 'sleep').mockImplementation(async () => {});
    const old = { start: '2026-09-24T17:00:00.000Z', end: '2026-10-24T17:00:00.000Z' };
    let n = 0;
    env.MS365_CLI_MAX_RETRIES = '0';
    const r = await run(['events', '--checkpoint', eventsCheckpoint(old), '--follow', '--max-passes', '2'], [['GET', EVENTS, () => (++n === 1
      ? jsonResponse({ error: { code: 'ServiceUnavailable', message: 'down' } }, 503)
      : { value: [{ id: 'E9', changeKey: 'k' }], '@odata.deltaLink': `${GRAPH}${EVENTS}?$deltatoken=NEW` })]] as never);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.calls).toHaveLength(2);
    for (const c of r.calls) expect(query(c).get('startDateTime')).toBe(NOW);
    const out = lines(r.stdout);
    expect(out.filter(l => l.type === 'backfill').map(l => l.item?.id ?? l.id)).toEqual(['E9']);
  });
});
