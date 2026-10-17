#!/usr/bin/env node
// AHT-15: the shared contract fixture (Chief MSC-020 vendors it). Every case runs the CLI in host mode through the
// offline Graph double (test/fixtures/fetch-stub.mjs) with a frozen clock (MS365_TEST_NOW), a synthetic token and a
// clean environment, and keeps the exact bytes a host parses: stdout, or the error JSON (the last non-empty stderr
// line). `manifest.json` lists each file with its argv, exit code and sha256, plus the ms365 commit (with `dirty`
// when the inputs differed from it) and the dist/ sha256 it was generated from.
//
//   node scripts/contract.mjs                 regenerate test/fixtures/contract/ from dist/ (run `npm run build` first)
//   node scripts/contract.mjs --check         regenerate into a temp dir and fail on any byte difference
//   --entry src                               run src/index.ts through tsx instead of dist/ (the npm test check)
import { spawn, execFileSync } from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const FIXTURE_DIR = path.join(root, 'test', 'fixtures', 'contract');
const STUB = path.join(root, 'test', 'fixtures', 'fetch-stub.mjs');
const TSX = path.join(root, 'node_modules', 'tsx', 'dist', 'esm', 'index.mjs');

const NOW = '2026-10-17T12:00:00.000Z';
const TENANT = '11111111-2222-3333-4444-555555555555';
const OID = '22222222-3333-4444-5555-666666666666';
const UPN = 'ceo@contoso.com';
const SCOPES = ['User.Read', 'Mail.Read', 'Calendars.Read'];
const EXP = 4102444800; // 2100-01-01: valid at NOW on any machine

const b64 = o => Buffer.from(JSON.stringify(o)).toString('base64url');
const jwt = claims => `${b64({ alg: 'RS256', typ: 'JWT' })}.${b64(claims)}.sig`;
const LIVE = jwt({ aud: 'https://graph.microsoft.com', tid: TENANT, oid: OID, upn: UPN, scp: 'User.Read Mail.Read Calendars.Read', exp: EXP });
const EXPIRED = jwt({ aud: 'https://graph.microsoft.com', tid: TENANT, oid: OID, scp: 'Mail.Read', exp: 1 });
const OTHER_TENANT = jwt({ aud: 'https://graph.microsoft.com', tid: '99999999-2222-3333-4444-555555555555', oid: OID, scp: 'Mail.Read', exp: EXP });
const TASKS = jwt({ aud: 'https://graph.microsoft.com', tid: TENANT, oid: OID, scp: 'Tasks.ReadWrite', exp: EXP });

const G = 'https://graph.microsoft.com/v1.0';
const ME = { match: '/me\\?', body: { id: OID, userPrincipalName: UPN, mail: UPN } };
const MAIL_DELTA = `${G}/me/mailFolders/inbox/messages/delta`;
const msg = (id, extra = {}) => ({
  id, changeKey: `ck-${id}`, subject: `Subject ${id}`, createdDateTime: '2026-10-16T09:00:00Z', lastModifiedDateTime: '2026-10-16T09:00:00Z',
  receivedDateTime: '2026-10-16T09:00:00Z', from: { emailAddress: { address: 'cfo@contoso.com' } }, webLink: `https://outlook.office.com/mail/${id}`, ...extra,
});
const mailCheckpoint = b64({ v: 1, resource: 'mail', endpoint: '/me/mailFolders/inbox/messages/delta', deltaLink: `${MAIL_DELTA}?$deltatoken=m1`, syncedAt: '2026-10-17T11:00:00.000Z', coverageStart: '2026-10-01T00:00:00.000Z', folder: 'inbox' });
const eventsCheckpoint = b64({ v: 1, resource: 'events', endpoint: '/me/calendarView/delta', deltaLink: `${G}/me/calendarView/delta?$deltatoken=e1`, syncedAt: '2026-10-17T11:00:00.000Z', coverageStart: '2026-09-20T00:00:00.000Z', window: { start: '2026-09-20T00:00:00.000Z', end: '2026-10-20T00:00:00.000Z' } });
const driveCursor = Buffer.from(`${G}/me/drive/root/children?$skiptoken=x`).toString('base64url');
const page = (items, next) => ({ value: items, ...(next && { '@odata.nextLink': next }) });

const host = ['--profile', 'chief', '--error-format', 'json'];

/**
 * One case per fixture file. `capture`: `stdout` (the bytes a host reads from stdout) or `error` (the last non-empty
 * stderr line, the error-line contract). `token`: the MS365_CLI_TOKEN value (default the valid host token).
 */
export const CASES = [
  // The agent-host surface
  { file: 'describe.json', argv: [...host, 'describe', '--json'], exit: 0, capture: 'stdout' },
  { file: 'auth-status.json', argv: [...host, 'auth', 'status', '-o', 'json'], exit: 0, capture: 'stdout' },
  // Health (connector-health mapping)
  { file: 'health-healthy.json', argv: [...host, 'health'], exit: 0, capture: 'stdout',
    routes: [ME, { match: '/me/messages', body: page([]) }, { match: '/me/calendarView|/me/events|/me/calendar', body: page([]) }] },
  { file: 'health-degraded.json', argv: [...host, 'health'], exit: 0, capture: 'stdout',
    routes: [ME, { match: '/me/messages', status: 429, headers: { 'Retry-After': '7' }, body: { error: { code: 'TooManyRequests', message: 'throttled' } } },
      { match: '/me/calendarView|/me/events|/me/calendar', status: 403, body: { error: { code: 'ErrorAccessDenied', message: 'denied' } } }] },
  { file: 'health-stale.json', argv: [...host, 'health'], exit: 0, capture: 'stdout',
    routes: [ME, { match: '/me/messages', status: 503, body: { error: { code: 'ServiceUnavailable', message: 'unavailable' } } },
      { match: '/me/calendarView|/me/events|/me/calendar', body: page([]) }] },
  // One error envelope per exit-code family
  { file: 'error-1-graph_error.json', argv: [...host, 'user', 'me'], exit: 1, capture: 'error', env: { MS365_CLI_MAX_RETRIES: '0' },
    routes: [{ match: '/me', status: 503, body: { error: { code: 'ServiceUnavailable', message: 'unavailable' } } }] },
  { file: 'error-2-cursor_mismatch.json', argv: [...host, 'mail', 'list', '--cursor', driveCursor], exit: 2, capture: 'error' },
  { file: 'error-2-checkpoint_mismatch.json', argv: [...host, 'watch', 'mail', '--checkpoint', eventsCheckpoint], exit: 2, capture: 'error' },
  { file: 'error-3-token_missing.json', argv: [...host, 'user', 'me'], exit: 3, capture: 'error', token: '' },
  { file: 'error-3-token_expired.json', argv: [...host, 'user', 'me'], exit: 3, capture: 'error', token: EXPIRED },
  { file: 'error-3-token_identity_mismatch.json', argv: [...host, 'user', 'me'], exit: 3, capture: 'error', token: OTHER_TENANT },
  { file: 'error-3-token_rejected.json', argv: [...host, 'user', 'me'], exit: 3, capture: 'error',
    routes: [{ match: '/me', status: 401, body: { error: { code: 'InvalidAuthenticationToken', message: 'expired' } } }] },
  { file: 'error-4-scope_ceiling.json', argv: [...host, 'mail', 'send', '--to', 'cfo@contoso.com', '--subject', 's', '--body', 'b'], exit: 4, capture: 'error' },
  { file: 'error-5-not_found.json', argv: [...host, 'mail', 'get', 'M404'], exit: 5, capture: 'error',
    routes: [{ match: '/me/messages/M404', status: 404, body: { error: { code: 'ErrorItemNotFound', message: 'not found' } } }] },
  { file: 'error-6-throttled.json', argv: [...host, 'mail', 'list'], exit: 6, capture: 'error', env: { MS365_CLI_MAX_RETRIES: '0' },
    routes: [{ match: '/me/messages', status: 429, headers: { 'Retry-After': '7' }, body: { error: { code: 'TooManyRequests', message: 'throttled' } } }] },
  // No profile: a destructive command without a terminal (the host profile's ceiling would refuse it first)
  { file: 'error-7-confirmation_refused.json', argv: ['--error-format', 'json', 'todo', 'delete', 'L1', 'T1'], exit: 7, capture: 'error', token: TASKS },
  // List pages: complete / incompleteReason / nextCursor / NDJSON _meta / _ref
  { file: 'list-next-cursor.json', argv: [...host, '--refs', 'mail', 'list', '--top', '2', '-o', 'json'], exit: 0, capture: 'stdout',
    routes: [ME, { match: '/me/messages', body: page([msg('m1'), msg('m2')], `${G}/me/messages?$top=2&$skip=2`) }] },
  { file: 'list-next-cursor.ndjson', argv: [...host, '--refs', 'mail', 'list', '--top', '2', '-o', 'ndjson'], exit: 0, capture: 'stdout',
    routes: [ME, { match: '/me/messages', body: page([msg('m1'), msg('m2')], `${G}/me/messages?$top=2&$skip=2`) }] },
  { file: 'list-page-limit.json', argv: [...host, '--refs', 'mail', 'list', '--all', '--page-limit', '1', '-o', 'json'], exit: 0, capture: 'stdout',
    routes: [ME, { match: '/me/messages', body: page([msg('m1')], `${G}/me/messages?$skip=1`) }] },
  { file: 'list-search-cap.json', argv: [...host, 'mail', 'list', '--search', 'budget', '--top', '250', '--select', 'id', '-o', 'json'], exit: 0, capture: 'stdout',
    routes: [{ match: '/me/messages', body: page(Array.from({ length: 250 }, (_, i) => ({ id: `s${String(i + 1).padStart(3, '0')}` }))) }] },
  { file: 'list-complete.ndjson', argv: [...host, '--refs', 'calendar', 'view', '--start', '2026-10-19T00:00:00Z', '--end', '2026-10-20T00:00:00Z', '-o', 'ndjson'], exit: 0, capture: 'stdout',
    routes: [ME, { match: '/me/calendarView', body: page([{ id: 'e1', changeKey: 'ck-e1', subject: 'Board', start: { dateTime: '2026-10-19T09:00:00', timeZone: 'UTC' }, end: { dateTime: '2026-10-19T10:00:00', timeZone: 'UTC' } }]) }] },
  // Every watch line kind: checkpoint, backfill, created/updated/deleted (removedReason), gap (sync_state_expired, window_rolled)
  { file: 'watch-mail-baseline.ndjson', argv: [...host, '--refs', 'watch', 'mail', '--checkpoint', 'new', '--include-existing'], exit: 0, capture: 'stdout',
    routes: [ME, { match: '/messages/delta', body: { value: [msg('m1')], '@odata.deltaLink': `${MAIL_DELTA}?$deltatoken=m1` } }] },
  { file: 'watch-mail-changes.ndjson', argv: [...host, '--refs', 'watch', 'mail', '--checkpoint', mailCheckpoint], exit: 0, capture: 'stdout',
    routes: [ME, { match: 'deltatoken=m1', body: { value: [
      msg('m2', { createdDateTime: '2026-10-17T11:30:00Z', lastModifiedDateTime: '2026-10-17T11:30:00Z' }),
      msg('m1', { lastModifiedDateTime: '2026-10-17T11:40:00Z' }),
      { id: 'm3', '@removed': { reason: 'deleted' } },
      { id: 'm4', '@removed': { reason: 'changed' } },
    ], '@odata.deltaLink': `${MAIL_DELTA}?$deltatoken=m2` } }] },
  { file: 'watch-mail-expired.ndjson', argv: [...host, '--refs', 'watch', 'mail', '--checkpoint', mailCheckpoint], exit: 0, capture: 'stdout',
    routes: [ME, { match: 'deltatoken=m1', status: 410, body: { error: { code: 'SyncStateNotFound', message: 'expired' } } },
      { match: '/messages/delta', body: { value: [msg('m1')], '@odata.deltaLink': `${MAIL_DELTA}?$deltatoken=m3` } }] },
  { file: 'watch-events-window-rolled.ndjson', argv: [...host, '--refs', 'watch', 'events', '--checkpoint', eventsCheckpoint], exit: 0, capture: 'stdout',
    routes: [ME, { match: '/calendarView/delta', body: { value: [{ id: 'e1', changeKey: 'ck-e1', subject: 'Board', lastModifiedDateTime: '2026-10-10T09:00:00Z' }],
      '@odata.deltaLink': `${G}/me/calendarView/delta?$deltatoken=e2` } }] },
];

const sha256 = data => crypto.createHash('sha256').update(data).digest('hex');

/** The dist/ digest smoke prints and hosts pin: sha256 over `sha256sum` lines of every file, sorted by path. */
export function distSha256() {
  const dist = path.join(root, 'dist');
  if (!fs.existsSync(dist)) return null;
  const files = [];
  const walk = dir => {
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      const p = path.join(dir, e.name);
      if (e.isDirectory()) walk(p); else files.push(`./${path.relative(dist, p).split(path.sep).join('/')}`);
    }
  };
  walk(dist);
  files.sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));
  return sha256(files.map(f => `${sha256(fs.readFileSync(path.join(dist, f)))}  ${f}\n`).join(''));
}

/**
 * The commit the fixture was generated on, and whether the inputs that shape its bytes (src/, package files, the
 * generator and the fetch-stub) differed from it. Generated before committing, `dirty` is true and only distSha256
 * pins the bytes; regenerated on the committed tree, `commit` rebuilds the exact fixture.
 */
function gitProvenance() {
  const git = args => execFileSync('git', args, { cwd: root, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });
  try {
    const inputs = ['src', 'package.json', 'package-lock.json', 'tsconfig.json', 'tsup.config.ts', 'scripts/contract.mjs', 'test/fixtures/fetch-stub.mjs'];
    return { commit: git(['rev-parse', 'HEAD']).trim(), dirty: git(['status', '--porcelain', '--', ...inputs]).trim() !== '' };
  } catch {
    return { commit: null, dirty: null };
  }
}

/** Run one case in its own HOME/config dir with a clean environment; resolves the fixture bytes. */
function runCase(c, entry, work) {
  const home = fs.mkdtempSync(path.join(work, 'case-'));
  const cfg = path.join(home, 'cfg');
  fs.mkdirSync(path.join(cfg, 'profiles'), { recursive: true });
  fs.writeFileSync(path.join(cfg, 'profiles', 'chief.json'), JSON.stringify({ type: 'host', tenantId: TENANT, scopes: SCOPES }));
  const args = [...(entry === 'src' ? ['--import', TSX] : []), '--import', STUB, path.join(root, entry === 'src' ? 'src/index.ts' : 'dist/index.js'), ...c.argv];
  const env = {
    PATH: process.env.PATH ?? '', HOME: home, TZ: 'UTC', LANG: 'C', MS365_CLI_CONFIG_DIR: cfg, MS365_CLI_STORAGE: 'file', MS365_CLI_AUDIT: '0',
    MS365_CLI_TOKEN: c.token ?? LIVE, MS365_TEST_NOW: NOW, MS365_TEST_FETCH_ROUTES: JSON.stringify(c.routes ?? []), ...c.env,
  };
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, args, { cwd: home, env, stdio: ['ignore', 'pipe', 'pipe'] });
    const out = [], err = [];
    child.stdout.on('data', d => out.push(d));
    child.stderr.on('data', d => err.push(d));
    const timer = setTimeout(() => child.kill('SIGKILL'), 30_000);
    child.on('error', reject);
    child.on('close', code => {
      clearTimeout(timer);
      const stdout = Buffer.concat(out).toString('utf8'), stderr = Buffer.concat(err).toString('utf8');
      if (code !== c.exit) return reject(new Error(`${c.file}: exit ${code}, expected ${c.exit}\n${stderr}`));
      let bytes = c.capture === 'stdout' ? stdout : `${stderr.split('\n').filter(l => l.trim()).pop() ?? ''}\n`;
      if (c.capture === 'error') JSON.parse(bytes);
      // No host-specific bytes: the per-case temp dir never reaches a fixture.
      if (bytes.includes(home) || bytes.includes(os.tmpdir())) return reject(new Error(`${c.file}: output names a temp path`));
      resolve(bytes);
    });
  });
}

/** Generate every case into `outDir` (created), with its manifest. */
export async function generate(outDir, { entry = 'dist', concurrency = 6 } = {}) {
  if (entry === 'dist' && !fs.existsSync(path.join(root, 'dist', 'index.js'))) throw new Error('dist/ is missing: run npm run build first');
  fs.mkdirSync(outDir, { recursive: true });
  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-contract-'));
  const results = new Array(CASES.length);
  try {
    let next = 0;
    await Promise.all(Array.from({ length: Math.min(concurrency, CASES.length) }, async () => {
      while (next < CASES.length) {
        const i = next++;
        results[i] = await runCase(CASES[i], entry, work);
      }
    }));
  } finally {
    fs.rmSync(work, { recursive: true, force: true });
  }
  const files = {};
  CASES.forEach((c, i) => {
    fs.writeFileSync(path.join(outDir, c.file), results[i]);
    files[c.file] = { argv: c.argv, exitCode: c.exit, capture: c.capture, sha256: sha256(results[i]) };
  });
  const manifest = {
    v: 1,
    ms365: { ...gitProvenance(), distSha256: entry === 'dist' ? distSha256() : null },
    clock: NOW,
    profile: { name: 'chief', type: 'host', tenantId: TENANT, scopes: SCOPES },
    files,
  };
  fs.writeFileSync(path.join(outDir, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
  return manifest;
}

/**
 * Regenerate into a temp dir and compare with `fixtureDir`: every fixture file byte for byte, and the manifest
 * except its `ms365` block (the commit, dirty flag and dist digest it was generated from, which move with every change).
 * Returns the list of differences (empty: the contract is unchanged).
 */
export async function check({ entry = 'dist', fixtureDir = FIXTURE_DIR } = {}) {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-contract-check-'));
  try {
    const fresh = await generate(tmp, { entry });
    const diffs = [];
    const want = new Set([...Object.keys(fresh.files), 'manifest.json']);
    const have = fs.existsSync(fixtureDir) ? fs.readdirSync(fixtureDir) : [];
    for (const f of have) if (!want.has(f)) diffs.push(`${f}: not generated any more`);
    for (const f of Object.keys(fresh.files)) {
      const p = path.join(fixtureDir, f);
      if (!fs.existsSync(p)) diffs.push(`${f}: missing from the fixture`);
      else if (!fs.readFileSync(p).equals(fs.readFileSync(path.join(tmp, f)))) diffs.push(`${f}: bytes differ`);
    }
    let committed = null;
    try { committed = JSON.parse(fs.readFileSync(path.join(fixtureDir, 'manifest.json'), 'utf8')); } catch { diffs.push('manifest.json: missing or not JSON'); }
    if (committed) {
      const strip = m => JSON.stringify({ ...m, ms365: undefined });
      if (strip(committed) !== strip(fresh)) diffs.push('manifest.json: differs (outside its ms365 block)');
      if (entry === 'dist' && committed.ms365?.distSha256 !== fresh.ms365.distSha256) {
        process.stderr.write(`note: manifest distSha256 ${committed.ms365?.distSha256} is not the current dist/ (${fresh.ms365.distSha256}); regenerate before pinning\n`);
      }
    }
    return diffs;
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const argv = process.argv.slice(2);
  const entry = argv.includes('--entry') ? argv[argv.indexOf('--entry') + 1] : 'dist';
  if (entry !== 'dist' && entry !== 'src') { console.error('--entry must be dist or src'); process.exit(2); }
  try {
    if (argv.includes('--check')) {
      const diffs = await check({ entry });
      if (diffs.length) {
        console.error(`contract: ${diffs.length} difference(s) from test/fixtures/contract/ — a contract change must regenerate the fixture (npm run contract) in the same commit:\n  ${diffs.join('\n  ')}`);
        process.exit(1);
      }
      console.log(`contract: ok (${CASES.length} files unchanged)`);
    } else {
      fs.rmSync(FIXTURE_DIR, { recursive: true, force: true });
      const m = await generate(FIXTURE_DIR, { entry });
      console.log(`contract: wrote ${Object.keys(m.files).length} files + manifest.json (dist sha256 ${m.ms365.distSha256})`);
    }
  } catch (err) {
    console.error(`contract: ${err instanceof Error ? err.message : String(err)}`);
    process.exit(1);
  }
}
