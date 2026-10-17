import { describe, it, expect, vi, beforeAll, afterAll, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { spawnSync } from 'child_process';
import { Command } from 'commander';
import { runCommand } from '../helpers/cli.js';
import { registerAuditTrailCommands, importAuditCsv, IMPORT_FIELDS, IMPORT_CHUNK_BYTES } from '../../src/commands/audit-trail.js';
import { CsvParser, parseCsv } from '../../src/csv.js';
import { commandMeta } from '../../src/catalog.js';

/**
 * audit-trail import (AO-31, AC 14): offline, streaming analysis of a Purview audit export.
 *
 * test/fixtures/purview-export.csv is the portal variant (BOM, CRLF records, AuditData with `""` escapes and CRLF
 * whitespace inside the JSON). Its worked-example shape: Pat reads her own mailbox (omitted) and three delegated
 * mailboxes through two apps; SYSTEM makes one admin logon; one sign-in row is not a mailbox operation. The other two
 * header variants are re-emitted from the same rows in-test.
 */

const root = path.resolve(__dirname, '..', '..');
const FIXTURE = path.join(root, 'test', 'fixtures', 'purview-export.csv');
const APP1 = 'a1a1a1a1-1111-4111-8111-111111111111';
const APP2 = 'b2b2b2b2-2222-4222-8222-222222222222';
const PAT = 'pat.admin@contoso.com';
const SYSTEM = 'nt authority\\system (microsoft.exchange.servicehost)';
const REST = 'Client=REST;Client=RESTSystem;;';
/** Content in the fixture that must never reach any output stream (AC 7). */
const FORBIDDEN = ['SECRET-', 'Q3 layoffs', 'secret-recipient', 'ZZLEAK_', 'InternetMessageId', 'Subject'];

const row = (mailbox: string, events: number, items: number, first: string, last: string, appId: string | null,
  actor = PAT, logonType = 'delegated', client: string | null = REST) =>
  ({ actor, mailbox, logonType, events, items, first: `2026-10-15T${first}:00:00.000Z`, last: `2026-10-15T${last}:00:00.000Z`, client, appId });

// Hand-computed from the fixture rows.
const PAT_A = row('mbx-a@contoso.com', 2, 5, '09', '10', APP1);
const PAT_B = row('mbx-b@contoso.com', 1, 1, '11', '11', APP2);
const PAT_C = row('mbx-c@contoso.com', 1, 4, '12', '12', APP2);
const SYS_B = row('mbx-b@contoso.com', 1, 1, '13', '13', null, SYSTEM, 'admin', 'Client=Microsoft.Exchange.ServiceHost');

let tmp: string;
beforeAll(() => { tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-import-')); });
afterAll(() => { fs.rmSync(tmp, { recursive: true, force: true }); });
afterEach(() => { vi.restoreAllMocks(); });

const q = (s: string) => `"${s.replace(/"/g, '""')}"`;
function writeTmp(name: string, text: string): string {
  const p = path.join(tmp, name);
  fs.writeFileSync(p, text);
  return p;
}

/** The fixture's records as objects (header → value). */
function fixtureRows(): Record<string, string>[] {
  const [header, ...recs] = parseCsv(fs.readFileSync(FIXTURE, 'utf8'));
  return recs.map(r => Object.fromEntries(header.map((h, i) => [h, r[i] ?? ''])));
}

/** Older portal: CreationDate, UserIds, Operations, AuditData (LF, no BOM). */
function olderPortalVariant(): string {
  const lines = fixtureRows().map(r => [r.CreationDate, r.UserId, r.Operation, q(r.AuditData)].join(','));
  return ['CreationDate,UserIds,Operations,AuditData', ...lines].join('\n') + '\n';
}

/** Search-UnifiedAuditLog | Export-Csv: CreationDate (locale format), UserIds, RecordType, AuditData — every field quoted, no Operation column. */
function exportCsvVariant(): string {
  const lines = fixtureRows().map(r => {
    const d = new Date(r.CreationDate);
    const h = d.getUTCHours();
    const locale = `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()} ${h % 12 || 12}:00:00 ${h < 12 ? 'AM' : 'PM'}`;
    return [locale, r.UserId, r.RecordType, r.AuditData].map(q).join(',');
  });
  return ['"CreationDate","UserIds","RecordType","AuditData"', ...lines].join('\r\n') + '\r\n';
}

const noClient = () => { throw new Error('audit-trail import must not create a Graph client'); };
const register = (p: Command) => registerAuditTrailCommands(p, noClient as never);

async function run(argv: string[]) {
  const r = await runCommand(register, ['audit-trail', 'import', ...argv]);
  expect(r.calls, 'no Graph/network call').toEqual([]);
  return r;
}

async function rows(file: string, view: string[]) {
  const r = await run([file, ...view, '-o', 'json']);
  expect(r.exitCode, r.stderr).toBe(0);
  return JSON.parse(r.stdout).value as unknown[];
}

describe('audit-trail import: aggregates (AC 14 worked example)', () => {
  it('--by reproduces the per-(actor, mailbox) aggregates; own-mailbox and non-mailbox rows are left out', async () => {
    const r = await run([FIXTURE, '--by', 'Pat.Admin@Contoso.com', '-o', 'json']);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(JSON.parse(r.stdout).value).toEqual([PAT_A, PAT_B, PAT_C]);
    expect(r.stderr).toContain('7 rows');
    expect(r.stderr).toContain('1 own-mailbox events omitted');
    expect(r.stderr).toContain('1 rows of other operations ignored');
    expect(r.stderr).toContain('2026-10-15T07:00:00.000Z → 2026-10-15T13:00:00.000Z');
    expect(r.stderr).not.toContain('truncated');
  });

  it('--app, --mailbox and --cross-mailbox give the matching subsets', async () => {
    expect(await rows(FIXTURE, ['--app', APP1])).toEqual([PAT_A]);
    expect(await rows(FIXTURE, ['--app', APP2.toUpperCase()])).toEqual([PAT_B, PAT_C]);
    expect(await rows(FIXTURE, ['--mailbox', 'MBX-B@contoso.com'])).toEqual([SYS_B, PAT_B]);
    expect(await rows(FIXTURE, ['--cross-mailbox'])).toEqual([SYS_B, PAT_A, PAT_B, PAT_C]);
    expect(await rows(FIXTURE, ['--mailbox', 'pat.admin@contoso.com'])).toEqual([]); // the owner's own access is never shown
  });

  it('prints only the access fields, without the (online-only) target annotation', async () => {
    for (const r of await rows(FIXTURE, ['--cross-mailbox'])) expect(Object.keys(r as object).sort()).toEqual([...IMPORT_FIELDS].sort());
    expect(IMPORT_FIELDS).not.toContain('target');
  });

  it('the older-portal and Export-Csv header variants parse identically', async () => {
    const want = await rows(FIXTURE, ['--cross-mailbox']);
    expect(await rows(writeTmp('older.csv', olderPortalVariant()), ['--cross-mailbox'])).toEqual(want);
    expect(await rows(writeTmp('exportcsv.csv', exportCsvVariant()), ['--cross-mailbox'])).toEqual(want);
  });

  it('header names are matched case-insensitively and unknown columns are ignored', async () => {
    const lines = fixtureRows().map(r => [q('ZZLEAK_extra_ZZ'), r.CreationDate, r.UserId, r.Operation, q(r.AuditData)].join(','));
    const text = ['Extra,creationdate , USERIDS,operations,auditdata', ...lines].join('\n') + '\n';
    const r = await run([writeTmp('case.csv', text), '--cross-mailbox', '-o', 'json']);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(JSON.parse(r.stdout).value).toEqual([SYS_B, PAT_A, PAT_B, PAT_C]);
    expect(r.stdout).not.toContain('ZZLEAK_');
  });

  it('a chunk boundary inside a quoted AuditData (embedded CRLF and "") is handled at every split point', async () => {
    const text = fs.readFileSync(FIXTURE, 'utf8');
    const want = (await importAuditCsv([text], {})).agg.rows();
    const crlf = text.indexOf('\r\n', text.indexOf('{'));
    const esc = text.indexOf('""', text.indexOf('{'));
    expect(crlf).toBeGreaterThan(0);
    expect(esc).toBeGreaterThan(0);
    for (const at of [1, 2, crlf, crlf + 1, esc, esc + 1]) {
      const got = await importAuditCsv([text.slice(0, at), text.slice(at)], {});
      expect(got.agg.rows(), `split at ${at}`).toEqual(want);
    }
    const bytes = Buffer.from(text, 'utf8'); // the BOM split across byte chunks
    const got = await importAuditCsv([bytes.subarray(0, 2), bytes.subarray(2)], {});
    expect(got.agg.rows()).toEqual(want);
    expect(want.length).toBe(4);
  });

  it('needs a view flag (exit 2) and validates --app / --mailbox', async () => {
    expect((await run([FIXTURE])).exitCode).toBe(2);
    expect((await run([FIXTURE, '--app', 'not-a-guid'])).exitCode).toBe(2);
    expect((await run([FIXTURE, '--mailbox', 'a@b.com,c@d.com'])).exitCode).toBe(2);
    expect((await run([path.join(tmp, 'missing.csv'), '--cross-mailbox'])).exitCode).toBe(5);
  });

  it('a file without an AuditData column is refused without echoing the header', async () => {
    const r = await run([writeTmp('nohdr.csv', 'RecordId,SECRET-HEADER\r\n1,2\r\n'), '--cross-mailbox']);
    expect(r.exitCode).toBe(2);
    expect(r.stderr).toContain('no AuditData column');
    expect(r.stderr + r.stdout).not.toContain('SECRET-');
  });

  it('is catalogued as an offline read', () => {
    expect(commandMeta('audit-trail import')).toMatchObject({ scopes: [], offline: true, readOnly: true });
  });
});

describe('audit-trail import: privacy (AC 7 allow-list)', () => {
  const views = [['--cross-mailbox'], ['--by', PAT], ['--app', APP1], ['--mailbox', 'mbx-b@contoso.com']];
  for (const output of ['json', 'ndjson', 'table', 'csv', 'text', 'yaml']) {
    it(`-o ${output}: no subject, InternetMessageId, recipient or other non-allow-listed field in any view`, async () => {
      for (const view of views) {
        const r = await run([FIXTURE, ...view, '-o', output]);
        expect(r.exitCode, r.stderr).toBe(0);
        expect(r.stdout.length).toBeGreaterThan(0);
        for (const f of FORBIDDEN) {
          expect(r.stdout, `${output} ${view.join(' ')}: ${f}`).not.toContain(f);
          expect(r.stderr, `${output} ${view.join(' ')}: ${f}`).not.toContain(f);
        }
      }
    });
  }
});

describe('audit-trail import: malformed rows', () => {
  const header = 'RecordId,CreationDate,RecordType,Operation,UserId,AuditData\r\n';
  const good = fixtureRows()[1];
  const goodLine = [good.RecordId, good.CreationDate, good.RecordType, good.Operation, good.UserId, q(good.AuditData)].join(',');
  const id = '12345678-aaaa-4bbb-8ccc-000000000099';

  it('reports a non-JSON AuditData by row number and RecordId only (the subject never appears)', async () => {
    const broken = q('{"Item":{"Subject":"SECRET-MALFORMED-SUBJECT re: layoffs"');
    const file = writeTmp('bad.csv', `${header}${goodLine}\r\n${id},2026-10-15T09:00:00Z,ExchangeItem,MailItemsAccessed,x@contoso.com,${broken}\r\n`);
    const r = await run([file, '--cross-mailbox']);
    expect(r.exitCode).toBe(1);
    expect(r.stderr).toContain(`row 2 (RecordId ${id}): AuditData is not JSON`);
    for (const s of ['SECRET-', 'layoffs', 'Subject', 'x@contoso.com']) expect(r.stderr + r.stdout).not.toContain(s);
  });

  it('a RecordId that is not a GUID is not echoed either', async () => {
    const file = writeTmp('bad2.csv', `${header}SECRET-RECORD-ID,,,,,${q('["SECRET-ARRAY"]')}\r\n`);
    const r = await run([file, '--cross-mailbox']);
    expect(r.exitCode).toBe(1);
    expect(r.stderr).toContain('row 1: AuditData is not a JSON object');
    expect(r.stderr + r.stdout).not.toContain('SECRET-');
  });

  it('an empty AuditData is reported by row number', async () => {
    const r = await run([writeTmp('bad3.csv', `${header}${goodLine}\r\n${goodLine}\r\n${id},,,,,\r\n`), '--cross-mailbox']);
    expect(r.exitCode).toBe(1);
    expect(r.stderr).toContain(`row 3 (RecordId ${id}): AuditData is empty`);
  });
});

/** A synthetic export: `n` rows, each a delegated read by one of a few actors. */
function bigExport(file: string, n: number, minimal = false): string {
  const fd = fs.openSync(file, 'w');
  fs.writeSync(fd, 'RecordId,CreationDate,RecordType,Operation,UserId,AuditData\r\n');
  const ad = (i: number) => q(JSON.stringify({
    CreationTime: '2026-10-15T09:00:00', Operation: 'MailItemsAccessed', UserId: `u${i % 7}@contoso.com`, LogonType: 2,
    MailboxOwnerUPN: `m${i % 11}@contoso.com`, AppId: APP1, Item: { Subject: `SECRET-SUBJECT-${i}` },
    Folders: [{ FolderItems: [{ InternetMessageId: `<SECRET-${i}@x>` }] }],
  }));
  let buf = '';
  for (let i = 0; i < n; i++) {
    buf += minimal ? ',,,,,{}\r\n' : `,,,,,${ad(i)}\r\n`;
    if (buf.length > 1 << 20) { fs.writeSync(fd, buf); buf = ''; }
  }
  fs.writeSync(fd, buf);
  fs.closeSync(fd);
  return file;
}

describe('audit-trail import: truncation caps and streaming', () => {
  it('warns at exactly 50,000 rows (Audit Standard cap), streaming the file in chunks without reading it whole', async () => {
    const file = bigExport(path.join(tmp, 'standard.csv'), 50_000);
    const size = fs.statSync(file).size;
    expect(size).toBeGreaterThan(20 * IMPORT_CHUNK_BYTES);
    const chunks: number[] = [];
    const push = CsvParser.prototype.push;
    vi.spyOn(CsvParser.prototype, 'push').mockImplementation(function (this: CsvParser, chunk: string) {
      chunks.push(chunk.length);
      return push.call(this, chunk);
    });
    const readFileSync = vi.spyOn(fs, 'readFileSync');
    const readFile = vi.spyOn(fs.promises, 'readFile');

    const r = await run([file, '--cross-mailbox', '-o', 'json']);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.stderr).toMatch(/export probably truncated: 50,000 rows is the Audit \(Standard\) export cap/);
    const out = JSON.parse(r.stdout).value as { events: number; items: number }[];
    expect(out.reduce((s, x) => s + x.events, 0)).toBe(50_000);
    expect(out.reduce((s, x) => s + x.items, 0)).toBe(50_000);
    expect(r.stdout).not.toContain('SECRET-');

    expect(chunks.length).toBeGreaterThan(20);
    expect(Math.max(...chunks)).toBeLessThanOrEqual(IMPORT_CHUNK_BYTES);
    expect(readFileSync.mock.calls.filter(c => String(c[0]) === file)).toEqual([]);
    expect(readFile).not.toHaveBeenCalled();
  }, 60_000);

  it('does not warn at 49,999 rows', async () => {
    const r = await run([bigExport(path.join(tmp, 'under.csv'), 49_999, true), '--cross-mailbox']);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.stderr).not.toContain('truncated');
    expect(r.stderr).toContain('49999 rows');
  }, 60_000);

  it('warns at exactly 1,000,000 rows (Audit Premium cap)', async () => {
    const r = await run([bigExport(path.join(tmp, 'premium.csv'), 1_000_000, true), '--cross-mailbox']);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.stderr).toMatch(/export probably truncated: 1,000,000 rows is the Audit \(Premium\) export cap/);
  }, 120_000);
});

describe('audit-trail import: offline through the real entry point (AO-34, AC 11)', () => {
  // Any fetch writes a marker and throws; any @azure/msal-node load writes a marker. HOME and the config dir are
  // empty: no token cache, no selected account, no client ID.
  it('runs with no token cache and makes no network call', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-import-probe-'));
    try {
      const hooks = path.join(dir, 'hooks.mjs');
      fs.writeFileSync(hooks, 'export async function resolve(s, c, next) {\n'
        + "  if (s === '@azure/msal-node' || s.startsWith('@azure/msal-node/')) (await import('node:fs')).writeSync(2, '[probe] msal-node loaded\\n');\n"
        + '  return next(s, c);\n}\n');
      const probe = path.join(dir, 'probe.mjs');
      fs.writeFileSync(probe, "import { register } from 'node:module';\nimport fs from 'node:fs';\n"
        + `register(${JSON.stringify(new URL('file://' + hooks).href)});\n`
        + "globalThis.fetch = async () => { fs.writeSync(2, '[probe] fetch called\\n'); throw new Error('network is not allowed'); };\n");
      const home = path.join(dir, 'home');
      fs.mkdirSync(home);
      const tsxLoader = path.join(root, 'node_modules', 'tsx', 'dist', 'esm', 'index.mjs');
      const env: NodeJS.ProcessEnv = { ...process.env, HOME: home, MS365_CLI_CONFIG_DIR: path.join(home, 'cfg'), MS365_CLI_TOKEN: '', MS365_CLI_API_VERSION: '' };
      delete env.MS365_CLI_PREFLIGHT;
      delete env.MS365_CLI_CLIENT_ID;
      const r = spawnSync(process.execPath, ['--import', tsxLoader, '--import', probe, path.join(root, 'src/index.ts'),
        'audit-trail', 'import', FIXTURE, '--by', PAT, '-o', 'json'], { cwd: home, encoding: 'utf8', timeout: 30_000, input: '', env });
      expect(r.status, r.stderr).toBe(0);
      expect(r.stderr).not.toContain('[probe]');
      expect(JSON.parse(r.stdout).value).toEqual([PAT_A, PAT_B, PAT_C]);
      expect(fs.existsSync(path.join(home, 'cfg', 'audit.log')) ? fs.readFileSync(path.join(home, 'cfg', 'audit.log'), 'utf8') : '').not.toContain('SECRET-');
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  }, 30_000);
});
