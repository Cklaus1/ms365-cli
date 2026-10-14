import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { spawn, spawnSync } from 'child_process';
import { runCommand } from './helpers/cli.js';
import { registerMailCommands } from '../src/commands/mail.js';
import { registerCalendarCommands } from '../src/commands/calendar.js';
import { registerIdempotencyCommands } from '../src/commands/idempotency.js';
import { checkKey, recordKey, withIdempotency, fingerprint } from '../src/idempotency.js';
import { CliError } from '../src/errors.js';

let dir: string;
let ledger: string;
const readLedger = () => JSON.parse(fs.readFileSync(ledger, 'utf8')) as Array<Record<string, unknown>>;
const writeLedger = (entries: unknown[]) => {
  fs.mkdirSync(path.dirname(ledger), { recursive: true });
  fs.writeFileSync(ledger, JSON.stringify(entries));
};
const pending = (extra: Record<string, unknown>) =>
  ({ key: 'k1', command: 'mail send', fingerprint: 'fp', status: 'pending', ts: new Date().toISOString(), ...extra });
/** The pid of a process that has already exited (and been reaped). */
const deadPid = () => spawnSync(process.execPath, ['-e', '0']).pid!;
const codeOf = (fn: () => unknown) => {
  try { fn(); } catch (err) { return { code: (err as CliError).code, exitCode: (err as CliError).exitCode, fix: (err as CliError).fix }; }
  return undefined;
};
const sendArgs = (key: string, body = 'hi') =>
  ['mail', 'send', '--to', 'a@b.com', '--subject', 'S', '--body', body, '--idempotency-key', key];

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-idem-'));
  vi.stubEnv('MS365_CLI_CONFIG_DIR', path.join(dir, 'cfg'));
  ledger = path.join(dir, 'cfg', 'idempotency.json');
});
afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllEnvs();
  fs.rmSync(dir, { recursive: true, force: true });
});

describe('idempotency ledger (§6 Recover: retries, D23)', () => {
  it('fingerprint is canonical (key order does not matter)', () => {
    expect(fingerprint({ a: 1, b: { c: 2, d: 3 } })).toBe(fingerprint({ b: { d: 3, c: 2 }, a: 1 }));
    expect(fingerprint({ a: 1 })).not.toBe(fingerprint({ a: 2 }));
  });

  it('new → record → replay; ledger file is mode 0600', () => {
    expect(checkKey('k1', 'mail send', 'fp')).toEqual({ status: 'new' });
    recordKey('k1', 'mail send', 'fp', { sentAt: 'T' });
    expect(checkKey('k1', 'mail send', 'fp')).toMatchObject({ status: 'replay', result: { sentAt: 'T' } });
    expect(fs.statSync(ledger).mode & 0o777).toBe(0o600);
  });

  it('same key with a different fingerprint or command is a conflict (exit 2)', () => {
    recordKey('k1', 'mail send', 'fp', { sentAt: 'T' });
    for (const [cmd, fp] of [['mail send', 'other'], ['mail reply', 'fp']]) {
      try {
        checkKey('k1', cmd, fp);
        expect.unreachable();
      } catch (err) {
        expect(err).toBeInstanceOf(CliError);
        expect((err as CliError).code).toBe('idempotency_conflict');
        expect((err as CliError).exitCode).toBe(2);
      }
    }
  });

  it('entries expire after 24 h and are pruned on write', () => {
    vi.useFakeTimers({ toFake: ['Date'] });
    vi.setSystemTime(new Date('2026-10-14T00:00:00Z'));
    recordKey('old', 'mail send', 'fp', { sentAt: 'T' });
    vi.setSystemTime(new Date('2026-10-15T00:00:01Z'));
    expect(checkKey('old', 'mail send', 'other')).toEqual({ status: 'new' });
    recordKey('new', 'mail send', 'fp', { sentAt: 'T2' });
    expect(readLedger().map(e => e.key)).toEqual(['new']);
  });

  it('a pending entry makes a concurrent caller fail with idempotency_in_progress (exit 1)', async () => {
    let inner: unknown;
    await withIdempotency('k1', 'mail send', 'fp', async () => {
      expect(readLedger()[0]).toMatchObject({ key: 'k1', status: 'pending' });
      inner = await withIdempotency('k1', 'mail send', 'fp', async () => ({ sentAt: 'X' })).catch(e => e);
      return { sentAt: 'T' };
    });
    expect((inner as CliError).code).toBe('idempotency_in_progress');
    expect((inner as CliError).exitCode).toBe(1);
    expect(readLedger()[0]).toMatchObject({ status: 'done', result: { sentAt: 'T' } });
  });

  it('a failed call removes its pending entry so the key can be retried', async () => {
    await expect(withIdempotency('k1', 'mail send', 'fp', async () => { throw new Error('boom'); })).rejects.toThrow('boom');
    expect(checkKey('k1', 'mail send', 'fp')).toEqual({ status: 'new' });
  });
});

describe('stale and unknown claims (R-6, R-7, F17)', () => {
  it('a pending claim records its owner {pid, host, startedAt}', async () => {
    await withIdempotency('k1', 'mail send', 'fp', async () => {
      expect(readLedger()[0]).toMatchObject({ status: 'pending', pid: process.pid, host: os.hostname() });
      expect(Date.parse(readLedger()[0].startedAt as string)).toBeGreaterThan(Date.now() - 60_000);
      return { sentAt: 'T' };
    });
  });

  it('a live owner is in_progress; a dead owner on this host is outcome_unknown (exit 1, fix names release)', () => {
    writeLedger([pending({ pid: process.pid, host: os.hostname(), startedAt: new Date().toISOString() })]);
    expect(codeOf(() => checkKey('k1', 'mail send', 'fp'))).toMatchObject({ code: 'idempotency_in_progress' });
    writeLedger([pending({ pid: deadPid(), host: os.hostname(), startedAt: new Date().toISOString() })]);
    const e = codeOf(() => checkKey('k1', 'mail send', 'fp'));
    expect(e).toMatchObject({ code: 'idempotency_outcome_unknown', exitCode: 1 });
    expect(e!.fix).toContain('ms365 idempotency release k1');
    expect(readLedger()).toHaveLength(1); // never auto-released
  });

  it('a dead pid on another host is not judged; a claim older than 1 h is outcome_unknown', () => {
    writeLedger([pending({ pid: deadPid(), host: 'elsewhere', startedAt: new Date().toISOString() })]);
    expect(codeOf(() => checkKey('k1', 'mail send', 'fp'))).toMatchObject({ code: 'idempotency_in_progress' });
    const old = new Date(Date.now() - 61 * 60_000).toISOString();
    writeLedger([pending({ pid: 1, host: 'elsewhere', startedAt: old, ts: old })]);
    expect(codeOf(() => checkKey('k1', 'mail send', 'fp'))).toMatchObject({ code: 'idempotency_outcome_unknown' });
  });

  it('a failure tagged maybeProcessed keeps the claim as unknown; the retry fails outcome_unknown', async () => {
    const err = new CliError('timed out', { code: 'timeout', maybeProcessed: true });
    await expect(withIdempotency('k1', 'mail send', 'fp', async () => { throw err; })).rejects.toBe(err);
    expect(readLedger()[0]).toMatchObject({ key: 'k1', status: 'unknown' });
    expect(codeOf(() => checkKey('k1', 'mail send', 'fp'))).toMatchObject({ code: 'idempotency_outcome_unknown', exitCode: 1 });
  });

  it('mail send: a 5xx on sendMail leaves the key unknown and the retry sends nothing', async () => {
    const fail = await runCommand(registerMailCommands, sendArgs('k1'), {
      routes: [['POST', '/me/sendMail', () => new Response(JSON.stringify({ error: { code: 'X', message: 'boom' } }), { status: 503 })]],
    });
    expect(fail.exitCode).not.toBe(0);
    const retry = await runCommand(registerMailCommands, sendArgs('k1'));
    expect(retry.exitCode).toBe(1);
    expect(retry.calls).toHaveLength(0);
    expect(retry.stderr).toContain('idempotency release k1');
  });

  it('8 processes racing for one key: exactly one claims it, the rest are in_progress', async () => {
    const go = path.join(dir, 'go');
    const root = path.resolve(__dirname, '..');
    const outs: Promise<string>[] = [];
    const firstLines: Promise<void>[] = [];
    for (let i = 0; i < 8; i++) {
      const p = spawn(process.execPath, ['--import', 'tsx', path.join(root, 'test/fixtures/idem-claimer.ts'), 'race', go], { cwd: root });
      let out = '';
      firstLines.push(new Promise(resolve => p.stdout.on('data', (d) => { out += d; if (out.includes('\n')) resolve(); })));
      outs.push(new Promise((resolve, reject) => { p.on('exit', () => resolve(out.trim())); p.on('error', reject); }));
    }
    await Promise.all(firstLines);
    fs.writeFileSync(go, '');
    const results = (await Promise.all(outs)).sort();
    expect(results).toEqual(['idempotency_in_progress', ...Array(6).fill('idempotency_in_progress'), 'new'].sort());
    expect(readLedger()).toMatchObject([{ key: 'race', status: 'done' }]);
  }, 60_000);

  it('a claim held by a killed process reports outcome_unknown', async () => {
    const root = path.resolve(__dirname, '..');
    const p = spawn(process.execPath, ['--import', 'tsx', path.join(root, 'test/fixtures/idem-claimer.ts'), 'k1'], { cwd: root });
    await new Promise<void>(resolve => p.stdout.on('data', () => resolve()));
    const exited = new Promise(resolve => p.on('exit', resolve));
    p.kill('SIGKILL');
    await exited;
    expect(readLedger()[0]).toMatchObject({ key: 'k1', status: 'pending', pid: p.pid });
    expect(codeOf(() => checkKey('k1', 'mail send', 'fp'))).toMatchObject({ code: 'idempotency_outcome_unknown' });
  }, 30_000);
});

describe('ms365 idempotency list / release', () => {
  it('list prints the ledger entries as JSON', async () => {
    recordKey('k1', 'mail send', 'fp', { sentAt: 'T' });
    const r = await runCommand(registerIdempotencyCommands, ['idempotency', 'list']);
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(0);
    expect(JSON.parse(r.stdout)).toMatchObject([{ key: 'k1', command: 'mail send', status: 'done' }]);
  });

  it('release removes the entry and the next send proceeds; --dry-run leaves the ledger unchanged', async () => {
    writeLedger([pending({ pid: deadPid(), host: os.hostname(), startedAt: new Date().toISOString() })]);
    const before = fs.readFileSync(ledger, 'utf8');
    const dry = await runCommand(registerIdempotencyCommands, ['idempotency', 'release', 'k1'], { dryRun: true });
    expect(dry.exitCode).toBe(0);
    expect(dry.stderr).toContain('[dry-run]');
    expect(fs.readFileSync(ledger, 'utf8')).toBe(before);

    const r = await runCommand(registerIdempotencyCommands, ['idempotency', 'release', 'k1']);
    expect(r.exitCode).toBe(0);
    expect(readLedger()).toEqual([]);
    const send = await runCommand(registerMailCommands, sendArgs('k1'));
    expect(send.exitCode).toBe(0);
    expect(send.calls).toHaveLength(1);
  });

  it('release of an unknown key exits 5', async () => {
    const r = await runCommand(registerIdempotencyCommands, ['idempotency', 'release', 'nope']);
    expect(r.exitCode).toBe(5);
  });
});

describe('mail sends with --idempotency-key', () => {
  it('mail send: second run replays without calling Graph', async () => {
    const first = await runCommand(registerMailCommands, sendArgs('k1'));
    expect(first.exitCode).toBe(0);
    expect(first.calls).toHaveLength(1);

    const second = await runCommand(registerMailCommands, sendArgs('k1'));
    expect(second.exitCode).toBe(0);
    expect(second.calls).toHaveLength(0);
    expect(second.stderr).toMatch(/already sent at .+; not sent again/i);
    expect(JSON.parse(second.stdout)).toHaveProperty('sentAt');
  });

  it('mail send: same key, different body → conflict exit 2, nothing sent', async () => {
    await runCommand(registerMailCommands, sendArgs('k1'));
    const r = await runCommand(registerMailCommands, sendArgs('k1', 'different'));
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toMatch(/idempotency/i);
  });

  it('mail send: a failed send does not record the key', async () => {
    const fail = await runCommand(registerMailCommands, sendArgs('k1'), {
      routes: [['POST', '/me/sendMail', () => new Response(JSON.stringify({ error: { code: 'ErrorX', message: 'bad' } }), { status: 400 })]],
    });
    expect(fail.exitCode).not.toBe(0);
    const retry = await runCommand(registerMailCommands, sendArgs('k1'));
    expect(retry.calls).toHaveLength(1);
  });

  it.each([
    ['reply', ['mail', 'reply', 'M1', '--body', 'b']],
    ['forward', ['mail', 'forward', 'M1', '--to', 'a@b.com']],
    ['send-draft', ['mail', 'send-draft', 'M1']],
  ])('mail %s replays without calling Graph', async (_name, argv) => {
    const args = [...argv, '--idempotency-key', 'k-' + _name];
    expect((await runCommand(registerMailCommands, args)).calls).toHaveLength(1);
    const second = await runCommand(registerMailCommands, args);
    expect(second.exitCode).toBe(0);
    expect(second.calls).toHaveLength(0);
    expect(second.stderr).toMatch(/not sent again/);
  });

  it('dry-run neither reads nor writes the ledger', async () => {
    recordKey('k1', 'mail send', 'stale-fingerprint', { sentAt: 'T' });
    const before = fs.readFileSync(ledger, 'utf8');
    const r = await runCommand(registerMailCommands, sendArgs('k1'), { dryRun: true });
    expect(r.exitCode).toBe(0); // no conflict: the ledger was not consulted
    expect(r.calls).toHaveLength(0);
    expect(fs.readFileSync(ledger, 'utf8')).toBe(before);

    const fresh = await runCommand(registerMailCommands, sendArgs('k2'), { dryRun: true });
    expect(fresh.exitCode).toBe(0);
    expect(readLedger().map(e => e.key)).toEqual(['k1']);
  });
});

describe('cal create with --idempotency-key', () => {
  const args = ['cal', 'create', '--subject', 'S', '--start', '2026-10-20T09:00', '--end', '2026-10-20T09:30',
    '--timezone', 'UTC', '--idempotency-key', 'evt-1'];
  const routes: Parameters<typeof runCommand>[2] = {
    routes: [['POST', '/me/events', () => ({ id: 'E1', webLink: 'https://outlook/E1', subject: 'S' })]],
  };

  it('sends the key as transactionId and replays the stored id/webLink', async () => {
    const first = await runCommand(registerCalendarCommands, args, routes);
    expect(first.exitCode).toBe(0);
    expect((first.calls[0].body as Record<string, unknown>).transactionId).toBe('evt-1');
    expect(readLedger()[0]).toMatchObject({ key: 'evt-1', command: 'cal create', result: { id: 'E1', webLink: 'https://outlook/E1' } });

    const second = await runCommand(registerCalendarCommands, args, routes);
    expect(second.exitCode).toBe(0);
    expect(second.calls).toHaveLength(0);
    expect(JSON.parse(second.stdout)).toEqual({ id: 'E1', webLink: 'https://outlook/E1' });
  });

  it('dry-run includes transactionId but leaves the ledger untouched', async () => {
    const r = await runCommand(registerCalendarCommands, args, { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(fs.existsSync(ledger)).toBe(false);
    expect(r.stderr).toContain('"transactionId":"evt-1"');
  });
});
