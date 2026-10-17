import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { runCommand, jsonResponse, type RecordedCall } from '../helpers/cli.js';
import { registerWatchCommands, watchDeps } from '../../src/commands/watch.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const TZ = 'America/Los_Angeles';

let dir: string;
let env: Record<string, string>;
const cursorPath = (name: string) => path.join(dir, 'watch', `${name}.json`);
const readCursor = (name: string) => JSON.parse(fs.readFileSync(cursorPath(name), 'utf8'));
const writeCursor = (name: string, data: unknown) => {
  fs.mkdirSync(path.join(dir, 'watch'), { recursive: true });
  fs.writeFileSync(cursorPath(name), JSON.stringify(data));
};
const lines = (stdout: string) => stdout.trim().split('\n').filter(Boolean).map(l => JSON.parse(l));
const query = (call: RecordedCall) => new URLSearchParams(call.path.split('?')[1] ?? '');
const run = (argv: string[], routes: Parameters<typeof runCommand>[2] = {}) =>
  runCommand(registerWatchCommands, ['watch', ...argv], { env, ...routes });

const MAIL = '/me/mailFolders/inbox/messages/delta';
const MAIL_CURSOR = { resource: 'mail', endpoint: MAIL, deltaLink: `${GRAPH}${MAIL}?$deltatoken=OLD`, syncedAt: '2026-10-14T16:00:00.000Z' };

// Initial sync of two pages ending in a deltaLink.
const mailPages: Parameters<typeof runCommand>[2] = {
  routes: [['GET', MAIL, (call) => call.path.includes('skiptoken')
    ? { value: [{ id: 'M2', subject: 'b' }], '@odata.deltaLink': `${GRAPH}${MAIL}?$deltatoken=D1` }
    : { value: [{ id: 'M1', subject: 'a' }], '@odata.nextLink': `${GRAPH}${MAIL}?$skiptoken=P2` }]],
};

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-watch-'));
  env = { MS365_CLI_CONFIG_DIR: dir, MS365_CLI_TIMEZONE: TZ };
  vi.useFakeTimers({ toFake: ['Date'] });
  vi.setSystemTime(new Date('2026-10-14T17:00:00Z'));
});
afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
  fs.rmSync(dir, { recursive: true, force: true });
});

describe('watch: first run (baseline)', () => {
  it('follows nextLinks to the deltaLink, emits nothing, and stores the cursor', async () => {
    const r = await run(['mail'], mailPages);
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(2);
    expect(r.calls[0].path.split('?')[0]).toBe(MAIL);
    expect(query(r.calls[0]).get('$select')).toBe('id,subject,from,receivedDateTime,createdDateTime,lastModifiedDateTime,changeKey,isRead,hasAttachments,bodyPreview');
    expect(r.calls[1].url).toBe(`${GRAPH}${MAIL}?$skiptoken=P2`);
    expect(r.stdout).toBe('');
    expect(readCursor('mail')).toEqual({ resource: 'mail', endpoint: MAIL, deltaLink: `${GRAPH}${MAIL}?$deltatoken=D1`, syncedAt: '2026-10-14T17:00:00.000Z', coverageStart: '2026-10-14T17:00:00.000Z' });
    expect(fs.statSync(cursorPath('mail')).mode & 0o777).toBe(0o600);
  });

  it('--include-existing emits the initial items as backfill (AH-4(3))', async () => {
    const r = await run(['mail', '--include-existing'], mailPages);
    expect(r.exitCode).toBe(0);
    const line = (id: string, subject: string) => ({ type: 'backfill', resource: 'mail', id, version: null, lastModified: null, untrusted: true, item: { id, subject } });
    expect(lines(r.stdout)).toEqual([line('M1', 'a'), line('M2', 'b')]);
  });

  it('--folder with a well-known name targets it without a lookup and names the cursor after it', async () => {
    const r = await run(['mail', '--folder', 'Archive'], {
      routes: [['GET', '/me/mailFolders/archive/messages/delta', () => ({ value: [], '@odata.deltaLink': `${GRAPH}/x?$deltatoken=A` })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(1);
    expect(readCursor('mail-archive').endpoint).toBe('/me/mailFolders/archive/messages/delta');
  });

  it('--folder <name> resolves to the folder ID; the cursor is named from the ID (Q-13)', async () => {
    const r = await run(['mail', '--folder', 'Receipts'], {
      routes: [
        ['GET', /^\/me\/mailFolders\/AAMkRcpt=\/messages\/delta/, () => ({ value: [], '@odata.deltaLink': `${GRAPH}/x?$deltatoken=R` })],
        ['GET', '/me/mailFolders', () => ({ value: [{ id: 'AAMkRcpt=', displayName: 'Receipts', childFolderCount: 0 }] })],
      ],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls.map(c => c.path.split('?')[0])).toEqual(['/me/mailFolders', '/me/mailFolders/AAMkRcpt=/messages/delta']);
    expect(readCursor('mail-AAMkRcpt_').endpoint).toBe('/me/mailFolders/AAMkRcpt%3D/messages/delta');
  });

  it('--folder Inbox and --folder inbox share the default mail cursor (Q-13)', async () => {
    for (const folder of ['Inbox', 'inbox']) {
      const r = await run(['mail', '--folder', folder], mailPages);
      expect(r.exitCode).toBe(0);
      expect(r.calls[0].path.split('?')[0]).toBe(MAIL);
    }
    expect(fs.readdirSync(path.join(dir, 'watch'))).toEqual(['mail.json']);
  });

  it('--name sets the cursor file; an unsafe name is a usage error', async () => {
    const ok = await run(['contacts', '--name', 'crm.sync'], {
      routes: [['GET', '/me/contacts/delta', () => ({ value: [], '@odata.deltaLink': `${GRAPH}/me/contacts/delta?$deltatoken=C` })]],
    });
    expect(ok.exitCode).toBe(0);
    expect(readCursor('crm.sync').resource).toBe('contacts');

    const bad = await run(['contacts', '--name', '../evil']);
    expect(bad.exitCode).toBe(2);
    expect(bad.calls).toHaveLength(0);
  });

  it('rejects an unknown resource', async () => {
    const r = await run(['tasks']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});

describe('watch: resume', () => {
  it('resumes from the saved deltaLink and classifies created/updated/deleted', async () => {
    writeCursor('mail', MAIL_CURSOR);
    const r = await run(['mail'], {
      routes: [['GET', MAIL, () => ({
        value: [
          { id: 'N1', createdDateTime: '2026-10-14T16:30:00Z' },
          { id: 'U1', createdDateTime: '2026-10-01T09:00:00Z' },
          { id: 'R1', '@removed': { reason: 'deleted' } },
        ],
        '@odata.deltaLink': `${GRAPH}${MAIL}?$deltatoken=NEW`,
      })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(1);
    expect(r.calls[0].url).toBe(MAIL_CURSOR.deltaLink);
    expect(lines(r.stdout).map(l => [l.type, l.id])).toEqual([['created', 'N1'], ['updated', 'U1'], ['deleted', 'R1']]);
    expect(lines(r.stdout)[2].item).toEqual({ id: 'R1', '@removed': { reason: 'deleted' } });
    expect(readCursor('mail').deltaLink).toBe(`${GRAPH}${MAIL}?$deltatoken=NEW`);
  });

  it('--reset discards the cursor and baselines again', async () => {
    writeCursor('mail', MAIL_CURSOR);
    const r = await run(['mail', '--reset'], mailPages);
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].path.split('?')[0]).toBe(MAIL);
    expect(r.calls[0].url).not.toContain('deltatoken');
    expect(r.stdout).toBe('');
    expect(readCursor('mail').deltaLink).toBe(`${GRAPH}${MAIL}?$deltatoken=D1`);
  });

  it('refuses a cursor that belongs to a different resource', async () => {
    writeCursor('mail', { ...MAIL_CURSOR, resource: 'contacts', endpoint: '/me/contacts/delta' });
    const r = await run(['mail']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});

describe('watch: resources', () => {
  it('events: calendarView delta over --start..--end with the Prefer time zone', async () => {
    const r = await run(['events', '--start', '2026-10-15T00:00', '--end', '+7d'], {
      routes: [['GET', '/me/calendarView/delta', () => ({ value: [], '@odata.deltaLink': `${GRAPH}/me/calendarView/delta?$deltatoken=E` })]],
    });
    expect(r.exitCode).toBe(0);
    const q = query(r.calls[0]);
    expect(q.get('startDateTime')).toBe('2026-10-15T07:00:00.000Z');
    expect(q.get('endDateTime')).toBe('2026-10-21T17:00:00.000Z');
    expect(r.calls[0].headers.Prefer).toBe(`outlook.timezone="${TZ}", IdType="ImmutableId"`);
  });

  it('events: defaults to now..+30d', async () => {
    const r = await run(['events'], {
      routes: [['GET', '/me/calendarView/delta', () => ({ value: [], '@odata.deltaLink': `${GRAPH}/me/calendarView/delta?$deltatoken=E` })]],
    });
    const q = query(r.calls[0]);
    expect(q.get('startDateTime')).toBe('2026-10-14T17:00:00.000Z');
    expect(q.get('endDateTime')).toBe('2026-11-13T18:00:00.000Z'); // +30d wall-clock across the DST change
  });

  it('events: resumed requests keep the Prefer header', async () => {
    writeCursor('events', { resource: 'events', endpoint: '/me/calendarView/delta', deltaLink: `${GRAPH}/me/calendarView/delta?$deltatoken=E`, syncedAt: '2026-10-14T16:00:00.000Z' });
    const r = await run(['events'], {
      routes: [['GET', '/me/calendarView/delta', () => ({ value: [], '@odata.deltaLink': `${GRAPH}/me/calendarView/delta?$deltatoken=E2` })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].headers.Prefer).toBe(`outlook.timezone="${TZ}", IdType="ImmutableId"`);
  });

  it('files: baseline asks for token=latest; the deleted facet means deleted', async () => {
    const base = await run(['files'], {
      routes: [['GET', '/me/drive/root/delta', () => ({ value: [], '@odata.deltaLink': `${GRAPH}/me/drive/root/delta?token=T1` })]],
    });
    expect(base.exitCode).toBe(0);
    expect(query(base.calls[0]).get('token')).toBe('latest');

    const r = await run(['files'], {
      routes: [['GET', '/me/drive/root/delta', () => ({
        value: [{ id: 'F1', deleted: { state: 'deleted' } }, { id: 'F2', createdDateTime: '2026-10-14T17:00:00Z', name: 'a.txt' }],
        '@odata.deltaLink': `${GRAPH}/me/drive/root/delta?token=T2`,
      })]],
    });
    expect(r.calls[0].url).toBe(`${GRAPH}/me/drive/root/delta?token=T1`);
    expect(lines(r.stdout).map(l => [l.type, l.resource, l.id])).toEqual([['deleted', 'files', 'F1'], ['created', 'files', 'F2']]);
  });

  it('files --include-existing enumerates from the start (no token=latest)', async () => {
    const r = await run(['files', '--include-existing'], {
      routes: [['GET', '/me/drive/root/delta', () => ({ value: [{ id: 'F1' }], '@odata.deltaLink': `${GRAPH}/me/drive/root/delta?token=T1` })]],
    });
    expect(query(r.calls[0]).get('token')).toBeNull();
    expect(lines(r.stdout)).toHaveLength(1);
  });

  it('contacts: /me/contacts/delta', async () => {
    const r = await run(['contacts'], {
      routes: [['GET', '/me/contacts/delta', () => ({ value: [], '@odata.deltaLink': `${GRAPH}/me/contacts/delta?$deltatoken=C` })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].path).toBe('/me/contacts/delta');
  });
});

describe('watch: dry-run', () => {
  it('prints the request and leaves the cursor untouched', async () => {
    writeCursor('mail', MAIL_CURSOR);
    const r = await runCommand(registerWatchCommands, ['watch', 'mail'], { env, dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toContain(`[dry-run] GET ${MAIL_CURSOR.deltaLink}`);
    expect(readCursor('mail')).toEqual(MAIL_CURSOR);

    const fresh = await runCommand(registerWatchCommands, ['watch', 'contacts'], { env, dryRun: true });
    expect(fresh.exitCode).toBe(0);
    expect(fs.existsSync(cursorPath('contacts'))).toBe(false);
  });

  it('--reset previews a new baseline but never deletes the cursor (DR-18)', async () => {
    writeCursor('mail', MAIL_CURSOR);
    const r = await runCommand(registerWatchCommands, ['watch', 'mail', '--reset'], { env, dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain(`[dry-run] GET ${GRAPH}${MAIL}`);
    expect(r.stderr).not.toContain('deltatoken');
    expect(readCursor('mail')).toEqual(MAIL_CURSOR);
  });
});

describe('watch: expired delta token', () => {
  const expired = (code: string, status = 410): Parameters<typeof runCommand>[2] => ({
    routes: [['GET', MAIL, (call) => call.url.includes('deltatoken=OLD')
      ? jsonResponse({ error: { code, message: 'gone' } }, status)
      : { value: [{ id: 'M1' }], '@odata.deltaLink': `${GRAPH}${MAIL}?$deltatoken=FRESH` }]],
  });

  for (const [code, status] of [['syncStateNotFound', 410], ['resyncRequired', 410], ['SyncStateInvalid', 400]] as const) {
    it(`${code} (${status}) emits a gap, backfills existing items and baselines again (AH-4(3))`, async () => {
      writeCursor('mail', MAIL_CURSOR);
      const r = await run(['mail'], expired(code, status));
      expect(r.exitCode).toBe(0);
      expect(r.stderr).toMatch(/expired/i);
      expect(lines(r.stdout)).toEqual([
        { type: 'gap', resource: 'mail', reason: 'sync_state_expired', from: MAIL_CURSOR.syncedAt, to: '2026-10-14T17:00:00.000Z' },
        expect.objectContaining({ type: 'backfill', id: 'M1' }),
      ]);
      expect(r.calls).toHaveLength(2);
      expect(r.calls[1].url).not.toContain('deltatoken');
      expect(readCursor('mail').deltaLink).toBe(`${GRAPH}${MAIL}?$deltatoken=FRESH`);
    });
  }

  it('--strict exits 1 and keeps the old cursor', async () => {
    writeCursor('mail', MAIL_CURSOR);
    const r = await run(['mail', '--strict'], expired('syncStateNotFound'));
    expect(r.exitCode).toBe(1);
    expect(r.calls).toHaveLength(1);
    expect(readCursor('mail')).toEqual(MAIL_CURSOR);
  });

  it('other errors propagate', async () => {
    writeCursor('mail', MAIL_CURSOR);
    const r = await run(['mail'], { routes: [['GET', MAIL, () => jsonResponse({ error: { code: 'ErrorAccessDenied', message: 'no' } }, 403)]] });
    expect(r.exitCode).toBe(4);
  });
});

describe('watch --follow', () => {
  const counting = () => {
    let n = 0;
    return {
      routes: [['GET', '/me/contacts/delta', () => {
        n++;
        return { value: n === 1 ? [] : [{ id: `C${n}`, createdDateTime: '2026-10-14T18:00:00Z' }], '@odata.deltaLink': `${GRAPH}/me/contacts/delta?$deltatoken=T${n}` };
      }]],
    } as Parameters<typeof runCommand>[2];
  };

  it('sleeps --interval seconds between passes and saves the cursor each pass', async () => {
    const sleeps: number[] = [];
    vi.spyOn(watchDeps, 'sleep').mockImplementation(async (ms) => { sleeps.push(ms); });
    const r = await run(['contacts', '--follow', '--interval', '20', '--max-passes', '3'], counting());
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(3);
    expect(sleeps).toEqual([20_000, 20_000]);
    expect(lines(r.stdout).map(l => l.id)).toEqual(['C2', 'C3']);
    expect(r.calls[2].url).toContain('deltatoken=T2');
    expect(readCursor('contacts').deltaLink).toContain('deltatoken=T3');
  });

  it('defaults to a 60s interval and rejects one under 15s', async () => {
    const sleeps: number[] = [];
    vi.spyOn(watchDeps, 'sleep').mockImplementation(async (ms) => { sleeps.push(ms); });
    await run(['contacts', '--follow', '--max-passes', '2'], counting());
    expect(sleeps).toEqual([60_000]);

    const bad = await run(['contacts', '--follow', '--interval', '5']);
    expect(bad.exitCode).toBe(2);
    expect(bad.calls).toHaveLength(0);
  });

  it('stops cleanly on SIGINT with the cursor saved', async () => {
    vi.spyOn(watchDeps, 'sleep').mockImplementation(async () => { process.emit('SIGINT'); });
    const before = process.listenerCount('SIGINT');
    const r = await run(['contacts', '--follow'], counting());
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(1);
    expect(readCursor('contacts').deltaLink).toContain('deltatoken=T1');
    expect(process.listenerCount('SIGINT')).toBe(before);
  });

  const denied = () => jsonResponse({ error: { code: 'ErrorAccessDenied', message: 'no' } }, 403);
  /** Contacts feed whose pass n (1-based) fails when fail(n). */
  const flaky = (fail: (n: number) => boolean) => {
    let n = 0;
    return {
      routes: [['GET', '/me/contacts/delta', () => {
        n++;
        if (fail(n)) return denied();
        return { value: [], '@odata.deltaLink': `${GRAPH}/me/contacts/delta?$deltatoken=T${n}` };
      }]],
    } as Parameters<typeof runCommand>[2];
  };

  it('a failed pass warns, backs off interval x 2^n and continues (R-10)', async () => {
    const sleeps: number[] = [];
    vi.spyOn(watchDeps, 'sleep').mockImplementation(async (ms) => { sleeps.push(ms); });
    const r = await run(['contacts', '--follow', '--interval', '20', '--max-passes', '4'], flaky(n => n <= 2));
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(4);
    expect(sleeps).toEqual([40_000, 80_000, 20_000]);
    expect(r.stderr).toMatch(/Warning: watch pass failed .*retrying in 80s/i);
    expect(readCursor('contacts').deltaLink).toContain('deltatoken=T4');
  });

  it('backoff is capped at 15 minutes', async () => {
    const sleeps: number[] = [];
    vi.spyOn(watchDeps, 'sleep').mockImplementation(async (ms) => { sleeps.push(ms); });
    await run(['contacts', '--follow', '--interval', '600', '--max-passes', '3'], flaky(() => true));
    expect(sleeps).toEqual([900_000, 900_000]);
  });

  it('--max-failures N consecutive failures exits with the last error code', async () => {
    const sleeps: number[] = [];
    vi.spyOn(watchDeps, 'sleep').mockImplementation(async (ms) => { sleeps.push(ms); });
    const r = await run(['contacts', '--follow', '--interval', '20', '--max-failures', '2'], flaky(() => true));
    expect(r.exitCode).toBe(4);
    expect(r.calls).toHaveLength(2);
    expect(sleeps).toEqual([40_000]);
  });

  it('defaults to 10 consecutive failures; a success resets the count', async () => {
    vi.spyOn(watchDeps, 'sleep').mockImplementation(async () => {});
    const r = await run(['contacts', '--follow'], flaky(n => n !== 5));
    expect(r.exitCode).toBe(4);
    expect(r.calls).toHaveLength(15);
  });

  it('rejects a --max-failures that is not a positive integer', async () => {
    const r = await run(['contacts', '--follow', '--max-failures', '0']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('auth errors end --follow at once (exit 3)', async () => {
    const sleeps: number[] = [];
    vi.spyOn(watchDeps, 'sleep').mockImplementation(async (ms) => { sleeps.push(ms); });
    const r = await run(['contacts', '--follow'], {
      routes: [['GET', '/me/contacts/delta', () => jsonResponse({ error: { code: 'InvalidAuthenticationToken', message: 'expired' } }, 401)]],
    });
    expect(r.exitCode).toBe(3);
    expect(sleeps).toEqual([]);
  });

  it('a pass that fails after printing saves nothing, so its lines repeat next pass (R-11)', async () => {
    vi.spyOn(watchDeps, 'sleep').mockImplementation(async () => {});
    writeCursor('mail', MAIL_CURSOR);
    let page2 = 0;
    const r = await run(['mail', '--follow', '--max-passes', '2'], {
      routes: [['GET', MAIL, (call) => {
        if (call.url.includes('skiptoken')) {
          return ++page2 === 1 ? denied() : { value: [], '@odata.deltaLink': `${GRAPH}${MAIL}?$deltatoken=NEW` };
        }
        return { value: [{ id: 'N1', createdDateTime: '2026-10-14T16:30:00Z' }], '@odata.nextLink': `${GRAPH}${MAIL}?$skiptoken=P2` };
      }]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[2].url).toBe(MAIL_CURSOR.deltaLink);
    expect(lines(r.stdout).map(l => l.id)).toEqual(['N1', 'N1']);
    expect(readCursor('mail').deltaLink).toContain('deltatoken=NEW');
  });

  it('--help documents at-least-once delivery (R-11)', async () => {
    const r = await run(['--help']);
    expect(r.stdout).toMatch(/at-least-once/i);
    expect(r.stdout).toMatch(/dedupe on id/i);
  });

  it('SIGINT mid-pass stops after the current page, saves nothing and exits 130 (R-12, F22)', async () => {
    writeCursor('mail', MAIL_CURSOR);
    const r = await run(['mail', '--follow'], {
      routes: [['GET', MAIL, () => {
        process.emit('SIGINT');
        return { value: [{ id: 'N1', createdDateTime: '2026-10-14T16:30:00Z' }], '@odata.nextLink': `${GRAPH}${MAIL}?$skiptoken=P2` };
      }]],
    });
    expect(r.exitCode).toBe(130);
    expect(r.calls).toHaveLength(1);
    expect(lines(r.stdout).map(l => l.id)).toEqual(['N1']);
    expect(readCursor('mail')).toEqual(MAIL_CURSOR);
  });

  it('a second SIGINT exits 130 at once', async () => {
    const exit = vi.spyOn(process, 'exit').mockImplementation((() => {}) as never);
    const r = await run(['contacts', '--follow'], {
      routes: [['GET', '/me/contacts/delta', () => {
        process.emit('SIGINT');
        process.emit('SIGINT');
        return { value: [], '@odata.nextLink': `${GRAPH}/me/contacts/delta?$skiptoken=P2` };
      }]],
    });
    expect(exit).toHaveBeenCalledWith(130);
    expect(r.exitCode).toBe(130);
    expect(fs.existsSync(cursorPath('contacts'))).toBe(false);
  });
});

describe('watch under MCP (X-6)', () => {
  const empty = (prefix: string): Parameters<typeof runCommand>[2] => ({
    routes: [['GET', prefix, () => ({ value: [], '@odata.deltaLink': `${GRAPH}${prefix}?$deltatoken=X` })]],
  });

  it('MS365_CLI_MCP=1 prefixes the default cursor name with mcp-', async () => {
    env.MS365_CLI_MCP = '1';
    expect((await run(['contacts'], empty('/me/contacts/delta'))).exitCode).toBe(0);
    expect((await run(['mail', '--folder', 'archive'], empty('/me/mailFolders/archive/messages/delta'))).exitCode).toBe(0);
    expect(fs.readdirSync(path.join(dir, 'watch')).sort()).toEqual(['mcp-contacts.json', 'mcp-mail-archive.json']);
  });

  it('an explicit --name is used as given', async () => {
    env.MS365_CLI_MCP = '1';
    await run(['contacts', '--name', 'agent'], empty('/me/contacts/delta'));
    expect(fs.readdirSync(path.join(dir, 'watch'))).toEqual(['agent.json']);
  });
});

