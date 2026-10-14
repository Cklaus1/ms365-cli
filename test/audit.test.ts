import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { GraphClient } from '../src/graph-client.js';
import { recordAudit, setActor, setCommandPath, auditLogPath, AUDIT_MAX_BYTES } from '../src/audit.js';
import { registerAuditCommands } from '../src/commands/audit.js';
import { runCommand } from './helpers/cli.js';

const json = (data: unknown, status = 200, headers: Record<string, string> = {}) =>
  new Response(JSON.stringify(data), { status, headers: { 'content-type': 'application/json', ...headers } });

const readLines = () => fs.existsSync(auditLogPath())
  ? fs.readFileSync(auditLogPath(), 'utf8').trim().split('\n').filter(Boolean).map(l => JSON.parse(l))
  : [];

describe('audit log (§6 accountability, T37)', () => {
  let dir: string;
  let originalFetch: typeof globalThis.fetch;
  let originalArgv: string[];
  const auth = { getToken: vi.fn().mockResolvedValue('secret-token') };
  const envKeys = ['MS365_CLI_CONFIG_DIR', 'MS365_CLI_AUDIT', 'MS365_CLI_ACTOR'] as const;
  const envBackup: Record<string, string | undefined> = {};

  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-audit-'));
    for (const k of envKeys) envBackup[k] = process.env[k];
    process.env.MS365_CLI_CONFIG_DIR = path.join(dir, 'cfg');
    delete process.env.MS365_CLI_AUDIT;
    delete process.env.MS365_CLI_ACTOR;
    originalFetch = globalThis.fetch;
    originalArgv = process.argv;
    process.argv = ['node', 'ms365', 'mail', 'send', '--to', 'a@b.com', '--body', 'TOP-SECRET-BODY'];
    setActor(undefined);
    setCommandPath(undefined);
    vi.spyOn(console, 'error').mockImplementation(() => {});
  });
  afterEach(() => {
    globalThis.fetch = originalFetch;
    process.argv = originalArgv;
    for (const k of envKeys) { if (envBackup[k] === undefined) delete process.env[k]; else process.env[k] = envBackup[k]; }
    vi.restoreAllMocks();
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it('a POST writes exactly one line with no body, token or option values; file 0600, dir 0700', async () => {
    globalThis.fetch = vi.fn(async () => json({ id: 'RESULT-SECRET' }, 201, { 'request-id': 'req-1' })) as never;
    await new GraphClient(auth as never).request('/me/sendMail', {
      method: 'POST', body: JSON.stringify({ message: 'BODY-SECRET' }), queryParams: { $select: 'id' },
    });
    const raw = fs.readFileSync(auditLogPath(), 'utf8');
    expect(raw.trim().split('\n')).toHaveLength(1);
    for (const s of ['BODY-SECRET', 'RESULT-SECRET', 'secret-token', 'TOP-SECRET-BODY', 'a@b.com', '$select']) expect(raw).not.toContain(s);
    const [e] = readLines();
    expect(e).toMatchObject({ actor: 'cli', method: 'POST', endpoint: '/me/sendMail', status: 201, requestId: 'req-1', ok: true });
    expect(e.command).toBe('mail send --to --body');
    expect(e.account).toBeUndefined();
    expect(new Date(e.ts).toISOString()).toBe(e.ts);
    expect(fs.statSync(auditLogPath()).mode & 0o777).toBe(0o600);
    expect(fs.statSync(path.dirname(auditLogPath())).mode & 0o777).toBe(0o700);
  });

  it('a GET writes nothing; dry-run and read-only-blocked calls write nothing', async () => {
    globalThis.fetch = vi.fn(async () => json({})) as never;
    await new GraphClient(auth as never).request('/me');
    vi.spyOn(console, 'log').mockImplementation(() => {});
    await new GraphClient(auth as never, true).request('/me/sendMail', { method: 'POST' });
    await expect(new GraphClient(auth as never, false, 'v1.0', true).request('/me/sendMail', { method: 'POST' })).rejects.toThrow(/Read-only/);
    expect(fs.existsSync(auditLogPath())).toBe(false);
  });

  it('a 403 writes ok:false with status and requestId', async () => {
    globalThis.fetch = vi.fn(async () => json({ error: { code: 'Forbidden', message: 'no' } }, 403, { 'request-id': 'req-403' })) as never;
    await expect(new GraphClient(auth as never).request('/me/events/1', { method: 'DELETE' })).rejects.toThrow(/403/);
    expect(readLines()).toEqual([expect.objectContaining({ method: 'DELETE', endpoint: '/me/events/1', status: 403, requestId: 'req-403', ok: false })]);
  });

  it('a network error writes status 0, ok:false', async () => {
    globalThis.fetch = vi.fn(async () => { throw new TypeError('fetch failed'); }) as never;
    await expect(new GraphClient(auth as never).request('/me/sendMail', { method: 'POST' })).rejects.toThrow();
    expect(readLines()).toEqual([expect.objectContaining({ method: 'POST', status: 0, ok: false })]);
  });

  it('batch writes one line per non-GET sub-request, not one for $batch', async () => {
    globalThis.fetch = vi.fn(async (_u: string, init: RequestInit) => {
      const reqs = JSON.parse(init.body as string).requests as Array<{ id: string; url: string }>;
      return json({ responses: reqs.map(q => ({ id: q.id, status: q.url.startsWith('/bad') ? 404 : 200, body: { secret: 'ITEM-SECRET' }, headers: { 'request-id': `r-${q.id}` } })) });
    }) as never;
    await new GraphClient(auth as never).batch([
      { method: 'GET', url: '/me/messages/1' },
      { method: 'POST', url: '/me/messages/2/move', body: { destinationId: 'DEST-SECRET' } },
      { method: 'PATCH', url: '/bad?$select=id' },
    ]);
    const raw = fs.readFileSync(auditLogPath(), 'utf8');
    expect(raw).not.toContain('SECRET');
    expect(readLines().map(e => [e.method, e.endpoint, e.status, e.ok, e.requestId])).toEqual([
      ['POST', '/me/messages/2/move', 200, true, 'r-1'],
      ['PATCH', '/bad', 404, false, 'r-2'],
    ]);
  });

  it('uploadFile writes one line for the upload (simple and session)', async () => {
    const file = path.join(dir, 'f.bin');
    fs.writeFileSync(file, Buffer.alloc(10));
    globalThis.fetch = vi.fn(async (u: string) => String(u).includes('createUploadSession')
      ? json({ uploadUrl: 'https://up.sharepoint.com/x' })
      : json({ id: 'item' }, 201)) as never;
    const client = new GraphClient(auth as never);
    await client.uploadFile('/me/drive/root:/f.bin:/content', file, 'application/octet-stream');
    await client.uploadFile('/me/drive/root:/f.bin:/content', file, 'application/octet-stream', { maxSimpleSize: 1 });
    expect(readLines().map(e => [e.method, e.endpoint, e.status, e.ok])).toEqual([
      ['PUT', '/me/drive/root:/f.bin:/content', 201, true],
      ['PUT', '/me/drive/root:/f.bin:/content', 200, true],
    ]);
  });

  it('records the actor from setActor, then MS365_CLI_ACTOR, and a command path from setCommandPath', () => {
    process.env.MS365_CLI_ACTOR = 'env-bot';
    recordAudit({ method: 'POST', endpoint: '/a', status: 200, ok: true });
    setActor('flag-bot');
    setCommandPath('mail send');
    process.argv = ['node', 'ms365', '--actor', 'flag-bot', 'mail', 'send', '--body=SECRET2', 'positional-secret'];
    recordAudit({ method: 'POST', endpoint: '/a', status: 200, ok: true });
    const lines = readLines();
    expect(lines.map(e => e.actor)).toEqual(['env-bot', 'flag-bot']);
    expect(lines[1].command).toBe('mail send --actor --body');
    expect(fs.readFileSync(auditLogPath(), 'utf8')).not.toMatch(/SECRET2|positional-secret/);
  });

  it('MS365_CLI_AUDIT=0 disables logging', async () => {
    process.env.MS365_CLI_AUDIT = '0';
    globalThis.fetch = vi.fn(async () => json({}, 201)) as never;
    await new GraphClient(auth as never).request('/me/sendMail', { method: 'POST' });
    expect(fs.existsSync(auditLogPath())).toBe(false);
  });

  it('an audit failure never breaks the request', async () => {
    fs.writeFileSync(path.join(dir, 'cfg'), 'not a dir');
    globalThis.fetch = vi.fn(async () => json({ ok: 1 }, 201)) as never;
    const res = await new GraphClient(auth as never).request('/me/sendMail', { method: 'POST' });
    expect(res.status).toBe(201);
  });

  it('rotates to audit.log.1 when the log exceeds the size limit (keeps one)', () => {
    fs.mkdirSync(path.dirname(auditLogPath()), { recursive: true });
    fs.writeFileSync(auditLogPath() + '.1', 'older');
    fs.writeFileSync(auditLogPath(), 'x'.repeat(AUDIT_MAX_BYTES + 1));
    recordAudit({ method: 'POST', endpoint: '/a', status: 200, ok: true });
    expect(fs.statSync(auditLogPath() + '.1').size).toBe(AUDIT_MAX_BYTES + 1);
    expect(readLines()).toHaveLength(1);
    expect(fs.existsSync(auditLogPath() + '.2')).toBe(false);
  });

  it('R-13: a writer that saw a stale size never rotates a fresh log over the big .1', () => {
    const file = auditLogPath();
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, 'x'.repeat(AUDIT_MAX_BYTES + 1));
    // Simulate process A rotating and appending right after this process's unlocked stat saw > 5 MB.
    const realStat = fs.statSync;
    let raced = false;
    vi.spyOn(fs, 'statSync').mockImplementation(((p: fs.PathLike, o?: fs.StatSyncOptions) => {
      const st = realStat(p, o);
      if (!raced && p === file) {
        raced = true;
        fs.renameSync(file, `${file}.1`);
        fs.appendFileSync(file, '{"from":"A"}\n');
      }
      return st;
    }) as typeof fs.statSync);
    recordAudit({ method: 'POST', endpoint: '/b', status: 200, ok: true });
    expect(realStat(`${file}.1`).size).toBe(AUDIT_MAX_BYTES + 1);
    expect(fs.readFileSync(file, 'utf8').trim().split('\n')).toHaveLength(2);
  });

  it('R-13: rotation happens while holding audit.log.rotate.lock, and the lock is released', () => {
    const file = auditLogPath();
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, 'x'.repeat(AUDIT_MAX_BYTES + 1));
    const realRename = fs.renameSync;
    let lockedDuringRename = false;
    vi.spyOn(fs, 'renameSync').mockImplementation((from, to) => {
      lockedDuringRename = fs.existsSync(`${file}.rotate.lock`);
      realRename(from, to);
    });
    recordAudit({ method: 'POST', endpoint: '/a', status: 200, ok: true });
    expect(lockedDuringRename).toBe(true);
    expect(fs.existsSync(`${file}.rotate.lock`)).toBe(false);
    expect(readLines()).toHaveLength(1);
  });

  it('R-13: under the limit no lock file is taken (fast path)', () => {
    const open = vi.spyOn(fs, 'openSync');
    recordAudit({ method: 'POST', endpoint: '/a', status: 200, ok: true });
    recordAudit({ method: 'POST', endpoint: '/b', status: 200, ok: true });
    expect(open.mock.calls.filter(([p]) => String(p).endsWith('.lock'))).toHaveLength(0);
    expect(readLines()).toHaveLength(2);
  });

  it('audit tail -n 2 -o json prints the last two entries; audit path prints the path', async () => {
    for (const ep of ['/1', '/2', '/3']) recordAudit({ method: 'POST', endpoint: ep, status: 200, ok: true });
    const tail = await runCommand((p) => registerAuditCommands(p), ['audit', 'tail', '-n', '2', '-o', 'json']);
    expect(tail.calls).toHaveLength(0);
    expect(JSON.parse(tail.stdout).map((e: { endpoint: string }) => e.endpoint)).toEqual(['/2', '/3']);
    const nd = await runCommand((p) => registerAuditCommands(p), ['audit', 'tail', '-n', '1', '-o', 'ndjson']);
    expect(nd.stdout.trim().split('\n').map(l => JSON.parse(l).endpoint)).toEqual(['/3']);
    const p = await runCommand((p) => registerAuditCommands(p), ['audit', 'path']);
    expect(p.stdout.trim()).toBe(auditLogPath());
  });

  it('audit tail on a missing log prints an empty list', async () => {
    const tail = await runCommand((p) => registerAuditCommands(p), ['audit', 'tail', '-o', 'json']);
    expect(JSON.parse(tail.stdout)).toEqual([]);
  });
});
