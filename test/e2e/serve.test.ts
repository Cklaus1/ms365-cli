import { describe, it, expect, beforeAll } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { aadRoutes } from '../auth-fake-aad.js';
import { runCli, tempDir, type FakeRoute } from './run-cli.js';
import { APP_CAP, call, capsHeader, DEFAULT_CLIENT_ID, loginHome, maxAlive, rawCall, startServer } from './serve-harness.js';

const mode = (p: string) => fs.statSync(p).mode & 0o777;
const MESSAGES: FakeRoute = { match: '/me/messages', method: 'GET', body: { value: [{ id: 'm1', subject: 'Hello' }] } };
const childStarts = (events: { event: string; pid: number }[], serverPid?: number) =>
  events.filter(e => e.event === 'start' && e.pid !== serverPid).length;

describe('REMOTE-1a ms365 serve: transport', () => {
  it('refuses a non-loopback --host with exit 1', async () => {
    const r = await runCli(['serve', '--app-cap', APP_CAP, '--port', '0', '--host', '0.0.0.0']);
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('Refusing to listen on 0.0.0.0');
    expect(r.stdout).toBe('');
  });

  it('creates the default socket 0600 in the 0700 config dir and removes it on SIGTERM', async () => {
    const server = await startServer({ socket: null });
    const configDir = path.join(server.home, '.config', 'ms365-cli');
    expect(server.socket).toBe(path.join(configDir, 'serve.sock'));
    // Other Unix users cannot connect: only the owner may open a 0600 socket.
    expect(mode(server.socket!)).toBe(0o600);
    expect(mode(configDir)).toBe(0o700);
    expect((await call(server, { method: 'GET', path: '/v1/health' })).status).toBe(200);
    expect(await server.stop('SIGTERM')).toEqual({ code: 0, signal: null });
    expect(fs.existsSync(server.socket!)).toBe(false);
  });

  it('replaces a stale socket but refuses one that is still served', async () => {
    const dir = tempDir('ms365-sock-');
    const socket = path.join(dir, 's.sock');
    const crashed = await startServer({ socket });
    await crashed.stop('SIGKILL');
    expect(fs.existsSync(socket)).toBe(true);

    const server = await startServer({ socket });
    const second = await runCli(['serve', '--app-cap', APP_CAP, '--socket', socket, '--audit-log', path.join(dir, 'a.jsonl')]);
    expect(second.code).toBe(1);
    expect(second.stderr).toContain('Another server is listening');
    expect((await call(server, { method: 'GET', path: '/v1/health' })).status).toBe(200);
    await server.stop('SIGINT');
    expect(fs.existsSync(socket)).toBe(false);
  });

  it('--port binds 127.0.0.1 only and warns that local processes are trusted', async () => {
    const server = await startServer({ tcp: true });
    expect(server.url).toMatch(/^http:\/\/127\.0\.0\.1:\d+$/);
    expect(server.stderr()).toContain('Every local process can connect');
    expect((await call(server, { method: 'GET', path: '/v1/health' })).body).toMatchObject({ ok: true });
    await server.stop();
  });

  it('answers an unparseable request target with an error status and keeps serving', async () => {
    const server = await startServer();
    // llhttp accepts `//[`, but the WHATWG URL parser rejects it (unclosed IPv6 host).
    expect(await rawCall(server, 'GET //[ HTTP/1.1\r\nHost: x')).toBe(404);
    expect((await call(server, { method: 'GET', path: '/v1/health' })).status).toBe(200);
    expect(server.proc.exitCode).toBeNull();
    expect(await server.stop()).toEqual({ code: 0, signal: null });
  });

  it('waits for running requests before exiting on SIGTERM', async () => {
    const home = await loginHome();
    const server = await startServer({ home, routes: [{ ...MESSAGES, delayMs: 1500 }] });
    const pending = call(server, { caps: capsHeader(['mail.read']), body: { command: 'mail.list' } });
    await server.waitForChild();
    const stopped = server.stop('SIGTERM');
    const res = await pending;
    expect(res.status).toBe(200);
    expect(await stopped).toEqual({ code: 0, signal: null });
  });
});

describe('REMOTE-1b/1c authorization and execution', () => {
  let home: string;
  beforeAll(async () => {
    home = await loginHome();
  });

  it('denies requests without a usable capability grant: 403 and no child process', async () => {
    const server = await startServer({ home, routes: [MESSAGES] });
    const headers = [
      undefined,
      '{}',
      '=?utf-8?b?eyJleGFtcGxlLmNvbS9jYXAvbXMzNjUiOltdfQ==?=',
      'not json',
      JSON.stringify({ 'other.example/cap/x': [{ allow: ['mail.read'] }] }),
      JSON.stringify({ [APP_CAP]: [{ agent: 'chief', allow: [] }] }),
      JSON.stringify({ [APP_CAP]: [{ agent: 'chief', allow: 'mail.read' }] }),
    ];
    for (const caps of headers) {
      const res = await call(server, { caps, body: { command: 'mail.list', args: { dryRun: true } } });
      expect(res.status, String(caps)).toBe(403);
      expect(res.body).toMatchObject({ ok: false, error: { code: 'capability_missing' } });
      expect(res.body.request_id).toMatch(/^req_/);
    }
    await server.stop();
    expect(childStarts(server.procs(), server.proc.pid)).toBe(0);
  });

  it('denies a repeated Tailscale-App-Capabilities header even when the copies join into valid JSON', async () => {
    const server = await startServer({ home, routes: [MESSAGES] });
    // Node joins repeated custom headers with ", ": these halves would read as one grant for mail.read.
    const first = `{"${APP_CAP}":[{"agent":"chief","allow":["mail.read"]}`;
    const second = '{"allow":["mail.read"]}]}';
    const head = `POST /v1/execute HTTP/1.1\r\nHost: x\r\nTailscale-App-Capabilities: ${first}\r\nTailscale-App-Capabilities: ${second}`;
    expect(await rawCall(server, head, { command: 'mail.list' })).toBe(403);
    await server.stop();
    expect(childStarts(server.procs(), server.proc.pid)).toBe(0);
  });

  it('runs granted commands and refuses the rest per capability', async () => {
    const server = await startServer({ home, routes: [MESSAGES] });
    const caps = capsHeader(['mail.read']);

    const list = await call(server, { caps, body: { command: 'mail.list', args: { top: 5 } } });
    expect(list.status).toBe(200);
    // `data` is exactly what `ms365 mail list -o json` prints locally.
    expect(list.body).toMatchObject({ ok: true, data: { value: [{ id: 'm1', subject: 'Hello' }] } });

    const send = await call(server, { caps, body: { command: 'mail.send', args: { to: 'bob@example.com', subject: 's', body: 'b' } } });
    expect(send.status).toBe(403);
    expect(send.body.error).toEqual({ code: 'capability_missing', message: 'Missing capability: mail.send' });

    const login = await call(server, { caps: capsHeader(['mail.read', 'mail.send']), body: { command: 'auth.login' } });
    expect(login.status).toBe(403);
    expect(login.body.error?.code).toBe('capability_missing');

    const discovery = await call(server, { method: 'GET', path: '/v1/capabilities', caps });
    expect(discovery.body).toMatchObject({ ok: true, agent: 'chief', capabilities: ['mail.read'] });
    expect(discovery.body.commands).toContain('mail.list');
    expect(discovery.body.commands).not.toContain('mail.send');
    await server.stop();
    expect(server.requests().filter(q => q.url.includes('graph.microsoft.com'))).toHaveLength(1);
  });

  it('rejects unknown options with 400 and passes values as single literal argv elements', async () => {
    const server = await startServer({ home, routes: [MESSAGES] });
    const caps = capsHeader(['mail.read']);
    const bad = await call(server, { caps, body: { command: 'mail.list', args: { bogus: 1 } } });
    expect(bad.status).toBe(400);
    expect(bad.body.error).toMatchObject({ code: 'invalid_args', message: expect.stringContaining('unknown option: bogus') });
    const global = await call(server, { caps, body: { command: 'mail.list', args: { readOnly: true } } });
    expect(global.status).toBe(400);

    const hostile = '"; rm -rf ~ $(touch pwned) --dry-run';
    const res = await call(server, { caps, body: { command: 'mail.list', args: { filter: hostile } } });
    expect(res.status).toBe(200);
    await server.stop();
    const graph = server.requests().filter(q => q.url.includes('/me/messages'));
    expect(graph).toHaveLength(1);
    expect(new URL(graph[0].url).searchParams.get('$filter')).toBe(hostile);
  });

  it('mail.send with the mail.send capability → 200 ok, data null', async () => {
    const server = await startServer({ home, routes: [{ match: '/me/sendMail', method: 'POST', status: 202 }] });
    const res = await call(server, {
      caps: capsHeader(['mail.send']),
      body: { command: 'mail.send', args: { to: 'bob@example.com', subject: 'Hi', body: 'Text' } },
    });
    expect(res.status).toBe(200);
    expect(res.body).toMatchObject({ ok: true, data: null });
    await server.stop();
    const sent = server.requests().find(q => q.url.endsWith('/me/sendMail'));
    expect(JSON.parse(sent!.body!).message.toRecipients).toEqual([{ emailAddress: { address: 'bob@example.com' } }]);
  });

  it('reports command failures as 502 command_failed with the CLI error text', async () => {
    const server = await startServer({
      home,
      routes: [{ match: '/me/messages/missing', status: 404, body: { error: { code: 'ErrorItemNotFound', message: 'Item not found' } } }],
    });
    const res = await call(server, { caps: capsHeader(['mail.read']), body: { command: 'mail.get', positionals: ['missing'] } });
    expect(res.status).toBe(502);
    expect(res.body.error).toMatchObject({ code: 'command_failed', message: expect.stringContaining('Item not found') });
    await server.stop();
  });

  it('kills a command that exceeds --timeout: 504 timeout', async () => {
    const server = await startServer({ home, routes: [{ ...MESSAGES, delayMs: 10_000 }], args: ['--timeout', '1'] });
    const res = await call(server, { caps: capsHeader(['mail.read']), body: { command: 'mail.list' } });
    expect(res.status).toBe(504);
    expect(res.body.error?.code).toBe('timeout');
    await server.stop();
  });

  it('writes one audit line per request with agent and command but no argument values', async () => {
    const server = await startServer({ home, routes: [MESSAGES] });
    await call(server, {
      caps: capsHeader(['mail.read'], 'cfo'),
      headers: { 'x-forwarded-for': '100.64.0.7' },
      body: { command: 'mail.list', args: { filter: "subject eq 'secret-value-123'" } },
    });
    await call(server, { caps: capsHeader(['mail.read'], 'cfo'), body: { command: 'mail.send', args: { to: 'secret@example.com' } } });
    await server.stop();
    expect(mode(server.auditLog)).toBe(0o600);
    const text = fs.readFileSync(server.auditLog, 'utf8');
    expect(text).not.toContain('secret');
    const lines = text.trim().split('\n').map(l => JSON.parse(l));
    expect(lines).toHaveLength(2);
    expect(lines[0]).toMatchObject({ agent: 'cfo', src: '100.64.0.7', command: 'mail.list', args: ['filter'], status: 200 });
    expect(lines[0].request_id).toMatch(/^req_/);
    expect(typeof lines[0].duration_ms).toBe('number');
    expect(lines[1]).toMatchObject({ agent: 'cfo', command: 'mail.send', args: ['to'], status: 403, error: 'capability_missing' });
  });

  it('bounds caller-chosen audit fields: an unknown command and its arg keys are truncated', async () => {
    const server = await startServer({ home, routes: [MESSAGES] });
    const args = Object.fromEntries(Array.from({ length: 1000 }, (_, i) => [`${i}-${'k'.repeat(500)}`, 1]));
    const res = await call(server, { caps: capsHeader(['mail.read']), body: { command: 'x'.repeat(100_000), args } });
    expect(res.status).toBe(400);
    await server.stop();
    const text = fs.readFileSync(server.auditLog, 'utf8');
    expect(text.length).toBeLessThan(4096);
    const entry = JSON.parse(text) as { command: string; args: string[] };
    expect(entry.command.length).toBeLessThanOrEqual(128);
    expect(entry.command.startsWith('xxx')).toBe(true);
    expect(entry.args).toHaveLength(32);
    expect(entry.args.every(k => k.length <= 64)).toBe(true);
  });

  it('forwards --read-only from server start to every child', async () => {
    const server = await startServer({ home, globals: ['--read-only'], routes: [{ match: '/me/sendMail', method: 'POST', status: 202 }] });
    const res = await call(server, {
      caps: capsHeader(['mail.send']),
      body: { command: 'mail.send', args: { to: 'bob@example.com', subject: 'Hi', body: 'Text' } },
    });
    expect(res.status).toBe(502);
    expect(res.body.error?.message).toMatch(/read-only/i);
    await server.stop();
    expect(server.requests().filter(q => q.url.includes('sendMail'))).toHaveLength(0);
  });
});

describe('REMOTE-1c child configuration and load', () => {
  it('children use the client ID from the server .env and never MS365_CLI_TOKEN', async () => {
    const clientId = 'custom-client-id';
    const home = await loginHome({ clientId, expiresIn: 60, env: { MS365_CLI_CLIENT_ID: clientId } });
    const cwd = tempDir();
    fs.writeFileSync(path.join(cwd, '.env'), `MS365_CLI_CLIENT_ID=${clientId}\n`);
    const server = await startServer({
      home,
      cwd,
      env: { MS365_CLI_TOKEN: 'leaked-token' },
      routes: [...aadRoutes({ clientId, accessToken: 'at-2', refreshToken: 'rt-2', expiresIn: 3600 }), MESSAGES],
    });
    const res = await call(server, { caps: capsHeader(['mail.read']), body: { command: 'mail.list' } });
    expect(res.status).toBe(200);
    await server.stop();
    const token = server.requests().find(q => q.url.includes('/oauth2/v2.0/token'));
    expect(new URLSearchParams(token!.body!).get('client_id')).toBe(clientId);
    const graph = server.requests().find(q => q.url.includes('/me/messages'));
    expect(graph!.headers.authorization).toBe('Bearer at-2');
  });

  it('20 concurrent requests: all complete, at most 4 children alive, token cache still parses', async () => {
    const home = await loginHome({ expiresIn: 60 });
    const server = await startServer({
      home,
      routes: [
        ...aadRoutes({ clientId: DEFAULT_CLIENT_ID, accessToken: 'at-2', refreshToken: 'rt-2', expiresIn: 60, me: false }),
        { ...MESSAGES, delayMs: 300 },
      ],
    });
    const caps = capsHeader(['mail.read']);
    const results = await Promise.all(Array.from({ length: 20 }, () => call(server, { caps, body: { command: 'mail.list' } })));
    await server.stop();
    expect(results.map(r => r.status)).toEqual(Array(20).fill(200));
    expect(results.every(r => r.body.ok === true)).toBe(true);
    const events = server.procs();
    expect(childStarts(events, server.proc.pid)).toBe(20);
    const peak = maxAlive(events, server.proc.pid);
    expect(peak).toBeLessThanOrEqual(4);
    expect(peak).toBeGreaterThanOrEqual(2);
    const cache = fs.readFileSync(path.join(home, '.config', 'ms365-cli', 'token-cache.json'), 'utf8');
    expect(() => JSON.parse(cache)).not.toThrow();
  }, 90_000);
});

describe('REMOTE-1c files over the API', () => {
  it('accepts multipart uploads referenced by {"$file"} and returns output files as base64', async () => {
    const home = await loginHome();
    const bytes = Buffer.from(Array.from({ length: 4096 }, (_, i) => i % 256));
    const server = await startServer({
      home,
      routes: [
        { match: '/content\\?@microsoft\\.graph\\.conflictBehavior=rename', method: 'PUT', body: { id: 'new', name: 'a.bin' } },
        { match: '/drives/d1/items/i1/children/c1/content', method: 'GET', bodyBase64: bytes.toString('base64'), headers: { 'content-type': 'application/octet-stream' } },
      ],
    });
    const caps = capsHeader(['drive.read', 'drive.write']);

    const plainPath = await call(server, { caps, body: { command: 'drive.upload', positionals: ['d1', 'p1', '/etc/passwd'] } });
    expect(plainPath.status).toBe(400);
    expect(plainPath.body.error?.message).toContain('$file');

    const form = new FormData();
    form.append('request', JSON.stringify({ command: 'drive.upload', positionals: ['d1', 'p1', { $file: 'f' }] }));
    form.append('f', new Blob([bytes]), '../a.bin');
    const encoded = new Response(form);
    const upload = await call(server, {
      caps,
      headers: { 'content-type': encoded.headers.get('content-type')! },
      raw: Buffer.from(await encoded.arrayBuffer()),
    });
    expect(upload.status, JSON.stringify(upload.body)).toBe(200);
    expect(upload.body.data).toMatchObject({ id: 'new' });

    const download = await call(server, { caps, body: { command: 'drive.download', positionals: ['d1', 'i1', 'c1'], args: { out: '/etc/x.bin', force: true } } });
    expect(download.status).toBe(200);
    expect(download.body.files).toHaveLength(1);
    expect(download.body.files![0]).toMatchObject({ name: 'x.bin', size: bytes.length });
    expect(Buffer.from(download.body.files![0].base64, 'base64').equals(bytes)).toBe(true);
    await server.stop();

    const put = server.requests().find(q => q.method === 'PUT');
    expect(put!.url).toContain('/drives/d1/items/p1:/a.bin:/content');
    expect(put!.body).toBe(`base64:${bytes.toString('base64')}`);
    expect(fs.existsSync('/etc/x.bin')).toBe(false);
  });

  it('kills a download that grows past the 25 MB output limit: 502, nothing left on disk', async () => {
    const home = await loginHome();
    const tmp = tempDir('ms365-srv-tmp-');
    // 30 MB, then the body never ends: only a size bound (not the 5 s timeout) can end this request with 502.
    const huge = { match: '/drives/d1/items/i1/children/c1/content', method: 'GET', bodySize: 30 * 1024 * 1024, stall: true, headers: { 'content-type': 'application/octet-stream' } };
    const server = await startServer({ home, env: { TMPDIR: tmp }, args: ['--timeout', '5'], routes: [huge] });
    const res = await call(server, { caps: capsHeader(['drive.read']), body: { command: 'drive.download', positionals: ['d1', 'i1', 'c1'] } });
    expect(res.status, JSON.stringify(res.body)).toBe(502);
    expect(res.body.error).toEqual({ code: 'command_failed', message: 'Output file exceeds the 25 MB remote limit' });
    expect((await call(server, { method: 'GET', path: '/v1/health' })).status).toBe(200);
    expect(fs.readdirSync(tmp)).toEqual([]);
    await server.stop();
  }, 20_000);
});
