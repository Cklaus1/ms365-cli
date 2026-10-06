import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { REMOTE_ENTRY, runCli, tempDir, type FakeRoute, type RunOptions } from './run-cli.js';
import { capsHeader, loginHome, startServeProxy, startServer, type Server } from './serve-harness.js';

const MESSAGES: FakeRoute = { match: '/me/messages\\?', method: 'GET', body: { value: [{ id: 'm1', subject: 'Hello, "world"' }] } };
const BYTES = Buffer.from(Array.from({ length: 70_000 }, (_, i) => (i * 7) % 256));
const LOCAL_ENV = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_TOKEN: 'fake' };

const remote = (args: string[], opts: RunOptions = {}) => runCli(args, { ...opts, entry: REMOTE_ENTRY });

describe('REMOTE-1d ms365-remote over MS365_REMOTE_URL', () => {
  let server: Server;
  let proxy: { url: string; close(): Promise<void> };
  let env: Record<string, string>;

  beforeAll(async () => {
    const home = await loginHome();
    server = await startServer({
      home,
      routes: [
        MESSAGES,
        { match: '/me/drive/root:/Docs/a\\.pdf:/content\\?', method: 'PUT', body: { id: 'up1', name: 'a.pdf' } },
        { match: '/drives/d1/items/i1/children/c1/content', method: 'GET', bodyBase64: BYTES.toString('base64'), headers: { 'content-type': 'application/pdf' } },
        { match: '/me/drive/root:/Docs/notes\\.bin:/content', method: 'GET', bodyBase64: BYTES.toString('base64'), headers: { 'content-type': 'application/octet-stream' } },
        { match: '/me/messages/missing', status: 404, body: { error: { code: 'ErrorItemNotFound', message: 'Item not found' } } },
      ],
    });
    // tailscale serve injects the caller's grant; the shim itself never sends capabilities.
    proxy = await startServeProxy(server, capsHeader(['mail.read', 'drive.read', 'drive.write']));
    env = { MS365_REMOTE_URL: proxy.url };
  });

  afterAll(async () => {
    await proxy.close();
    await server.stop();
  });

  it('uploads a local file byte-identically and writes downloads to the local --out', async () => {
    const cwd = tempDir();
    fs.writeFileSync(path.join(cwd, 'a.pdf'), BYTES);
    const up = await remote(['drive', 'cp', './a.pdf', '/Docs/', '-o', 'json'], { cwd, env });
    expect(up.code, up.stderr).toBe(0);
    expect(JSON.parse(up.stdout)).toMatchObject({ id: 'up1' });
    const put = server.requests().find(q => q.method === 'PUT');
    expect(put!.body).toBe(`base64:${BYTES.toString('base64')}`);

    const down = await remote(['drive', 'download', 'd1', 'i1', 'c1', '--out', './b.pdf'], { cwd, env });
    expect(down.code, down.stderr).toBe(0);
    expect(down.stdout).toBe('');
    expect(fs.readFileSync(path.join(cwd, 'b.pdf')).equals(BYTES)).toBe(true);

    // Same no-overwrite rule as the local command, checked before anything is sent.
    const before = server.requests().length;
    const again = await remote(['drive', 'download', 'd1', 'i1', 'c1', '--out', './b.pdf'], { cwd, env });
    expect(again.code).toBe(1);
    expect(again.stderr).toContain('--force');
    expect(server.requests()).toHaveLength(before);
    const forced = await remote(['drive', 'download', 'd1', 'i1', 'c1', '--out', './b.pdf', '--force'], { cwd, env });
    expect(forced.code, forced.stderr).toBe(0);
  });

  it('streams binary output to stdout when --out is omitted', async () => {
    const r = await remote(['drive', 'cat', '/Docs/notes.bin'], { env });
    expect(r.code, r.stderr).toBe(0);
    expect(r.stdoutBytes.equals(BYTES)).toBe(true);
  });

  it('renders output locally, byte-identical to the local CLI (default and -o)', async () => {
    for (const args of [['mail', 'list'], ['mail', 'list', '-o', 'csv'], ['mail', 'list', '-o', 'yaml']]) {
      const local = await runCli(args, { env: LOCAL_ENV, routes: [MESSAGES] });
      expect(local.code, local.stderr).toBe(0);
      const r = await remote(args, { env });
      expect(r.code, r.stderr).toBe(0);
      expect(r.stdout, args.join(' ')).toBe(local.stdout);
    }
  });

  it('sends --dry-run as dryRun: nothing reaches Graph', async () => {
    const before = server.requests().length;
    const r = await remote(['--dry-run', 'mail', 'list', '-o', 'json'], { env });
    expect(r.code, r.stderr).toBe(0);
    expect(JSON.parse(r.stdout)).toMatchObject({ dryRun: true, method: 'GET' });
    expect(server.requests()).toHaveLength(before);
  });

  it('reports a server-side command failure on stderr with exit 1', async () => {
    const r = await remote(['mail', 'get', 'missing'], { env });
    expect(r.code).toBe(1);
    expect(r.stdout).toBe('');
    expect(r.stderr).toContain('Item not found');
  });

  it('reports a missing capability with exit 1', async () => {
    const r = await remote(['mail', 'send', '--to', 'a@example.com', '--subject', 's', '--body', 'b'], { env });
    expect(r.code).toBe(1);
    expect(r.stdout).toBe('');
    expect(r.stderr).toContain('Missing capability: mail.send (capability_missing)');
  });
});

describe('REMOTE-1d ms365-remote argv handling', () => {
  it('prints the same help and usage errors as ms365', async () => {
    const local = await runCli(['mail', 'list', '--help']);
    const r = await remote(['mail', 'list', '--help']);
    expect(r.code).toBe(0);
    expect(r.stdout.replace('ms365-remote', 'ms365')).toBe(local.stdout);

    const localBad = await runCli(['mail', 'list', '--bogus'], { env: LOCAL_ENV });
    const bad = await remote(['mail', 'list', '--bogus'], { env: { MS365_REMOTE_URL: 'http://127.0.0.1:9' } });
    expect(bad.code).toBe(1);
    expect(bad.stderr).toBe(localBad.stderr);
  });

  it('exits 1 with a clear message when no transport is configured', async () => {
    const r = await remote(['mail', 'list']);
    expect(r.code).toBe(1);
    expect(r.stdout).toBe('');
    expect(r.stderr).toMatch(/MS365_REMOTE_URL.*MS365_REMOTE_SSH/);
  });
});

describe('REMOTE-0 ms365-remote over MS365_REMOTE_SSH', () => {
  it('runs `tailscale ssh <host> -- ms365 <argv>` with literal args, stdio and exit code passed through', async () => {
    // Fake tailscale: like ssh, joins the words after `--` and runs them in the remote POSIX shell.
    // Fake ms365 on the "remote" PATH: prints each argument on its own line, then exits 3.
    const bin = tempDir('ms365-bin-');
    const log = path.join(bin, 'tailscale.log');
    fs.writeFileSync(path.join(bin, 'tailscale'), [
      '#!/bin/sh',
      `printf '%s\\n' "$@" > '${log}'`,
      'shift 3', // ssh <host> --
      'exec sh -c "$*"',
    ].join('\n'), { mode: 0o755 });
    fs.writeFileSync(path.join(bin, 'ms365'), '#!/bin/sh\nprintf \'[%s]\\n\' "$@"\nexit 3\n', { mode: 0o755 });

    const tricky = `it's "$HOME" $(rm -rf ~); x`;
    const r = await remote(['--dry-run', 'mail', 'list', '--filter', tricky], {
      env: { MS365_REMOTE_SSH: 'mailhost', PATH: `${bin}:${process.env.PATH}` },
    });
    expect(r.code).toBe(3);
    expect(r.stdout).toBe(`[--dry-run]\n[mail]\n[list]\n[--filter]\n[${tricky}]\n`);
    expect(fs.readFileSync(log, 'utf8').split('\n').slice(0, 4)).toEqual(['ssh', 'mailhost', '--', 'ms365']);
  });
});
