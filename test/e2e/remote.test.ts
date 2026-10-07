import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { REMOTE_ENTRY, runCli, tempDir, type FakeRoute, type RunOptions } from './run-cli.js';
import { capsHeader, loginHome, startServeProxy, startServer, type Server } from './serve-harness.js';

const MESSAGES: FakeRoute = { match: '/me/messages\\?', method: 'GET', body: { value: [{ id: 'm1', subject: 'Hello, "world"' }] } };
const BYTES = Buffer.from(Array.from({ length: 70_000 }, (_, i) => (i * 7) % 256));
const LOCAL_ENV = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_TOKEN: 'fake' };

const remote = (args: string[], opts: RunOptions = {}) => runCli(args, { ...opts, entry: REMOTE_ENTRY });

/** A malicious or compromised `ms365 serve`: answers every execute with `files` and every GET with `abc`. */
async function hostileServe(files: unknown[]): Promise<{ url: string; gets: string[]; close(): Promise<void> }> {
  const gets: string[] = [];
  const server = http.createServer((req, res) => {
    req.resume().on('end', () => {
      if (req.method === 'GET') {
        gets.push(req.url ?? '');
        res.writeHead(200, { 'content-type': 'application/octet-stream' });
        res.end('abc');
        return;
      }
      res.writeHead(200, { 'content-type': 'application/json' });
      res.end(JSON.stringify({ ok: true, request_id: 'req_x', data: null, files }));
    });
  });
  const listening = Promise.withResolvers<void>();
  server.listen(0, '127.0.0.1', listening.resolve);
  await listening.promise;
  const { port } = server.address() as { port: number };
  return {
    url: `http://127.0.0.1:${port}`,
    gets,
    async close() {
      const closed = Promise.withResolvers<void>();
      server.close(() => closed.resolve());
      server.closeAllConnections();
      await closed.promise;
    },
  };
}

/** Peak resident set size of `pid` in kB (Linux `/proc`). */
function peakRssKb(pid: number): number {
  const status = fs.readFileSync(`/proc/${pid}/status`, 'utf8');
  return Number(/^VmHWM:\s+(\d+) kB$/m.exec(status)![1]);
}

function sha256File(file: string): string {
  const hash = createHash('sha256');
  const fd = fs.openSync(file, 'r');
  const buffer = Buffer.alloc(1024 * 1024);
  for (let n; (n = fs.readSync(fd, buffer)) > 0;) hash.update(buffer.subarray(0, n));
  fs.closeSync(fd);
  return hash.digest('hex');
}

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
        { match: '/drives/d1/items/i1/content', method: 'GET', bodyBase64: BYTES.toString('base64'), headers: { 'content-type': 'application/pdf' } },
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

    const down = await remote(['drive', 'download', 'd1', 'i1', '--out', './b.pdf'], { cwd, env });
    expect(down.code, down.stderr).toBe(0);
    expect(down.stdout).toBe('');
    expect(fs.readFileSync(path.join(cwd, 'b.pdf')).equals(BYTES)).toBe(true);

    // Same no-overwrite rule as the local command, checked before anything is sent.
    const before = server.requests().length;
    const again = await remote(['drive', 'download', 'd1', 'i1', '--out', './b.pdf'], { cwd, env });
    expect(again.code).toBe(1);
    expect(again.stderr).toContain('--force');
    expect(server.requests()).toHaveLength(before);
    const forced = await remote(['drive', 'download', 'd1', 'i1', '--out', './b.pdf', '--force'], { cwd, env });
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

describe('REMOTE-5 ms365-remote streams large files', () => {
  it('a 100 MB download round-trips byte-identically with --max-file-mb 200; serve peak RSS grows < 64 MB', async () => {
    const size = 100 * 1024 * 1024;
    const home = await loginHome();
    const server = await startServer({
      home,
      args: ['--max-file-mb', '200'],
      // `bodySize` streams generated bytes ('a'), so the fake Graph never holds the file either.
      routes: [{ match: '/drives/d1/items/i1/content', method: 'GET', bodySize: size, headers: { 'content-type': 'application/octet-stream' } }],
    });
    const proxy = await startServeProxy(server, capsHeader(['drive.read']));
    try {
      const linux = process.platform === 'linux';
      const before = linux ? peakRssKb(server.proc.pid!) : 0;
      const cwd = tempDir();
      const r = await remote(['drive', 'download', 'd1', 'i1', '--out', './big.bin'], { cwd, env: { MS365_REMOTE_URL: proxy.url }, timeoutMs: 110_000 });
      expect(r.code, r.stderr).toBe(0);
      const out = path.join(cwd, 'big.bin');
      expect(fs.statSync(out).size).toBe(size);
      const expected = createHash('sha256');
      const mb = Buffer.alloc(1024 * 1024, 0x61);
      for (let i = 0; i < 100; i++) expected.update(mb);
      expect(sha256File(out)).toBe(expected.digest('hex'));
      if (linux) expect(peakRssKb(server.proc.pid!) - before).toBeLessThan(64 * 1024);
    } finally {
      await proxy.close();
      await server.stop();
    }
  }, 120_000);

  it('an upload over the server\'s --max-file-mb fails with the 413 message: exit 1, nothing executed', async () => {
    const home = await loginHome();
    const server = await startServer({ home, args: ['--max-file-mb', '1'] });
    const proxy = await startServeProxy(server, capsHeader(['drive.write']));
    try {
      const cwd = tempDir();
      fs.writeFileSync(path.join(cwd, 'big.bin'), Buffer.alloc(2 * 1024 * 1024));
      const r = await remote(['drive', 'cp', './big.bin', '/Docs/'], { cwd, env: { MS365_REMOTE_URL: proxy.url } });
      expect(r.code).toBe(1);
      expect(r.stdout).toBe('');
      expect(r.stderr).toContain('Upload exceeds the 1 MB file limit (invalid_args)');
    } finally {
      await proxy.close();
      await server.stop();
    }
    expect(server.requests()).toEqual([]);
  });
});

describe('REMOTE-1d ms365-remote file names chosen by the server', () => {
  const file = (name: string, contentType = 'application/octet-stream') => ({ name, contentType, size: 3, url: '/v1/files/tok' });

  it('refuses empty and dot-leading names for named outputs: exit 1, nothing written', async () => {
    for (const name of ['.bashrc', 'x/.env', '..', '.', '', '/']) {
      const fake = await hostileServe([file(name)]);
      const cwd = tempDir();
      for (const args of [[], ['--force']]) {
        const r = await remote(['mail', 'download-attachment', 'm1', 'a1', ...args], { cwd, env: { MS365_REMOTE_URL: fake.url } });
        expect(r.code, `${JSON.stringify(name)} ${args}`).toBe(1);
        expect(r.stdout).toBe('');
        expect(r.stderr).toContain('--out');
      }
      expect(fs.readdirSync(cwd), JSON.stringify(name)).toEqual([]);
      expect(fake.gets).toEqual([]);
      await fake.close();
    }
  });

  it('names a profile photo profile-photo.<ext from content type> locally, like the local CLI', async () => {
    const fake = await hostileServe([file('.bashrc', 'image/png')]);
    const cwd = tempDir();
    const r = await remote(['user', 'download-photo'], { cwd, env: { MS365_REMOTE_URL: fake.url } });
    expect(r.code, r.stderr).toBe(0);
    expect(fs.readdirSync(cwd)).toEqual(['profile-photo.png']);
    expect(fs.readFileSync(path.join(cwd, 'profile-photo.png'), 'utf8')).toBe('abc');
    expect(fake.gets).toEqual(['/v1/files/tok']);
    await fake.close();
  });

  it('refuses a file URL outside /v1/files/ on the configured server: exit 1, nothing fetched', async () => {
    for (const url of ['http://evil.example/v1/files/tok', '//evil.example/v1/files/tok', '/v1/files/../execute', '/v1/files/']) {
      const fake = await hostileServe([{ name: 'a.bin', contentType: 'application/octet-stream', size: 3, url }]);
      const cwd = tempDir();
      const r = await remote(['drive', 'download', 'd1', 'i1', '--out', './a.bin'], { cwd, env: { MS365_REMOTE_URL: fake.url } });
      expect(r.code, url).toBe(1);
      expect(r.stderr, url).toContain('unexpected body');
      expect(fake.gets, url).toEqual([]);
      expect(fs.readdirSync(cwd)).toEqual([]);
      await fake.close();
    }
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
