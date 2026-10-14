import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { Readable } from 'stream';
import { runCommand } from '../helpers/cli.js';
import { registerDriveCommands } from '../../src/commands/drive.js';

// Bytes that are not valid UTF-8 — any text decoding along the way would corrupt them.
const PDF = Buffer.concat([Buffer.from('%PDF-1.7\n'), Buffer.from([0x00, 0xff, 0xfe, 0x80, 0xc3, 0x28, 0x0a]), Buffer.from('%%EOF')]);
const pdfResponse = () => new Response(PDF, { status: 200, headers: { 'content-type': 'application/pdf', 'content-length': String(PDF.length) } });

let dir: string;
beforeEach(() => { dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-drive-')); });
afterEach(() => { fs.rmSync(dir, { recursive: true, force: true }); });

describe('drive download (P-2, C-3, C-19, D12)', () => {
  it('streams a PDF byte-exact to --out and reports on stderr only', async () => {
    const out = path.join(dir, 'report.pdf');
    const r = await runCommand(registerDriveCommands, ['drive', 'download', 'D', 'P', 'C', '--out', out], {
      routes: [['GET', '/drives/D/items/P/children/C/content', pdfResponse]],
    });
    expect(r.exitCode).toBe(0);
    expect(fs.readFileSync(out).equals(PDF)).toBe(true);
    expect(r.stdout).toBe('');
    expect(r.stderr).toContain(`Downloaded to ${out}`);
  });

  it('refuses to overwrite an existing file without --force (exit 1, nothing written)', async () => {
    const out = path.join(dir, 'exists.pdf');
    fs.writeFileSync(out, 'original');
    const r = await runCommand(registerDriveCommands, ['drive', 'download', 'D', 'P', 'C', '--out', out], {
      routes: [['GET', '/drives/', pdfResponse]],
    });
    expect(r.exitCode).toBe(1);
    expect(fs.readFileSync(out, 'utf8')).toBe('original');
    expect(r.calls).toHaveLength(0);
  });

  it('overwrites with --force', async () => {
    const out = path.join(dir, 'exists.pdf');
    fs.writeFileSync(out, 'original');
    const r = await runCommand(registerDriveCommands, ['drive', 'download', 'D', 'P', 'C', '--out', out, '--force'], {
      routes: [['GET', '/drives/', pdfResponse]],
    });
    expect(r.exitCode).toBe(0);
    expect(fs.readFileSync(out).equals(PDF)).toBe(true);
  });

  it('works under --dry-run without writing a file', async () => {
    const out = path.join(dir, 'dry.pdf');
    const r = await runCommand(registerDriveCommands, ['drive', 'download', 'D', 'P', 'C', '--out', out], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stdout).toContain('"dryRun": true');
    expect(fs.existsSync(out)).toBe(false);
  });
});

describe('drive download-shared (M-20)', () => {
  it('encodes a sharing URL as u! + unpadded base64url', async () => {
    const link = 'https://contoso.sharepoint.com/:b:/g/EaBc?e=1';
    const expected = 'u!' + Buffer.from(link).toString('base64').replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_');
    const out = path.join(dir, 'shared.pdf');
    const r = await runCommand(registerDriveCommands, ['drive', 'download-shared', link, '--out', out], {
      routes: [['GET', '/shares/', pdfResponse]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].url).toBe(`https://graph.microsoft.com/v1.0/shares/${expected}/driveItem/content`);
    expect(expected).not.toMatch(/[=+/]/);
    expect(fs.readFileSync(out).equals(PDF)).toBe(true);
  });

  it('passes a share ID through unchanged', async () => {
    const out = path.join(dir, 's.bin');
    const r = await runCommand(registerDriveCommands, ['drive', 'download-shared', 's!abc', '--out', out], {
      routes: [['GET', '/shares/', pdfResponse]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].url).toBe('https://graph.microsoft.com/v1.0/shares/s!abc/driveItem/content');
  });
});

describe('drive cat (P-2, C-3, C-10)', () => {
  it('writes the raw bytes to stdout', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'cat', '/Docs/a.pdf'], {
      routes: [['GET', '/me/drive/root:/Docs/a.pdf:/content', pdfResponse]],
    });
    expect(r.exitCode).toBe(0);
    expect(Buffer.from(r.stdout, 'latin1').equals(PDF)).toBe(true);
  });

  it('writes a JSON file verbatim (no parsing, @odata keys kept)', async () => {
    const text = '{"@odata.type":"x","a":1}\n';
    const r = await runCommand(registerDriveCommands, ['drive', 'cat', 'n.json'], {
      routes: [['GET', '/me/drive/root:/n.json:/content', () => new Response(text, { headers: { 'content-type': 'application/json' } })]],
    });
    expect(r.stdout).toBe(text);
  });

  it('saves to --out and refuses to overwrite', async () => {
    const out = path.join(dir, 'a.pdf');
    const ok = await runCommand(registerDriveCommands, ['drive', 'cat', 'a.pdf', '--out', out], { routes: [['GET', '/me/', pdfResponse]] });
    expect(ok.exitCode).toBe(0);
    expect(fs.readFileSync(out).equals(PDF)).toBe(true);
    expect(ok.stdout).toBe('');
    const again = await runCommand(registerDriveCommands, ['drive', 'cat', 'a.pdf', '--out', out], { routes: [['GET', '/me/', pdfResponse]] });
    expect(again.exitCode).toBe(1);
  });

  it("encodes '#', '%', '?' and spaces in path segments", async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'cat', '/R#D/50% off?/Q3 Report.pdf'], {
      routes: [['GET', '/me/', pdfResponse]],
    });
    expect(r.calls[0].url).toBe('https://graph.microsoft.com/v1.0/me/drive/root:/R%23D/50%25%20off%3F/Q3%20Report.pdf:/content');
  });
});

describe('path / name / search encoding (C-10)', () => {
  it('ls encodes path segments', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'ls', '/A#1/b?c'], { routes: [['GET', '/me/', () => ({ value: [] })]] });
    expect(r.calls[0].url).toBe('https://graph.microsoft.com/v1.0/me/drive/root:/A%231/b%3Fc:/children');
  });

  it('upload encodes the file name', async () => {
    const file = path.join(dir, 'Q#1 50%.txt');
    fs.writeFileSync(file, 'hi');
    const r = await runCommand(registerDriveCommands, ['drive', 'upload', 'D', 'P', file], { routes: [['PUT', '/drives/', () => ({ id: 'x' })]] });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].url).toBe('https://graph.microsoft.com/v1.0/drives/D/items/P:/Q%231%2050%25.txt:/content');
  });

  it('cp encodes the remote path', async () => {
    const file = path.join(dir, 'x.txt');
    fs.writeFileSync(file, 'hi');
    const r = await runCommand(registerDriveCommands, ['drive', 'cp', file, '/Docs/Q#1.txt'], { routes: [['PUT', '/me/', () => ({ id: 'x' })]] });
    expect(r.calls[0].url).toBe('https://graph.microsoft.com/v1.0/me/drive/root:/Docs/Q%231.txt:/content');
  });

  it("search escapes ' and encodes the term", async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'search', "O'Brien #1?"], { routes: [['GET', '/me/', () => ({ value: [] })]] });
    expect(r.calls[0].url).toBe("https://graph.microsoft.com/v1.0/me/drive/root/search(q='O''Brien%20%231%3F')?$top=25");
  });
});

describe('drive share password (M-29)', () => {
  it('--password-stdin reads the password from stdin', async () => {
    const orig = Object.getOwnPropertyDescriptor(process, 'stdin')!;
    Object.defineProperty(process, 'stdin', { value: Readable.from([Buffer.from('s3cret\n')]), configurable: true });
    try {
      const r = await runCommand(registerDriveCommands, ['drive', 'share', 'D', 'I', '--scope', 'anonymous', '--password-stdin'], {
        routes: [['POST', '/drives/D/items/I/createLink', () => ({ link: { webUrl: 'https://x' } })]],
      });
      expect(r.exitCode).toBe(0);
      expect((r.calls[0].body as Record<string, unknown>).password).toBe('s3cret');
      expect(r.stdout + r.stderr).not.toContain('s3cret');
    } finally {
      Object.defineProperty(process, 'stdin', orig);
    }
  });

  const withStdin = async <T>(data: string, fn: () => Promise<T>): Promise<T> => {
    const orig = Object.getOwnPropertyDescriptor(process, 'stdin')!;
    Object.defineProperty(process, 'stdin', { value: Readable.from(data ? [Buffer.from(data)] : []), configurable: true });
    try { return await fn(); } finally { Object.defineProperty(process, 'stdin', orig); }
  };

  it('--password-stdin with empty stdin (e.g. /dev/null) is a usage error, not an empty password (X-15)', async () => {
    const r = await withStdin('', () => runCommand(registerDriveCommands, ['drive', 'share', 'D', 'I', '--password-stdin']));
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('--password-stdin is refused under MCP (MS365_CLI_MCP set), even with data on stdin (X-15)', async () => {
    const r = await withStdin('pw\n', () => runCommand(registerDriveCommands, ['drive', 'share', 'D', 'I', '--password-stdin'], { env: { MS365_CLI_MCP: '1' } }));
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('--password no longer exists: secrets never come from argv (D29)', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'share', 'D', 'I', '--password', 'pw']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});

describe('single-request lists page through printList (Q-8)', () => {
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const cases: [string[], string][] = [
    [['drive', 'list'], '/me/drives'],
    [['drive', 'search', 'q'], "/me/drive/root/search(q='q')"],
    [['drive', 'permissions', 'D', 'I'], '/drives/D/items/I/permissions'],
    [['drive', 'ls', '/Photos'], '/me/drive/root:/Photos:/children'],
    [['drive', 'ls'], '/me/drive/root/children'],
  ];
  const twoPages = (prefix: string) => [['GET', prefix, (c: { url: string }) => /skiptoken=2/.test(c.url)
    ? { value: [{ id: 'b' }] }
    : { value: [{ id: 'a' }], '@odata.nextLink': `${GRAPH}${prefix}?$skiptoken=2` }]] as never;

  for (const [argv, prefix] of cases) {
    it(`${argv.join(' ')}: page 1 carries nextCursor; --all fetches both pages`, async () => {
      const r1 = await runCommand(registerDriveCommands, [...argv, '-o', 'json'], { routes: twoPages(prefix) });
      expect(r1.exitCode).toBe(0);
      expect(JSON.parse(r1.stdout).nextCursor).toEqual(expect.any(String));
      const r2 = await runCommand(registerDriveCommands, [...argv, '--all', '-o', 'json'], { routes: twoPages(prefix) });
      expect(r2.exitCode).toBe(0);
      expect(JSON.parse(r2.stdout).value.map((v: { id: string }) => v.id)).toEqual(['a', 'b']);
    });
  }

  it('drive search keeps its --top default of 25', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'search', 'q'], { routes: [['GET', '/me/', () => ({ value: [] })]] });
    expect(r.calls[0].url).toContain('$top=25');
  });
});

describe('drive delete confirmation is silent under --dry-run (DR-1)', () => {
  it('prints no "Item deleted." when nothing was sent', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'delete', 'D', 'I'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).not.toContain('Item deleted.');
    expect(r.stderr).toContain('[dry-run] DELETE');
  });

  it('still confirms a real delete', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'delete', 'D', 'I'], {
      tty: true, answer: 'y', routes: [['DELETE', '/drives/D/items/I', () => new Response(null, { status: 204 })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain('Item deleted.');
  });
});

describe('drive validation (V-18, V-19, V-21)', () => {
  it('upload of a missing local file is a usage error (exit 2, no request)', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'upload', 'D', 'P', path.join(dir, 'missing.pdf')]);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('cp of a missing local file is a usage error (exit 2, no request)', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'cp', path.join(dir, 'missing.pdf'), '/x.pdf']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('share rejects an unknown --type / --scope with exit 2 and lists the choices', async () => {
    const t = await runCommand(registerDriveCommands, ['drive', 'share', 'D', 'I', '--type', 'read']);
    expect(t.exitCode).toBe(2);
    expect(t.stderr).toContain('view, edit, embed');
    const s = await runCommand(registerDriveCommands, ['drive', 'share', 'D', 'I', '--scope', 'everyone']);
    expect(s.exitCode).toBe(2);
    expect(t.calls.length + s.calls.length).toBe(0);
  });

  const link = () => ({ id: 'P1', link: { webUrl: 'https://x/s', type: 'view', scope: 'organization' } });
  it('share -o yaml prints YAML, not prose', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'share', 'D', 'I', '-o', 'yaml'], { routes: [['POST', '/drives/', link]] });
    expect(r.exitCode).toBe(0);
    expect(r.stdout).not.toContain('Sharing link (');
    expect(r.stdout).toContain('webUrl: "https://x/s"');
  });

  it('share -o text keeps the human summary', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'share', 'D', 'I', '-o', 'text'], { routes: [['POST', '/drives/', link]] });
    expect(r.stdout).toContain('Sharing link (view, organization):');
    expect(r.stdout).toContain('  https://x/s');
  });
});

describe('drive copy: server-side copy with a monitor URL (X-7, F8)', () => {
  const MONITOR = 'https://contoso-my.sharepoint.com/_api/v2.0/monitor/abc';
  const accepted = () => new Response(null, { status: 202, headers: { Location: MONITOR } });
  const monitorJson = (body: unknown, headers: Record<string, string> = {}) =>
    new Response(JSON.stringify(body), { status: 202, headers: { 'content-type': 'application/json', ...headers } });

  it('POSTs /copy with parentReference and name; without --wait prints {status:"accepted", monitorUrl}', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'copy', 'D', 'I', '--to-parent', 'P', '--name', 'b.txt', '-o', 'json'], {
      routes: [['POST', '/drives/D/items/I/copy', accepted]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(1);
    expect(r.calls[0].body).toEqual({ parentReference: { driveId: 'D', id: 'P' }, name: 'b.txt' });
    expect(JSON.parse(r.stdout)).toEqual({ status: 'accepted', monitorUrl: MONITOR });
  });

  it('--wait polls the monitor without the bearer token, then GETs the new item from the destination drive', async () => {
    let polls = 0;
    const r = await runCommand(registerDriveCommands, ['drive', 'copy', 'D', 'I', '--to-parent', 'P', '--to-drive', 'D2', '--wait', '-o', 'json'], {
      routes: [
        ['POST', '/drives/D/items/I/copy', accepted],
        ['GET', /\/monitor\/abc/, () => ++polls === 1
          ? monitorJson({ status: 'inProgress', percentageComplete: 50 }, { 'Retry-After': '0' })
          : monitorJson({ status: 'completed', resourceId: 'NEW' })],
        ['GET', '/drives/D2/items/NEW', () => ({ id: 'NEW', name: 'a.txt' })],
      ],
    });
    expect(r.exitCode).toBe(0);
    expect((r.calls[0].body as { parentReference: unknown }).parentReference).toEqual({ driveId: 'D2', id: 'P' });
    const monitorCalls = r.calls.filter((c) => c.url === MONITOR);
    expect(monitorCalls).toHaveLength(2);
    for (const c of monitorCalls) expect(Object.keys(c.headers).map((k) => k.toLowerCase())).not.toContain('authorization');
    expect(r.calls.at(-1)!.path).toBe('/drives/D2/items/NEW');
    expect(JSON.parse(r.stdout)).toMatchObject({ id: 'NEW' });
  });

  it('a 303 from the monitor counts as completed; the resource ID comes from its Location', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'copy', 'D', 'I', '--to-parent', 'P', '--wait', '-o', 'json'], {
      routes: [
        ['POST', '/drives/D/items/I/copy', accepted],
        ['GET', /\/monitor\/abc/, () => new Response(null, { status: 303, headers: { Location: 'https://graph.microsoft.com/v1.0/drives/D/items/N3' } })],
        ['GET', '/drives/D/items/N3', () => ({ id: 'N3' })],
      ],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls.filter((c) => c.url === MONITOR)).toHaveLength(1);
    expect(JSON.parse(r.stdout).id).toBe('N3');
  });

  it('a failed copy raises copy_failed (exit 1)', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'copy', 'D', 'I', '--to-parent', 'P', '--wait'], {
      routes: [
        ['POST', '/drives/D/items/I/copy', accepted],
        ['GET', /\/monitor\/abc/, () => monitorJson({ status: 'failed', error: { code: 'nameAlreadyExists', message: 'exists' } })],
      ],
    });
    expect(r.exitCode).toBe(1);
    expect(r.stderr).toContain('exists');
  });

  it('times out with the monitor URL in the fix', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'copy', 'D', 'I', '--to-parent', 'P', '--wait', '--timeout', '1'], {
      routes: [
        ['POST', '/drives/D/items/I/copy', accepted],
        ['GET', /\/monitor\/abc/, () => monitorJson({ status: 'inProgress' }, { 'Retry-After': '5' })],
      ],
    });
    expect(r.exitCode).toBe(1);
    expect(r.stderr).toContain('Timed out');
    expect(r.stderr).toContain(MONITOR);
  });

  it('refuses a non-https monitor URL', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'copy', 'D', 'I', '--to-parent', 'P', '--wait'], {
      routes: [['POST', '/drives/D/items/I/copy', () => new Response(null, { status: 202, headers: { Location: 'http://evil/monitor/abc' } })]],
    });
    expect(r.exitCode).toBe(1);
    expect(r.calls).toHaveLength(1);
  });

  it('--dry-run previews the POST and does not poll', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'copy', 'D', 'I', '--to-parent', 'P', '--wait', '-o', 'json'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toContain('[dry-run] POST https://graph.microsoft.com/v1.0/drives/D/items/I/copy');
    expect(JSON.parse(r.stdout).dryRun).toBe(true);
  });

  it('rejects a non-integer --timeout with exit 2', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'copy', 'D', 'I', '--to-parent', 'P', '--wait', '--timeout', 'soon']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('help cross-references drive cp (upload), and cp points back at copy', async () => {
    const copy = await runCommand(registerDriveCommands, ['drive', 'copy', '--help']);
    expect(copy.stdout).toContain('drive cp');
    const cp = await runCommand(registerDriveCommands, ['drive', 'cp', '--help']);
    expect(cp.stdout).toContain('drive copy');
  });
});
