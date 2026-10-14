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

  it('--password no longer exists: secrets never come from argv (D29)', async () => {
    const r = await runCommand(registerDriveCommands, ['drive', 'share', 'D', 'I', '--password', 'pw']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});
