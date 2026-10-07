import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { runCli, tempDir } from '../e2e/run-cli.js';

const ENV = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_TOKEN: 'fake' };

/** Bytes that any text/JSON decoding would corrupt. */
const PDF = Buffer.concat([Buffer.from('%PDF-1.7\r\n'), Buffer.from([0x00, 0xff, 0xfe, 0x80, 0x0d, 0x0a, 0xc3]), Buffer.alloc(5000, 7)]);
const PDF_ROUTE = { headers: { 'content-type': 'application/pdf' }, bodyBase64: PDF.toString('base64') };

function localFile(cwd: string, name = 'x.txt', content = 'hello'): string {
  fs.writeFileSync(path.join(cwd, name), content);
  return name;
}

describe('DRIVE-1 user paths are encoded per segment', () => {
  it('cp to "/Docs/notes#v2.md" previews root:/Docs/notes%23v2.md:/content with replace', async () => {
    const cwd = tempDir();
    const r = await runCli(['--dry-run', 'drive', 'cp', localFile(cwd), '/Docs/notes#v2.md', '-o', 'json'], { env: ENV, cwd, routes: [] });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toContain('/me/drive/root:/Docs/notes%23v2.md:/content?@microsoft.graph.conflictBehavior=replace');
    expect(JSON.parse(r.stdout).url).toContain('root:/Docs/notes%23v2.md:/content');
  });

  it('cp encodes ? and % and keeps them in the item name', async () => {
    const cwd = tempDir();
    const r = await runCli(['--dry-run', 'drive', 'cp', localFile(cwd), '/a?b/100%25.txt'], { env: ENV, cwd, routes: [] });
    expect(r.code).toBe(0);
    expect(r.stderr).toContain('root:/a%3Fb/100%2525.txt:/content');
  });

  it('cp to a folder path ending in / uploads under the local file name', async () => {
    const cwd = tempDir();
    const r = await runCli(['--dry-run', 'drive', 'cp', localFile(cwd, 'r e.pdf'), '/Docs/'], { env: ENV, cwd, routes: [] });
    expect(r.code).toBe(0);
    expect(r.stderr).toContain('root:/Docs/r%20e.pdf:/content');
  });

  it('ls / lists the root without a colon path; nested paths and trailing / are normalised', async () => {
    const routes = [{ match: '/children', body: { value: [] } }];
    const root = await runCli(['drive', 'ls', '/', '-o', 'json'], { env: ENV, routes });
    expect(root.code).toBe(0);
    expect(root.requests[0].url).toMatch(/\/me\/drive\/root\/children(\?|$)/);
    expect(JSON.parse(root.stdout)).toEqual({ value: [] });

    const nested = await runCli(['drive', 'ls', '/Docs//Q#1/', '-o', 'json'], { env: ENV, routes });
    expect(nested.code).toBe(0);
    expect(nested.requests[0].url).toContain('/me/drive/root:/Docs/Q%231:/children');
  });

  it('search escapes quotes and URL-encodes the query inside the function call', async () => {
    const r = await runCli(['drive', 'search', "it's #1?", '-o', 'json'], { env: ENV, routes: [{ match: 'search', body: { value: [] } }] });
    expect(r.code).toBe(0);
    expect(r.requests[0].url).toContain("/me/drive/root/search(q='it''s%20%231%3F')");
    expect(r.requests[0].url).toContain('$top=25');
  });
});

describe('GRAPH-3 drive upload conflict behaviour', () => {
  it('defaults to conflictBehavior=rename and reports the stored name', async () => {
    const cwd = tempDir();
    const stored = { id: '1', name: 'report 2.docx', size: 5 };
    const r = await runCli(['drive', 'upload', 'D1', 'P1', localFile(cwd, 'report.docx'), '-o', 'json'], {
      env: ENV,
      cwd,
      routes: [{ match: ':/content', method: 'PUT', status: 201, body: { '@odata.context': 'x', ...stored } }],
    });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(1);
    expect(r.requests[0].method).toBe('PUT');
    expect(r.requests[0].url).toContain('/drives/D1/items/P1:/report.docx:/content?@microsoft.graph.conflictBehavior=rename');
    expect(JSON.parse(r.stdout)).toEqual(stored);
    expect(r.stderr).toContain(stored.name);
  });

  it('--conflict fail is sent; invalid values are rejected', async () => {
    const cwd = tempDir();
    const ok = await runCli(['--dry-run', 'drive', 'upload', 'D1', 'P1', localFile(cwd, 'a#b.txt'), '--conflict', 'fail'], { env: ENV, cwd, routes: [] });
    expect(ok.code).toBe(0);
    expect(ok.stderr).toContain('/items/P1:/a%23b.txt:/content?@microsoft.graph.conflictBehavior=fail');

    const bad = await runCli(['--dry-run', 'drive', 'upload', 'D1', 'P1', localFile(cwd), '--conflict', 'keep'], { env: ENV, cwd, routes: [] });
    expect(bad.code).toBe(1);
  });

  it('--quiet hides the upload notice', async () => {
    const cwd = tempDir();
    const r = await runCli(['-q', 'drive', 'upload', 'D1', 'P1', localFile(cwd), '-o', 'json'], {
      env: ENV,
      cwd,
      routes: [{ match: ':/content', method: 'PUT', status: 201, body: { id: '1', name: 'x.txt' } }],
    });
    expect(r.code).toBe(0);
    expect(r.stderr).toBe('');
  });
});

describe('GRAPH-2 / DRIVE-2 binary downloads', () => {
  it('cat --out writes a mocked application/pdf byte-identical', async () => {
    const cwd = tempDir();
    const r = await runCli(['drive', 'cat', '/Docs/r#1.pdf', '--out', 'r.pdf'], {
      env: ENV,
      cwd,
      routes: [{ match: 'root:/Docs/r%231.pdf:/content', ...PDF_ROUTE }],
    });
    expect(r.code).toBe(0);
    expect(r.stdout).toBe('');
    expect(fs.readFileSync(path.join(cwd, 'r.pdf')).equals(PDF)).toBe(true);
  });

  it('cat without --out streams the raw bytes to stdout (JSON files are not re-rendered)', async () => {
    const text = '{"@odata.context":"x","a":1}\n';
    const r = await runCli(['drive', 'cat', '/a.json'], {
      env: ENV,
      routes: [{ match: 'root:/a.json:/content', body: text }],
    });
    expect(r.code).toBe(0);
    expect(r.stdout).toBe(text);
  });

  it('download --out is byte-identical', async () => {
    const cwd = tempDir();
    const r = await runCli(['drive', 'download', 'D1', 'I1', '--out', 'b.pdf'], {
      env: ENV,
      cwd,
      routes: [{ match: '/drives/D1/items/I1/content', ...PDF_ROUTE }],
    });
    expect(r.code).toBe(0);
    expect(fs.readFileSync(path.join(cwd, 'b.pdf')).equals(PDF)).toBe(true);
  });

  it('DRIVE-3 the old 3-argument form exits 1 with too many arguments and sends nothing', async () => {
    const cwd = tempDir();
    const r = await runCli(['drive', 'download', 'D1', 'I1', 'C1', '--out', 'b.pdf'], { env: ENV, cwd, routes: [{ match: '/content', ...PDF_ROUTE }] });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('too many arguments');
    expect(r.requests).toHaveLength(0);
    expect(fs.readdirSync(cwd)).toEqual([]);
  });

  it('download-shared --out is byte-identical', async () => {
    const cwd = tempDir();
    const r = await runCli(['drive', 'download-shared', 'u!abc', '--out', 's.pdf'], {
      env: ENV,
      cwd,
      routes: [{ match: '/shares/u!abc/driveItem/content', ...PDF_ROUTE }],
    });
    expect(r.code).toBe(0);
    expect(fs.readFileSync(path.join(cwd, 's.pdf')).equals(PDF)).toBe(true);
  });

  it('an existing --out without --force exits 1 and leaves the file intact; --force overwrites', async () => {
    const cwd = tempDir();
    fs.writeFileSync(path.join(cwd, 'keep.pdf'), 'original');
    const routes = [{ match: '/content', ...PDF_ROUTE }];
    const r = await runCli(['drive', 'download', 'D1', 'I1', '--out', 'keep.pdf'], { env: ENV, cwd, routes });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('--force');
    expect(fs.readFileSync(path.join(cwd, 'keep.pdf'), 'utf8')).toBe('original');
    expect(fs.readdirSync(cwd)).toEqual(['keep.pdf']);

    const f = await runCli(['drive', 'download', 'D1', 'I1', '--out', 'keep.pdf', '--force'], { env: ENV, cwd, routes });
    expect(f.code).toBe(0);
    expect(fs.readFileSync(path.join(cwd, 'keep.pdf')).equals(PDF)).toBe(true);
  });

  it('dry-run download writes no file and nothing to stdout', async () => {
    const cwd = tempDir();
    const r = await runCli(['--dry-run', 'drive', 'download', 'D1', 'I1', '--out', 'b.pdf'], { env: ENV, cwd, routes: [] });
    expect(r.code).toBe(0);
    expect(r.stdout).toBe('');
    expect(fs.readdirSync(cwd)).toEqual([]);
  });
});

describe('CORE-4 / PAGE-1 drive listings', () => {
  it('files --all -o json follows nextLink and prints one JSON array', async () => {
    const r = await runCli(['drive', 'files', 'D1', 'F1', '--all', '--top', '2', '-o', 'json'], {
      env: ENV,
      routes: [
        { match: 'skiptoken=2', body: { value: [{ id: 'b' }] } },
        { match: '/drives/D1/items/F1/children', body: { value: [{ id: 'a' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/drives/D1/items/F1/children?$skiptoken=2' } },
      ],
    });
    expect(r.code).toBe(0);
    expect(r.requests[0].url).toContain('$top=2');
    expect(JSON.parse(r.stdout)).toEqual([{ id: 'a' }, { id: 'b' }]);
  });

  it('share -o yaml renders the response instead of the text summary', async () => {
    const r = await runCli(['drive', 'share', 'D1', 'I1', '-o', 'yaml'], {
      env: ENV,
      routes: [{ match: 'createLink', method: 'POST', body: { link: { webUrl: 'https://x/y' } } }],
    });
    expect(r.code).toBe(0);
    expect(r.stdout).toContain('webUrl');
    expect(r.stdout).not.toContain('Sharing link');
  });

  it('quota -o json parses', async () => {
    const r = await runCli(['drive', 'quota', '-o', 'json'], { env: ENV, routes: [{ match: '/me/drive', body: { quota: { used: 1, total: 2 } } }] });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout).quota.used).toBe(1);
  });
});

describe('DRY-1 drive upload under --dry-run', () => {
  it('prints the preview but no Uploading line', async () => {
    const cwd = tempDir();
    const r = await runCli(['--dry-run', 'drive', 'upload', 'D1', 'P1', localFile(cwd)], { env: ENV, cwd, routes: [] });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toContain('[dry-run]');
    expect(r.stderr).not.toContain('Uploading');
  });
});

describe('PAGE-2 drive permissions drains every page', () => {
  it('prints one combined JSON array from two pages', async () => {
    const r = await runCli(['drive', 'permissions', 'D1', 'I1', '-o', 'json'], {
      env: ENV,
      routes: [
        { match: 'skiptoken=2', body: { value: [{ id: 'b' }] } },
        { match: '/drives/D1/items/I1/permissions', body: { value: [{ id: 'a' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/drives/D1/items/I1/permissions?$skiptoken=2' } },
      ],
    });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout)).toEqual([{ id: 'a' }, { id: 'b' }]);
    expect(r.requests).toHaveLength(2);
  });
});

describe('CHOICE-1 drive share choices', () => {
  for (const args of [['--type', 'bogus'], ['--scope', 'everyone']]) {
    it(`share ${args.join(' ')} exits 1 with Allowed choices before any request`, async () => {
      const r = await runCli(['drive', 'share', 'D1', 'I1', ...args], { env: ENV, routes: [] });
      expect(r.code).toBe(1);
      expect(r.stderr).toContain('Allowed choices');
      expect(r.requests).toHaveLength(0);
    });
  }

  it('share --type edit --scope users sends them', async () => {
    const r = await runCli(['drive', 'share', 'D1', 'I1', '--type', 'edit', '--scope', 'users', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/createLink', body: { link: { webUrl: 'https://x' } } }],
    });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.requests[0].body ?? '')).toEqual({ type: 'edit', scope: 'users' });
  });
});
