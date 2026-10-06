import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { runCli, tempDir } from '../e2e/run-cli.js';

const ENV = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_TOKEN: 'fake' };
const HTML = '<!DOCTYPE html>\r\n<html><head><title>Ünïcode "q"</title></head><body><p>a &amp; b</p></body></html>';

describe('ONENOTE-1', () => {
  it('page prints the raw HTML, not a {rawResponse} wrapper', async () => {
    const r = await runCli(['onenote', 'page', 'PG1'], {
      env: ENV,
      routes: [{ match: '/me/onenote/pages/PG1/content', headers: { 'content-type': 'text/html' }, body: HTML }],
    });
    expect(r.code).toBe(0);
    expect(r.stdout).toBe(HTML);
  });

  it('page --out saves the HTML and refuses to overwrite', async () => {
    const cwd = tempDir();
    const routes = [{ match: '/content', headers: { 'content-type': 'text/html' }, body: HTML }];
    const r = await runCli(['onenote', 'page', 'PG1', '--out', 'p.html'], { env: ENV, cwd, routes });
    expect(r.code).toBe(0);
    expect(fs.readFileSync(path.join(cwd, 'p.html'), 'utf8')).toBe(HTML);
    const again = await runCli(['onenote', 'page', 'PG1', '--out', 'p.html'], { env: ENV, cwd, routes });
    expect(again.code).toBe(1);
  });

  it('pages supports --top and --all (PAGE-1)', async () => {
    const r = await runCli(['onenote', 'pages', 'S1', '--top', '1', '--all', '-o', 'json'], {
      env: ENV,
      routes: [
        { match: 'skip=1', body: { value: [{ id: 'b' }] } },
        { match: '/me/onenote/sections/S1/pages', body: { value: [{ id: 'a' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/onenote/sections/S1/pages?$skip=1' } },
      ],
    });
    expect(r.code).toBe(0);
    expect(r.requests[0].url).toContain('$top=1');
    expect(JSON.parse(r.stdout)).toEqual([{ id: 'a' }, { id: 'b' }]);
  });
});
