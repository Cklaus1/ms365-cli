import { describe, it, expect } from 'vitest';
import { runCli } from '../e2e/run-cli.js';

const ENV = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_TOKEN: 'fake' };
const BASE = 'https://graph.microsoft.com/v1.0';

describe('SP-1', () => {
  it('sp items is gone', async () => {
    const r = await runCli(['sp', 'items', 'S1'], { env: ENV, routes: [{ match: '.', body: { value: [] } }] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toMatch(/unknown command|too many arguments/i);
  });
});

describe('SP-2 delta follows nextLink and surfaces the deltaLink', () => {
  it('--all drains all pages and prints value + deltaLink', async () => {
    const r = await runCli(['sp', 'delta', '--all', '-o', 'json'], {
      env: ENV,
      routes: [
        { match: 'token=p2', body: { value: [{ id: 'b' }], '@odata.deltaLink': `${BASE}/sites/delta()?token=d1` } },
        { match: '/sites/delta\\(\\)$', body: { value: [{ id: 'a' }], '@odata.nextLink': `${BASE}/sites/delta()?token=p2` } },
      ],
    });
    expect(r.code).toBe(0);
    expect(r.requests.map((q) => q.url)).toEqual([`${BASE}/sites/delta()`, `${BASE}/sites/delta()?token=p2`]);
    expect(JSON.parse(r.stdout)).toEqual({ value: [{ id: 'a' }, { id: 'b' }], deltaLink: `${BASE}/sites/delta()?token=d1` });
  });

  it('without --all prints one page and the nextLink to continue', async () => {
    const r = await runCli(['sp', 'delta', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/sites/delta\\(\\)$', body: { '@odata.context': 'x', value: [{ id: 'a' }], '@odata.nextLink': `${BASE}/sites/delta()?token=p2` } }],
    });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(1);
    expect(JSON.parse(r.stdout)).toEqual({ value: [{ id: 'a' }], nextLink: `${BASE}/sites/delta()?token=p2` });
  });

  it('--delta-link resumes from a previous link', async () => {
    const r = await runCli(['sp', 'delta', '--delta-link', `${BASE}/sites/delta()?token=d1`, '-o', 'json'], {
      env: ENV,
      routes: [{ match: 'token=d1', body: { value: [], '@odata.deltaLink': `${BASE}/sites/delta()?token=d2` } }],
    });
    expect(r.code).toBe(0);
    expect(r.requests.map((q) => q.url)).toEqual([`${BASE}/sites/delta()?token=d1`]);
    expect(JSON.parse(r.stdout)).toEqual({ value: [], deltaLink: `${BASE}/sites/delta()?token=d2` });
  });

  it('rejects a link that is not a Graph delta URL', async () => {
    const r = await runCli(['sp', 'delta', '--delta-link', 'https://evil.example/sites/delta()'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });
});

describe('DRIVE-1 sp site-by-path encodes the server-relative path', () => {
  it('encodes # and spaces per segment', async () => {
    const r = await runCli(['--dry-run', 'sp', 'site-by-path', 'contoso.sharepoint.com', 'sites/R&D #1'], { env: ENV, routes: [] });
    expect(r.code).toBe(0);
    expect(r.stderr).toContain(`${BASE}/sites/contoso.sharepoint.com:/sites/R%26D%20%231\n`);
  });

  it("refuses '..' segments instead of letting URL parsing climb to another Graph path", async () => {
    const r = await runCli(['--org-mode', 'sp', 'site-by-path', 'contoso.sharepoint.com', '../../me/messages', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/me/messages', body: { value: [{ id: 'm1', subject: 'secret' }] } }],
    });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain("'.' or '..'");
    expect(r.stdout).toBe('');
    expect(r.requests).toHaveLength(0);
  });
});

describe('PAGE-1 list-items', () => {
  it('--top and --all page through and print one JSON array', async () => {
    const r = await runCli(['sp', 'list-items', 'S1', 'L1', '--top', '1', '--all', '-o', 'json'], {
      env: ENV,
      routes: [
        { match: 'skiptoken=2', body: { value: [{ id: '2' }] } },
        { match: '/sites/S1/lists/L1/items', body: { value: [{ id: '1' }], '@odata.nextLink': `${BASE}/sites/S1/lists/L1/items?$skiptoken=2` } },
      ],
    });
    expect(r.code).toBe(0);
    expect(r.requests[0].url).toContain('$top=1');
    expect(JSON.parse(r.stdout)).toEqual([{ id: '1' }, { id: '2' }]);
  });
});
