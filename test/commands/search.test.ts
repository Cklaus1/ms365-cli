import { describe, it, expect } from 'vitest';
import { runCli } from '../e2e/run-cli.js';

const ENV = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_TOKEN: 'fake' };

describe('SEARCH-1 entity type combinations', () => {
  it('rejects driveItem,message,site before any request', async () => {
    const r = await runCli(['search', 'query', 'budget', '--entity', 'driveItem,message,site'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.stderr).toMatch(/cannot be combined/i);
    expect(r.requests).toHaveLength(0);
  });

  it('rejects unknown entity types', async () => {
    const r = await runCli(['search', 'query', 'x', '--entity', 'mail'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });

  it('accepts a SharePoint bucket (driveItem,site,listItem) and sends one request', async () => {
    const r = await runCli(['search', 'query', 'budget', '--entity', 'driveItem, site,listItem', '--top', '5', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/search/query', method: 'POST', body: { value: [] } }],
    });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.requests[0].body!)).toEqual({
      requests: [{ entityTypes: ['driveItem', 'site', 'listItem'], query: { queryString: 'budget' }, from: 0, size: 5 }],
    });
    expect(JSON.parse(r.stdout)).toEqual({ value: [] });
  });

  it('works under --read-only (read-only POST action)', async () => {
    const r = await runCli(['--read-only', 'search', 'query', 'x', '--entity', 'message'], {
      env: ENV,
      routes: [{ match: '/search/query', method: 'POST', body: { value: [] } }],
    });
    expect(r.code).toBe(0);
  });
});

describe('MAIL-8 people $search is quote-escaped', () => {
  it('wraps the phrase in quotes and escapes embedded quotes', async () => {
    const r = await runCli(['search', 'people', 'say "hi"', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/me/people', body: { value: [] } }],
    });
    expect(r.code).toBe(0);
    const url = new URL(r.requests[0].url);
    expect(url.searchParams.get('$search')).toBe('"say \\"hi\\""');
    expect(url.searchParams.get('$top')).toBe('10');
  });
});
