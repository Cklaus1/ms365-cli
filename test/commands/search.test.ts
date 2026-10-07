import { describe, it, expect } from 'vitest';
import { REMOTE_ENTRY, runCli } from '../e2e/run-cli.js';
import { call, capsHeader, startServer } from '../e2e/serve-harness.js';

const ENV = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_TOKEN: 'fake' };

describe('SEARCH-1 entity type combinations', () => {
  it('rejects driveItem,message,site before any request', async () => {
    const r = await runCli(['search', 'query', 'budget', '--entity', 'driveItem,message,site'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.stderr).toMatch(/cannot be combined/i);
    expect(r.requests).toHaveLength(0);
  });

  it('CHOICE-1 rejects unknown entity types with the allowed choices', async () => {
    const r = await runCli(['search', 'query', 'x', '--entity', 'mail'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('Allowed choices');
    expect(r.stderr).toContain('driveItem');
    expect(r.requests).toHaveLength(0);
  });

  it('CHOICE-1 rejects an empty entity list', async () => {
    const r = await runCli(['search', 'query', 'x', '--entity', ' , '], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('Allowed choices');
    expect(r.requests).toHaveLength(0);
  });

  it('CHOICE-1 ms365-remote rejects a bad entity client-side (no HTTP request)', async () => {
    // Port 9 (discard) is never served: any request would fail with a connection error instead.
    const r = await runCli(['search', 'query', 'x', '--entity', 'mail'], { entry: REMOTE_ENTRY, env: { MS365_REMOTE_URL: 'http://127.0.0.1:9' } });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('Allowed choices');
    expect(r.stderr).not.toMatch(/ECONNREFUSED|fetch failed/);
  });

  it('CHOICE-1 POST /v1/execute with a bad entity returns 400 invalid_args', async () => {
    const server = await startServer({ routes: [] });
    try {
      const caps = capsHeader(['search.read']);
      const bad = await call(server, { caps, body: { command: 'search.query', positionals: ['x'], args: { entity: 'driveItem,mail' } } });
      expect(bad.status).toBe(400);
      expect(bad.body.error).toMatchObject({ code: 'invalid_args', message: expect.stringContaining('entity must be a comma-separated list of') });
    } finally {
      await server.stop();
    }
    expect(server.requests()).toHaveLength(0);
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
