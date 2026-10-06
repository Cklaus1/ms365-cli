import { describe, it, expect } from 'vitest';
import { runCli } from '../e2e/run-cli.js';

const ENV = {
  MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000',
  MS365_CLI_TOKEN: 'fake',
};

describe('contacts list (MAIL-8, PAGE-1)', () => {
  it('quote-escapes --search and honours --top/--page-limit', async () => {
    const r = await runCli(['contacts', 'list', '--search', 'say "hi"', '--top', '2', '--all', '--page-limit', '2', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/me/contacts', method: 'GET', body: { value: [{ id: 'C1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/contacts?$skiptoken=n' } }],
    });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(2);
    const url = new URL(r.requests[0].url);
    expect(url.searchParams.get('$search')).toBe('"say \\"hi\\""');
    expect(url.searchParams.get('$top')).toBe('2');
    expect(JSON.parse(r.stdout)).toEqual([{ id: 'C1' }, { id: 'C1' }]);
  });
});

describe('contacts update (TODO-1)', () => {
  it('no fields: no request, exit 1', async () => {
    const r = await runCli(['contacts', 'update', 'C1'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toMatch(/nothing to update/i);
  });

  it('--replace-email without --email: no request, exit 1', async () => {
    const r = await runCli(['contacts', 'update', 'C1', '--replace-email'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });

  it('sends only the given fields', async () => {
    const r = await runCli(['contacts', 'update', 'C1', '--phone', '123', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/me/contacts/C1$', method: 'PATCH', body: { id: 'C1' } }],
    });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.requests[0].body!)).toEqual({ mobilePhone: '123' });
  });
});
