import { describe, it, expect } from 'vitest';
import { runCli } from '../e2e/run-cli.js';

const ENV = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_TOKEN: 'fake' };
const GRAPH = 'https://graph.microsoft.com/v1.0';

const DL = { id: 'dl', displayName: 'All Staff DL', groupTypes: [] };
const M365 = { id: 'm365', displayName: 'Project Team', groupTypes: ['Unified'] };
const DYN = { id: 'dyn', displayName: 'Dynamic DL', groupTypes: ['DynamicMembership'] };

describe('GROUPS-1 distribution-lists', () => {
  it('excludes Unified (Microsoft 365) groups and keeps the server filter', async () => {
    const r = await runCli(['groups', 'distribution-lists', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/groups', body: { value: [DL, M365, DYN] } }],
    });
    expect(r.code).toBe(0);
    const url = new URL(r.requests[0].url);
    expect(url.pathname).toBe('/v1.0/groups');
    expect(url.searchParams.get('$filter')).toBe('mailEnabled eq true and securityEnabled eq false');
    expect(url.searchParams.get('$select')!.split(',')).toContain('groupTypes');
    expect(url.searchParams.get('$top')).toBe('25');
    expect(JSON.parse(r.stdout).value.map((g: { id: string }) => g.id)).toEqual(['dl', 'dyn']);
  });

  it('excludes Unified groups on every page with --all and honours --page-limit', async () => {
    const r = await runCli(['groups', 'distribution-lists', '--all', '--page-limit', '2', '-o', 'json'], {
      env: ENV,
      routes: [
        { match: 'skiptoken=2', body: { value: [M365, DYN], '@odata.nextLink': `${GRAPH}/groups?$skiptoken=3` } },
        { match: '/groups', body: { value: [DL, M365], '@odata.nextLink': `${GRAPH}/groups?$skiptoken=2` } },
      ],
    });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(2);
    expect(JSON.parse(r.stdout).map((g: { id: string }) => g.id)).toEqual(['dl', 'dyn']);
  });
});

describe('GROUPS-1 list --search', () => {
  it('escapes quotes in $search and sends ConsistencyLevel: eventual', async () => {
    const r = await runCli(['groups', 'list', '--search', 'say "hi"', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/groups', body: { value: [] } }],
    });
    expect(r.code).toBe(0);
    const url = new URL(r.requests[0].url);
    expect(url.searchParams.get('$search')).toBe('"displayName:say \\"hi\\""');
    expect(r.requests[0].headers['consistencylevel']).toBe('eventual');
    expect(JSON.parse(r.stdout)).toEqual({ value: [] });
  });
});

describe('PAGE-1 groups lists', () => {
  for (const [cmd, path] of [['my', '/me/memberOf/microsoft.graph.group'], ['teams', '/groups'], ['list', '/groups']] as const) {
    it(`groups ${cmd} honours --page-limit with --all`, async () => {
      const r = await runCli(['groups', cmd, '--all', '--page-limit', '1', '-o', 'json'], {
        env: ENV,
        routes: [{ match: path, body: { value: [{ id: 'g1' }], '@odata.nextLink': `${GRAPH}${path}?$skiptoken=2` } }],
      });
      expect(r.code).toBe(0);
      expect(r.requests).toHaveLength(1);
      expect(new URL(r.requests[0].url).pathname).toBe(`/v1.0${path}`);
      expect(JSON.parse(r.stdout)).toEqual([{ id: 'g1' }]);
    });
  }

  it('groups members rejects a non-integer --top', async () => {
    const r = await runCli(['groups', 'members', 'G1', '--top', 'ten'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });
});
