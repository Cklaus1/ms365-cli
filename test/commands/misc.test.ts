import { describe, it, expect } from 'vitest';
import { runCommand } from '../helpers/cli.js';
import { registerSharePointCommands } from '../../src/commands/sharepoint.js';
import { registerGroupsCommands } from '../../src/commands/groups.js';
import { registerContactsCommands } from '../../src/commands/contacts.js';
import { registerOneNoteCommands } from '../../src/commands/onenote.js';
import { registerSearchCommands } from '../../src/commands/search.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';

describe('sharepoint site-by-path (C-10)', () => {
  it('encodes # and ? inside path segments', async () => {
    const r = await runCommand(registerSharePointCommands, ['sp', 'site-by-path', 'contoso.sharepoint.com', '/sites/R#D?x', '-o', 'json'], {
      routes: [['GET', '/sites/', () => ({ id: 'S1' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].url).toBe(`${GRAPH}/sites/contoso.sharepoint.com:/sites/R%23D%3Fx`);
  });

  it('keeps slashes and adds a leading slash; spaces are encoded', async () => {
    const r = await runCommand(registerSharePointCommands, ['sp', 'site-by-path', 'contoso.sharepoint.com', 'sites/Team Site', '-o', 'json'], {
      routes: [['GET', '/sites/', () => ({ id: 'S1' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].url).toBe(`${GRAPH}/sites/contoso.sharepoint.com:/sites/Team%20Site`);
  });
});

// Two pages available; --page-limit 1 must stop after the first.
function pagedRoutes(prefix: string): Parameters<typeof runCommand>[2] {
  return {
    routes: [['GET', prefix, (call) => call.path.includes('skiptoken')
      ? { value: [{ id: 'B' }] }
      : { value: [{ id: 'A' }], '@odata.nextLink': `${GRAPH}${prefix}?$skiptoken=2` }]],
  };
}

describe('--page-limit is passed through (M-26)', () => {
  const cases: Array<[string, Parameters<typeof runCommand>[0], string[], string]> = [
    ['contacts list', registerContactsCommands, ['contacts', 'list'], '/me/contacts'],
    ['groups list', registerGroupsCommands, ['groups', 'list'], '/groups'],
    ['groups my', registerGroupsCommands, ['groups', 'my'], '/me/memberOf/microsoft.graph.group'],
    ['groups distribution-lists', registerGroupsCommands, ['groups', 'distribution-lists'], '/groups'],
    ['groups teams', registerGroupsCommands, ['groups', 'teams'], '/groups'],
  ];
  for (const [name, register, argv, prefix] of cases) {
    it(name, async () => {
      const r = await runCommand(register, [...argv, '--all', '--page-limit', '1', '-o', 'json'], pagedRoutes(prefix));
      expect(r.exitCode).toBe(0);
      expect(r.calls).toHaveLength(1);
    });
  }
});

describe('onenote page', () => {
  it('prints the page HTML as-is', async () => {
    const html = '<html><body><p>Hello</p></body></html>';
    const r = await runCommand(registerOneNoteCommands, ['onenote', 'page', 'P1'], {
      routes: [['GET', '/me/onenote/pages/P1/content', () => new Response(html, { headers: { 'Content-Type': 'text/html' } })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.stdout.trim()).toBe(html);
  });
});

describe('sharepoint/onenote single-request lists page through printList (Q-8)', () => {
  const cases: [typeof registerSharePointCommands, string[], string][] = [
    [registerSharePointCommands, ['sp', 'search', 'mkt'], '/sites'],
    [registerSharePointCommands, ['sp', 'drives', 'S'], '/sites/S/drives'],
    [registerSharePointCommands, ['sp', 'lists', 'S'], '/sites/S/lists'],
    [registerOneNoteCommands, ['onenote', 'notebooks'], '/me/onenote/notebooks'],
    [registerOneNoteCommands, ['onenote', 'sections', 'N'], '/me/onenote/notebooks/N/sections'],
    [registerOneNoteCommands, ['onenote', 'pages', 'S'], '/me/onenote/sections/S/pages'],
  ];
  const twoPages = (prefix: string) => [['GET', prefix, (c: { url: string }) => /skiptoken=2/.test(c.url)
    ? { value: [{ id: 'b' }] }
    : { value: [{ id: 'a' }], '@odata.nextLink': `${GRAPH}${prefix}?$skiptoken=2` }]] as never;

  for (const [register, argv, prefix] of cases) {
    it(`${argv.join(' ')}: page 1 carries nextCursor; --all fetches both pages`, async () => {
      const r1 = await runCommand(register, [...argv, '-o', 'json'], { routes: twoPages(prefix) });
      expect(r1.exitCode).toBe(0);
      expect(JSON.parse(r1.stdout).nextCursor).toEqual(expect.any(String));
      const r2 = await runCommand(register, [...argv, '--all', '-o', 'json'], { routes: twoPages(prefix) });
      expect(r2.exitCode).toBe(0);
      expect(JSON.parse(r2.stdout).value.map((v: { id: string }) => v.id)).toEqual(['a', 'b']);
    });
  }

  it('sp delta follows nextLinks to the deltaLink', async () => {
    const r = await runCommand(registerSharePointCommands, ['sp', 'delta', '-o', 'json'], {
      routes: [['GET', '/sites/delta()', (c) => /skiptoken=2/.test(c.url)
        ? { value: [{ id: 'b' }], '@odata.deltaLink': `${GRAPH}/sites/delta()?token=T` }
        : { value: [{ id: 'a' }], '@odata.nextLink': `${GRAPH}/sites/delta()?$skiptoken=2` }]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(2);
    const out = JSON.parse(r.stdout);
    expect(out.value.map((v: { id: string }) => v.id)).toEqual(['a', 'b']);
    expect(out.deltaLink).toBe(`${GRAPH}/sites/delta()?token=T`);
  });
});

describe('$search terms are KQL-quoted (Q-1 to Q-3, AC-1)', () => {
  const term = 'a "b" \\c';
  const searchOf = (path: string) => new URLSearchParams(path.split('?')[1]).get('$search');

  it('contacts list --search', async () => {
    const r = await runCommand(registerContactsCommands, ['contacts', 'list', '--search', term, '-o', 'json']);
    expect(r.exitCode).toBe(0);
    expect(searchOf(r.calls[0].path)).toBe('"a \\"b\\" \\\\c"');
  });

  it('search people', async () => {
    const r = await runCommand(registerSearchCommands, ['search', 'people', term, '-o', 'json']);
    expect(r.exitCode).toBe(0);
    expect(searchOf(r.calls[0].path)).toBe('"a \\"b\\" \\\\c"');
  });

  it('groups list --search quotes the whole displayName clause', async () => {
    const r = await runCommand(registerGroupsCommands, ['groups', 'list', '--search', term, '-o', 'json']);
    expect(r.exitCode).toBe(0);
    expect(searchOf(r.calls[0].path)).toBe('"displayName:a \\"b\\" \\\\c"');
    expect(r.calls[0].headers['ConsistencyLevel'] ?? r.calls[0].headers['consistencylevel']).toBe('eventual');
  });
});

describe('contacts list --search rejects --filter/--orderby (Q-4)', () => {
  it.each([['--filter', 'x eq 1'], ['--orderby', 'displayName']])('--search with %s exits 2, nothing sent', async (flag, value) => {
    const r = await runCommand(registerContactsCommands, ['contacts', 'list', '--search', 'bob', flag, value]);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toMatch(/usage|--search cannot be combined/);
  });

  it('groups list --search may combine with --orderby (directory $search)', async () => {
    const r = await runCommand(registerGroupsCommands, ['groups', 'list', '--search', 'x', '--orderby', 'displayName', '-o', 'json']);
    expect(r.exitCode).toBe(0);
  });
});

describe('search people and groups owners paginate (Q-8)', () => {
  const twoPages = (prefix: string) => [['GET', prefix, (c: { url: string }) => /skiptoken=2/.test(c.url)
    ? { value: [{ id: 'b' }] }
    : { value: [{ id: 'a' }], '@odata.nextLink': `${GRAPH}${prefix}?$skiptoken=2` }]] as never;
  const cases: Array<[Parameters<typeof runCommand>[0], string[], string]> = [
    [registerSearchCommands, ['search', 'people', 'ada'], '/me/people'],
    [registerGroupsCommands, ['groups', 'owners', 'G1'], '/groups/G1/owners'],
  ];
  for (const [register, argv, prefix] of cases) {
    it(`${argv.join(' ')}: page 1 carries nextCursor; --all fetches both pages`, async () => {
      const r1 = await runCommand(register, [...argv, '-o', 'json'], { routes: twoPages(prefix) });
      expect(r1.exitCode).toBe(0);
      expect(JSON.parse(r1.stdout).nextCursor).toEqual(expect.any(String));
      const r2 = await runCommand(register, [...argv, '--all', '-o', 'json'], { routes: twoPages(prefix) });
      expect(r2.exitCode).toBe(0);
      expect(JSON.parse(r2.stdout).value.map((v: { id: string }) => v.id)).toEqual(['a', 'b']);
    });
  }

  it('search people keeps its --top 10 default', async () => {
    const r = await runCommand(registerSearchCommands, ['search', 'people', '-o', 'json']);
    expect(r.calls[0].path).toContain('$top=10');
  });
});

describe('contacts delete under --dry-run prints no confirmation (DR-1)', () => {
  it('stays silent and previews the DELETE', async () => {
    const r = await runCommand(registerContactsCommands, ['contacts', 'delete', 'C1'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).not.toContain('Contact deleted.');
  });
});
