import { describe, it, expect } from 'vitest';
import { runCommand } from '../helpers/cli.js';
import { registerSharePointCommands } from '../../src/commands/sharepoint.js';
import { registerGroupsCommands } from '../../src/commands/groups.js';
import { registerContactsCommands } from '../../src/commands/contacts.js';
import { registerOneNoteCommands } from '../../src/commands/onenote.js';

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
