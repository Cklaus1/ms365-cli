import { describe, it, expect, beforeAll, afterAll, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { runCommand, type RecordedCall } from './helpers/cli.js';
import { MODULES, leafCommands, sampleArgv, writeSampleFiles } from './helpers/commands.js';
import { commandMeta } from '../src/catalog.js';
import { sourceFiles } from './helpers/commands.js';

/**
 * Query-pagination guard (Q spec AC-3a): every catalog read command that prints a collection pages through
 * printList, so a first page with an @odata.nextLink surfaces a nextCursor. Runs each read leaf against a mock
 * whose every GET answers with one item and a link to a (final) second page.
 */
const SINGLE = 'GETs one resource (the mock\'s collection shape is echoed back)';
const NO_CURSOR: Record<string, string> = {
  'search query': 'pages by --from/--size, not @odata.nextLink',
  'mail triage': 'post-processes one page into a report',
  'calendar find-time': 'prints computed meeting suggestions, not a Graph collection',
  'sharepoint delta': 'follows every nextLink to the deltaLink (a delta feed is only complete there)',
  'drive cat': 'streams file content (the mock answers JSON)',
  'user photo': 'downloads binary content (the mock answers JSON)',
  'meetings transcript': 'prints transcript text (the mock answers JSON)',
  'mail get': SINGLE,
  'mail read': SINGLE,
  'mail attachment': SINGLE,
  'mail shared-get': SINGLE,
  'mail rule': SINGLE,
  'calendar get': SINGLE,
  'drive root': SINGLE,
  'drive quota': SINGLE,
  'drive info': SINGLE,
  'contacts get': SINGLE,
  'chat get': SINGLE,
  'sharepoint site': SINGLE,
  'sharepoint site-by-path': SINGLE,
  'sharepoint drive': SINGLE,
  'sharepoint list': SINGLE,
  'sharepoint list-item': SINGLE,
  'sharepoint item': SINGLE,
  'planner task': SINGLE,
  'planner task-details': SINGLE,
  'onenote page': SINGLE,
  'excel range': SINGLE,
  'user me': SINGLE,
  'user ooo': SINGLE,
  'user presence': SINGLE,
  'user mailbox-settings': SINGLE,
  'user working-hours': SINGLE,
  'groups get': SINGLE,
  'meetings get': SINGLE,
  'notes get': SINGLE,
};
const NO_CURSOR_PREFIX: Record<string, string> = {
  'workflow ': 'reports aggregate several reads; no collection to page',
  'reports ': 'usage reports print one aggregated document (at most 180 daily rows, or rolled-up departments); every Graph page of a JSON report is read internally (AO-23), so there is no cursor to surface',
};

const twoPages = (call: RecordedCall) => call.path.includes('$skiptoken=')
  ? { value: [{ id: '2' }] }
  : { value: [{ id: '1' }], '@odata.nextLink': `${call.url}${call.url.includes('?') ? '&' : '?'}$skiptoken=page2` };

let tmp: string;
const cwd = process.cwd();
beforeAll(() => {
  tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-lint-'));
  writeSampleFiles(tmp);
  process.chdir(tmp);
  vi.stubEnv('MS365_CLI_TIMEZONE', 'America/Los_Angeles');
});
afterAll(() => { process.chdir(cwd); vi.unstubAllEnvs(); fs.rmSync(tmp, { recursive: true, force: true }); });

function exempt(name: string): boolean {
  return name in NO_CURSOR || Object.keys(NO_CURSOR_PREFIX).some(p => name.startsWith(p));
}

let checked = 0;
describe('read commands that print a collection surface nextCursor (QP-6, AC-3a)', () => {
  for (const register of MODULES) {
    for (const { chain, cmd } of leafCommands(register)) {
      const name = chain.join(' ');
      if (!commandMeta(name)?.readOnly || exempt(name)) continue;
      it(name, async () => {
        const argv = sampleArgv(chain, cmd);
        if (cmd.options.some(o => o.long === '--output')) argv.push('-o', 'json');
        const r = await runCommand(register, argv, { routes: [['GET', /.*/, twoPages]] });
        let data: unknown;
        try { data = JSON.parse(r.stdout); } catch { return; } // not a JSON document (binary, text, error)
        const isCollection = Array.isArray(data) || (!!data && typeof data === 'object' && Array.isArray((data as { value?: unknown }).value));
        if (!isCollection) return;
        checked++;
        expect(data, `${name} prints a collection without nextCursor; page it with printList`).toHaveProperty('nextCursor');
      });
    }
  }

  it('the guard exercised the list commands', () => {
    expect(checked).toBeGreaterThan(40);
  });

  it('allow-listed commands still exist', () => {
    const names = new Set(MODULES.flatMap(r => leafCommands(r).map(l => l.chain.join(' '))));
    expect(Object.keys(NO_CURSOR).filter(n => !names.has(n))).toEqual([]);
  });
});

/** AC-3c: user text reaches OData/KQL only through the quoting helpers. */
describe('query literal quoting (AC-3c)', () => {
  const files = sourceFiles().filter(f => !f.rel.endsWith('commands/completion.ts')); // completion quotes for shells

  it('every $search value goes through kqlQuote', () => {
    const hits = files.flatMap(f => f.lines
      .map((line, i) => ({ line, at: `${f.rel}:${i + 1}` }))
      .filter(({ line }) => /\['\$search'\]\s*=|'\$search'\s*:/.test(line) && !line.includes('kqlQuote(')));
    expect(hits.map(h => `${h.at}: ${h.line.trim()}`)).toEqual([]);
  });

  it("every '${…}' inside an OData literal goes through odataQuote", () => {
    const hits = files.flatMap(f => f.lines
      .map((line, i) => ({ line, at: `${f.rel}:${i + 1}` }))
      .filter(({ line }) => [...line.matchAll(/'\$\{([^}]*)\}/g)].some(m => !/odataQuote\(|^safe/.test(m[1]))));
    expect(hits.map(h => `${h.at}: ${h.line.trim()}`)).toEqual([]);
  });
});
