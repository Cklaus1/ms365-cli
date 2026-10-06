import { describe, it, expect, afterEach, vi } from 'vitest';
import { Command } from 'commander';
import { GraphClient } from '../src/graph-client.js';
import { paginationOptions, listRequest, type ListOptions } from '../src/list.js';
import { mockFetch, jsonResponse, fakeAuth, type Handler } from './helpers/fetch-mock.js';

const BASE = 'https://graph.microsoft.com/v1.0';

let restore: (() => void) | undefined;
afterEach(() => {
  restore?.();
  restore = undefined;
  vi.restoreAllMocks();
});

function install(handler: Handler) {
  const m = mockFetch(handler);
  restore = m.restore;
  vi.spyOn(console, 'error').mockImplementation(() => {});
  return m.calls;
}

/** Three pages of two items each. */
const pages: Handler = (call) => {
  const page = Number(new URL(call.url).searchParams.get('page') ?? '1');
  return jsonResponse({
    value: [{ id: `${page}a` }, { id: `${page}b` }],
    ...(page < 3 ? { '@odata.nextLink': `${BASE}/items?page=${page + 1}` } : {}),
  });
};

/** Parse argv with the shared options; returns the parsed ListOptions (or throws the commander error). */
function parse(argv: string[], defaultTop?: number): ListOptions {
  let parsed: ListOptions | undefined;
  const cmd = paginationOptions(new Command('list'), defaultTop)
    .option('-o, --output <format>', 'fmt', 'json')
    .exitOverride()
    .configureOutput({ writeErr: () => {}, writeOut: () => {} })
    .action((opts: ListOptions) => { parsed = opts; });
  cmd.parse(argv, { from: 'user' });
  if (!parsed) throw new Error('action not run');
  return parsed;
}

async function run(argv: string[], defaultTop?: number, extra: Partial<Parameters<typeof listRequest>[3]> = {}): Promise<string> {
  const out: string[] = [];
  await listRequest(new GraphClient(fakeAuth()), '/items', {}, { ...parse(argv, defaultTop), ...extra }, (s) => out.push(s));
  return out.join('');
}

describe('PAGE-1 paginationOptions', () => {
  it('parses --top/--page-limit/--page-delay as numbers, --all as boolean', () => {
    expect(parse(['--top', '5', '--all', '--page-limit', '2', '--page-delay', '0'])).toMatchObject({ top: 5, all: true, pageLimit: 2, pageDelay: 0 });
  });

  it('applies the default --top', () => {
    expect(parse([], 25).top).toBe(25);
    expect(parse([]).top).toBeUndefined();
  });

  it.each([
    [['--top', 'abc']],
    [['--top', '0']],
    [['--top', '1.5']],
    [['--page-limit', '-1']],
    [['--page-delay', 'soon']],
  ])('rejects %j', (argv) => {
    expect(() => parse(argv)).toThrow();
  });
});

describe('PAGE-1 listRequest', () => {
  it('without --all: one request with $top, prints that page', async () => {
    const calls = install(pages);
    const out = await run(['--top', '2']);
    expect(calls).toHaveLength(1);
    expect(new URL(calls[0].url).searchParams.get('$top')).toBe('2');
    expect(JSON.parse(out)).toEqual({ value: [{ id: '1a' }, { id: '1b' }] });
  });

  it('--all follows nextLinks; json streams a single array', async () => {
    const calls = install(pages);
    const out = await run(['--all']);
    expect(calls).toHaveLength(3);
    expect(JSON.parse(out)).toEqual(['1a', '1b', '2a', '2b', '3a', '3b'].map((id) => ({ id })));
  });

  it('--all --page-limit stops early', async () => {
    const calls = install(pages);
    const out = await run(['--all', '--page-limit', '2']);
    expect(calls).toHaveLength(2);
    expect(JSON.parse(out)).toHaveLength(4);
  });

  it('--all for a buffered format (table) honours --page-limit too', async () => {
    const calls = install(pages);
    const out = await run(['--all', '--page-limit', '2', '-o', 'table']);
    expect(calls).toHaveLength(2);
    expect(out).toContain('2b');
    expect(out).not.toContain('3a');
  });

  it('--page-delay waits between pages', async () => {
    install(pages);
    const started = Date.now();
    await run(['--all', '--page-delay', '60']);
    expect(Date.now() - started).toBeGreaterThanOrEqual(110);
  });

  it('caller query params and headers are kept on every page', async () => {
    const calls = install(pages);
    await listRequest(new GraphClient(fakeAuth()), '/items', { queryParams: { $select: 'id' }, headers: { Prefer: 'x' } }, { ...parse(['--all', '--top', '2']) }, () => {});
    expect(new URL(calls[0].url).searchParams.get('$select')).toBe('id');
    expect(calls.every((c) => c.headers.prefer === 'x')).toBe(true);
  });

  it('filterItems applies to every page in both modes', async () => {
    install(pages);
    const keepB = (item: unknown) => typeof item === 'object' && item !== null && 'id' in item && String(item.id).endsWith('b');
    expect(JSON.parse(await run(['--all'], undefined, { filterItems: keepB }))).toEqual([{ id: '1b' }, { id: '2b' }, { id: '3b' }]);
    const text = await run(['--all', '-o', 'text'], undefined, { filterItems: keepB });
    expect(text).toContain('3b');
    expect(text).not.toContain('2a');
    expect(JSON.parse(await run([], undefined, { filterItems: keepB }))).toEqual({ value: [{ id: '1b' }] });
  });
});
