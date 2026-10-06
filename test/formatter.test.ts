import { describe, it, expect, vi } from 'vitest';
import { formatOutput, streamPages, type OutputFormat } from '../src/formatter.js';

describe('formatOutput', () => {
  describe('json', () => {
    it('formats objects as pretty JSON', () => {
      const result = formatOutput({ a: 1, b: 'two' }, 'json');
      expect(JSON.parse(result)).toEqual({ a: 1, b: 'two' });
    });

    it('handles null and undefined in objects', () => {
      const result = formatOutput({ a: null, b: undefined }, 'json');
      const parsed = JSON.parse(result);
      expect(parsed.a).toBeNull();
    });

    it('formats arrays', () => {
      const result = formatOutput([1, 2, 3], 'json');
      expect(JSON.parse(result)).toEqual([1, 2, 3]);
    });
  });

  describe('table', () => {
    it('formats value-wrapped arrays as table', () => {
      const data = { value: [{ id: '1', name: 'Alice' }, { id: '2', name: 'Bob' }] };
      const result = formatOutput(data, 'table');
      expect(result).toContain('id');
      expect(result).toContain('name');
      expect(result).toContain('Alice');
      expect(result).toContain('Bob');
      // Should have header, separator, and 2 data rows
      const lines = result.split('\n');
      expect(lines.length).toBe(4);
    });

    it('returns (no results) for empty value array', () => {
      const result = formatOutput({ value: [] }, 'table');
      expect(result).toBe('(no results)');
    });

    it('truncates long values to 50 chars', () => {
      const longVal = 'x'.repeat(100);
      const data = { value: [{ text: longVal }] };
      const result = formatOutput(data, 'table');
      const dataLine = result.split('\n')[2];
      expect(dataLine.trim().length).toBeLessThanOrEqual(50);
    });

    it('flattens nested objects to dot notation', () => {
      const data = { value: [{ from: { emailAddress: { address: 'a@b.com' } } }] };
      const result = formatOutput(data, 'table');
      expect(result).toContain('from.emailAddress.address');
      expect(result).toContain('a@b.com');
    });
  });

  describe('csv', () => {
    it('formats as CSV with header', () => {
      const data = { value: [{ id: '1', name: 'Alice' }] };
      const result = formatOutput(data, 'csv');
      const lines = result.split('\n');
      expect(lines[0]).toBe('id,name');
      expect(lines[1]).toBe('1,Alice');
    });

    it('escapes commas in values', () => {
      const data = { value: [{ text: 'hello, world' }] };
      const result = formatOutput(data, 'csv');
      expect(result).toContain('"');
    });

    it('escapes double quotes in values', () => {
      const data = { value: [{ text: 'say "hi"' }] };
      const result = formatOutput(data, 'csv');
      expect(result).toContain('""hi""');
    });

    it('prevents formula injection', () => {
      const data = { value: [{ text: '=CMD("calc")' }] };
      const result = formatOutput(data, 'csv');
      // Should be quoted and prefixed with tab
      expect(result).toContain('\t=CMD');
    });

    it('prevents +, -, @ formula injection', () => {
      for (const prefix of ['+', '-', '@']) {
        const data = { value: [{ text: `${prefix}danger` }] };
        const result = formatOutput(data, 'csv');
        expect(result).toContain(`\t${prefix}danger`);
      }
    });

    it('returns empty string for no items', () => {
      expect(formatOutput({ value: [] }, 'csv')).toBe('');
    });
  });

  describe('text', () => {
    it('formats numbered items', () => {
      const data = { value: [{ id: '1' }, { id: '2' }] };
      const result = formatOutput(data, 'text');
      expect(result).toContain('[1]');
      expect(result).toContain('[2]');
    });

    it('formats single object as key-value pairs', () => {
      const data = { name: 'Test', count: 5 };
      const result = formatOutput(data, 'text');
      // extractItems wraps single non-{value:[]} in array
      expect(result).toContain('name');
    });
  });

  describe('yaml', () => {
    it('quotes boolean-like strings', () => {
      const data = { value: 'true' };
      const result = formatOutput(data, 'yaml');
      expect(result).toContain('"true"');
    });

    it('quotes numeric-like strings', () => {
      const data = { value: '123' };
      const result = formatOutput(data, 'yaml');
      expect(result).toContain('"123"');
    });

    it('quotes yes/no/on/off strings', () => {
      for (const val of ['yes', 'no', 'on', 'off']) {
        const data = { value: val };
        const result = formatOutput(data, 'yaml');
        expect(result).toContain(`"${val}"`);
      }
    });

    it('does not quote normal strings', () => {
      const data = { value: 'hello world' };
      const result = formatOutput(data, 'yaml');
      expect(result).toContain('value: hello world');
      expect(result).not.toContain('"hello world"');
    });

    it('quotes strings with colons', () => {
      const data = { value: 'key: val' };
      const result = formatOutput(data, 'yaml');
      expect(result).toContain('"key: val"');
    });

    it('renders empty array', () => {
      const data = { items: [] };
      const result = formatOutput(data, 'yaml');
      expect(result).toContain('[]\n');
    });

    it('renders empty object', () => {
      const result = formatOutput({}, 'yaml');
      expect(result).toContain('{}\n');
    });
  });

  describe('default', () => {
    it('falls back to JSON for unknown format', () => {
      const data = { a: 1 };
      const result = formatOutput(data, 'unknown' as OutputFormat);
      expect(JSON.parse(result)).toEqual({ a: 1 });
    });
  });
});

describe('streamPages', () => {
  async function* makePages(pages: unknown[][]): AsyncGenerator<unknown[], void, undefined> {
    for (const page of pages) {
      yield page;
    }
  }

  it('streams JSON as valid array', async () => {
    const chunks: string[] = [];
    const pages = makePages([[{ id: 1 }, { id: 2 }], [{ id: 3 }]]);
    await streamPages(pages, 'json', (s) => chunks.push(s));
    const full = chunks.join('');
    const parsed = JSON.parse(full);
    expect(parsed).toHaveLength(3);
    expect(parsed[0].id).toBe(1);
    expect(parsed[2].id).toBe(3);
  });

  it('streams CSV with single header', async () => {
    const chunks: string[] = [];
    const pages = makePages([[{ a: '1', b: '2' }], [{ a: '3', b: '4' }]]);
    await streamPages(pages, 'csv', (s) => chunks.push(s));
    const lines = chunks.join('').trim().split('\n');
    expect(lines[0]).toBe('a,b');
    expect(lines[1]).toBe('1,2');
    expect(lines[2]).toBe('3,4');
    // Only one header line
    expect(lines.filter(l => l === 'a,b').length).toBe(1);
  });

  it('handles empty pages', async () => {
    const chunks: string[] = [];
    const pages = makePages([]);
    await streamPages(pages, 'json', (s) => chunks.push(s));
    const full = chunks.join('');
    expect(JSON.parse(full)).toEqual([]);
  });

  it('handles first page empty, second page has items', async () => {
    const chunks: string[] = [];
    const pages = makePages([[], [{ id: 1 }]]);
    await streamPages(pages, 'json', (s) => chunks.push(s));
    const parsed = JSON.parse(chunks.join(''));
    expect(parsed).toHaveLength(1);
    expect(parsed[0].id).toBe(1);
  });

  it('CSV handles missing columns gracefully', async () => {
    const chunks: string[] = [];
    // First page has {a,b}, second page has {a} only — b should be empty
    const pages = makePages([[{ a: '1', b: '2' }], [{ a: '3' }]]);
    await streamPages(pages, 'csv', (s) => chunks.push(s));
    const lines = chunks.join('').trim().split('\n');
    expect(lines[0]).toBe('a,b');
    expect(lines[2]).toBe('3,'); // missing b becomes empty
  });

  it('CSV warns on new columns in later pages', async () => {
    const chunks: string[] = [];
    const errSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    // First page {a}, second page {a, b} — b is new
    const pages = makePages([[{ a: '1' }], [{ a: '2', b: '3' }]]);
    await streamPages(pages, 'csv', (s) => chunks.push(s));
    expect(errSpy).toHaveBeenCalledWith(expect.stringContaining('Column "b"'));
    errSpy.mockRestore();
  });

  it('streams single-item pages in JSON', async () => {
    const chunks: string[] = [];
    const pages = makePages([[{ x: 1 }], [{ x: 2 }], [{ x: 3 }]]);
    await streamPages(pages, 'json', (s) => chunks.push(s));
    const parsed = JSON.parse(chunks.join(''));
    expect(parsed).toHaveLength(3);
  });

  it('falls back to buffered formatOutput for table format', async () => {
    const chunks: string[] = [];
    const pages = makePages([[{ id: '1', name: 'Alice' }], [{ id: '2', name: 'Bob' }]]);
    await streamPages(pages, 'table', (s) => chunks.push(s));
    const output = chunks.join('');
    expect(output).toContain('id');
    expect(output).toContain('Alice');
    expect(output).toContain('Bob');
  });

  it('handles primitive items in JSON streaming', async () => {
    const chunks: string[] = [];
    const pages = makePages([['hello', 'world'], [42]]);
    await streamPages(pages, 'json', (s) => chunks.push(s));
    const parsed = JSON.parse(chunks.join(''));
    expect(parsed).toEqual(['hello', 'world', 42]);
  });
});
