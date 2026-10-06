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

const MESSAGES = {
  value: [
    {
      id: 'AAMk1',
      subject: 'Quarterly report',
      isRead: false,
      importance: 'normal',
      size: 1024,
      categories: ['Red', 'Blue'],
      from: { emailAddress: { name: 'Alice', address: 'alice@example.com' } },
      toRecipients: [{ emailAddress: { address: 'bob@example.com' } }, { emailAddress: { address: 'carol@example.com' } }],
      flag: { flagStatus: 'notFlagged' },
      bodyPreview: 'See attached',
      parentFolderId: null,
    },
    { id: 'AAMk2', subject: 'Lunch?', isRead: true, size: 12, categories: [], from: { emailAddress: { name: 'Bob', address: 'bob@example.com' } } },
  ],
};

/** The same payload as Graph returns it: annotations at the top level and on items. */
function withAnnotations<T>(data: T): T {
  const copy = JSON.parse(JSON.stringify(data));
  copy['@odata.context'] = 'https://graph.microsoft.com/v1.0/$metadata#users/messages';
  copy['@odata.nextLink'] = 'https://graph.microsoft.com/v1.0/me/messages?$skip=2';
  for (const item of copy.value) {
    item['@odata.etag'] = 'W/"CQAAABYAAAA"';
    item['@odata.type'] = '#microsoft.graph.message';
    if (item.from) item.from['@odata.type'] = '#microsoft.graph.recipient';
  }
  return copy;
}

const ALL_FORMATS: OutputFormat[] = ['json', 'table', 'csv', 'text', 'yaml'];

describe('GRAPH-4 rendered output (fixture snapshot, byte-identical across the move)', () => {
  it.each(ALL_FORMATS)('%s output of the message fixture is unchanged', (format) => {
    expect(formatOutput(MESSAGES, format)).toMatchSnapshot();
  });
});

describe('GRAPH-4 @odata annotations are stripped at output time', () => {
  it.each(ALL_FORMATS)('%s: annotated payload renders exactly like the clean one', (format) => {
    expect(formatOutput(withAnnotations(MESSAGES), format)).toBe(formatOutput(MESSAGES, format));
  });

  it('does not mutate the caller data (annotations stay usable after printing)', () => {
    const data = withAnnotations(MESSAGES);
    formatOutput(data, 'json');
    expect(data.value[0]).toHaveProperty(['@odata.etag'], 'W/"CQAAABYAAAA"');
  });

  it.each(['json', 'csv', 'table'] as OutputFormat[])('streamPages %s strips item annotations', async (format) => {
    async function* pages(): AsyncGenerator<unknown[], void, undefined> {
      yield withAnnotations(MESSAGES).value;
    }
    const chunks: string[] = [];
    await streamPages(pages(), format, (s) => chunks.push(s));
    expect(chunks.join('')).not.toContain('@odata');
  });
});

describe('CORE-4 streamPages JSON brackets', () => {
  it('writes nothing to stdout when the first page fails', async () => {
    async function* failing(): AsyncGenerator<unknown[], void, undefined> {
      throw new Error('network down');
    }
    const chunks: string[] = [];
    await expect(streamPages(failing(), 'json', (s) => chunks.push(s))).rejects.toThrow('network down');
    expect(chunks.join('')).toBe('');
  });

  it('emits [] for a result with only empty pages', async () => {
    async function* empty(): AsyncGenerator<unknown[], void, undefined> {
      yield [];
    }
    const chunks: string[] = [];
    await streamPages(empty(), 'json', (s) => chunks.push(s));
    expect(chunks.join('')).toBe('[]\n');
  });
});

describe('FMT-1 yaml scalars and keys', () => {
  const yaml = (v: unknown) => formatOutput(v, 'yaml');

  it.each([
    ['[EXTERNAL] Hi', '"[EXTERNAL] Hi"'],
    ['- dash', '"- dash"'],
    ['~', '"~"'],
    ['%done', '"%done"'],
    ['@mention', '"@mention"'],
    ['*star', '"*star"'],
    ['&anchor', '"&anchor"'],
    ['!tag', '"!tag"'],
    ['|pipe', '"|pipe"'],
    ['>fold', '">fold"'],
    ["'single", '"\'single"'],
    ['`tick', '"`tick"'],
    ['{brace', '"{brace"'],
    [',comma', '",comma"'],
    ['?q', '"?q"'],
    ['True', '"True"'],
    ['NULL', '"NULL"'],
    ['Yes', '"Yes"'],
    [' lead', '" lead"'],
    ['trail ', '"trail "'],
    ['a\r\nb\tc', '"a\\r\\nb\\tc"'],
    ['back\\slash "q"', '"back\\\\slash \\"q\\""'],
    ['plain words', 'plain words'],
    ['Lunch?', 'Lunch?'],
  ])('value %j renders as %s', (input, expected) => {
    expect(yaml({ v: input })).toBe(`v: ${expected}\n`);
    expect(yaml([input])).toBe(`- ${expected}\n`);
  });

  it.each([
    '2026-10-06', '2026-1-6', '0b101', '-0b1_0', '1_000', '1_000.5', '0777', '0x_1F', '1.2.3', '5.', '=', '<<',
    '1:20', '-1:20:30.5', '.inf', '-.Inf', '.NaN', '~', '0o17', '0x1F', '1e5', '+12', 'off', 'Y',
  ])('YAML 1.1 non-string %j is quoted (as value and key)', (input) => {
    expect(yaml({ v: input })).toBe(`v: "${input}"\n`);
    expect(yaml({ [input]: 1 })).toBe(`"${input}": 1\n`);
  });

  it.each(['v1.2', '2026-10', '12345-6789', 'a=b', '<<a', '_1', '1.2e', '0b2', 'x1_000', '2026-10-06 notes'])(
    'look-alike %j stays plain', (input) => {
      expect(yaml({ v: input })).toBe(`v: ${input}\n`);
    });

  it('quotes keys the same way (driveItem fixture)', () => {
    const item = {
      '@microsoft.graph.downloadUrl': 'https://x.example/d?a=1',
      name: '- notes.md',
      file: { mimeType: 'text/markdown' },
    };
    expect(yaml(item)).toBe([
      '"@microsoft.graph.downloadUrl": "https://x.example/d?a=1"',
      'name: "- notes.md"',
      'file:',
      '  mimeType: text/markdown',
      '',
    ].join('\n'));
    expect(yaml({ value: [item] })).toBe([
      'value:',
      '  - "@microsoft.graph.downloadUrl": "https://x.example/d?a=1"',
      '    name: "- notes.md"',
      '    file: {"mimeType":"text/markdown"}',
      '',
    ].join('\n'));
  });

  it('a multi-line value stays on one line', () => {
    const out = yaml({ body: 'line1\nline2\r\n', count: 1 });
    expect(out.split('\n')).toEqual(['body: "line1\\nline2\\r\\n"', 'count: 1', '']);
  });
});

describe('FMT-1 table and csv', () => {
  it('table rows do not break on CR/LF or tabs in values', () => {
    const out = formatOutput({ value: [{ id: '1', subject: 'a\r\nb\tc' }, { id: '2', subject: 'x' }] }, 'table');
    const lines = out.split('\n');
    expect(lines).toHaveLength(4);
    expect(lines[2]).toMatch(/^1\s+a b c\s*$/);
  });

  it('csv keeps numeric literals unmangled', () => {
    const out = formatOutput({ value: [{ n: -5, s: '+3.5e2' }] }, 'csv');
    expect(out).toBe('n,s\n-5,+3.5e2');
  });

  it('csv still guards formulas', () => {
    const out = formatOutput({ value: [{ f: '=1+2', g: '-cmd', h: '@SUM(A1)' }] }, 'csv');
    expect(out).toBe('f,g,h\n"\t=1+2","\t-cmd","\t@SUM(A1)"');
  });

  it('csv quotes a bare CR', () => {
    expect(formatOutput({ value: [{ s: 'a\rb' }] }, 'csv')).toBe('s\n"a\rb"');
  });
});
