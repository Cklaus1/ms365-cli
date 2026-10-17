import { describe, it, expect } from 'vitest';
import { CsvParser, parseCsv, parseCsvObjects, csvRecords, csvObjects } from '../src/csv.js';

/** Feed `text` to a fresh parser in chunks of `size` characters. */
function chunked(text: string, size: number): string[][] {
  const p = new CsvParser();
  const out: string[][] = [];
  for (let i = 0; i < text.length; i += size) out.push(...p.push(text.slice(i, i + size)));
  out.push(...p.end());
  return out;
}

async function collect<T>(gen: AsyncGenerator<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const x of gen) out.push(x);
  return out;
}

const TRICKY = '﻿Name,Assigned Products,Note\r\n'
  + 'a@x.com,"OFFICE 365 E3,MICROSOFT 365 E5","He said ""hi""\r\nthen left"\r\n'
  + 'b@x.com,,"line1\nline2"\n'
  + '"",plain,"x"\r\n';

const EXPECTED = [
  ['Name', 'Assigned Products', 'Note'],
  ['a@x.com', 'OFFICE 365 E3,MICROSOFT 365 E5', 'He said "hi"\r\nthen left'],
  ['b@x.com', '', 'line1\nline2'],
  ['', 'plain', 'x'],
];

describe('parseCsv (RFC 4180)', () => {
  it('handles quoted commas, "" escapes, CRLF/LF inside quotes and a UTF-8 BOM', () => {
    expect(parseCsv(TRICKY)).toEqual(EXPECTED);
  });

  it('parses identically at every chunk boundary (streaming)', () => {
    for (const size of [1, 2, 3, 5, 7, 13, 64]) expect(chunked(TRICKY, size), `chunk ${size}`).toEqual(EXPECTED);
  });

  it('keeps empty fields, skips blank lines, accepts a missing final newline and a lone CR', () => {
    expect(parseCsv('a,,c\n\n,\r\nx')).toEqual([['a', '', 'c'], ['', ''], ['x']]);
    expect(parseCsv('a,b\rc,d\r')).toEqual([['a', 'b'], ['c', 'd']]);
    expect(parseCsv('')).toEqual([]);
    expect(parseCsv('﻿')).toEqual([]);
  });

  it('is lenient about a stray quote inside an unquoted field', () => {
    expect(parseCsv('5" disk,ok\n')).toEqual([['5" disk', 'ok']]);
  });

  it('rejects an unterminated quoted field', () => {
    expect(() => parseCsv('a,"open\n')).toThrow(/never closed/);
  });

  it('push() only returns completed records and holds the partial one', () => {
    const p = new CsvParser();
    expect(p.push('a,b\nc,"d')).toEqual([['a', 'b']]);
    expect(p.push('\n",e')).toEqual([]);
    expect(p.push('\n')).toEqual([['c', 'd\n', 'e']]);
    expect(p.end()).toEqual([]);
  });

  it('parseCsvObjects keys records by the header and fills missing fields', () => {
    expect(parseCsvObjects('User Principal Name,Send Count\nA@x.com,3\nb@x.com\n')).toEqual([
      { 'User Principal Name': 'A@x.com', 'Send Count': '3' },
      { 'User Principal Name': 'b@x.com', 'Send Count': '' },
    ]);
    expect(parseCsvObjects('')).toEqual([]);
  });
});

describe('streaming over bytes', () => {
  it('decodes UTF-8 split across chunks (incl. multi-byte chars and the BOM)', async () => {
    const bytes = new TextEncoder().encode('﻿dept,name\n"Ventes, Europe","Zoë ✓"\n');
    const chunks = Array.from(bytes, b => Uint8Array.of(b)); // one byte per chunk
    async function* src() { yield* chunks; }
    expect(await collect(csvRecords(src()))).toEqual([['dept', 'name'], ['Ventes, Europe', 'Zoë ✓']]);
  });

  it('csvObjects streams a large input without returning it in one piece', async () => {
    const rows = 20000;
    async function* src() {
      yield 'id,products\r\n';
      for (let i = 0; i < rows; i++) yield `${i},"A,B"\r\n`;
    }
    let count = 0;
    let last: Record<string, string> | undefined;
    for await (const o of csvObjects(src())) { count++; last = o; }
    expect(count).toBe(rows);
    expect(last).toEqual({ id: String(rows - 1), products: 'A,B' });
  });
});
