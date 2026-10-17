import { CliError } from './errors.js';

/**
 * RFC 4180 CSV reading, in-repo (no dependency).
 *
 * - Fields separated by commas, records by CRLF or LF (a lone CR also ends a record).
 * - Quoted fields may contain commas, CR/LF and `""` (an escaped quote).
 * - A leading UTF-8 byte-order mark is dropped; blank lines are skipped.
 * - Lenient where RFC 4180 is silent: a quote in the middle of an unquoted field is kept as text,
 *   and text after a closing quote is appended to the field.
 *
 * `CsvParser` is incremental: feed it chunks with `push()` and it returns only the records completed so
 * far, holding at most one partial record. `csvRecords()` / `csvObjects()` wrap it over a byte or string
 * stream, so a 100+ MB export is never held in memory at once.
 */
export class CsvParser {
  private field = '';
  private record: string[] = [];
  /** True until the current field has consumed any character (an opening quote is only legal here). */
  private atFieldStart = true;
  private inQuotes = false;
  /** A quote closed a quoted field in the previous step; the next char decides `""` vs end of quotes. */
  private afterQuote = false;
  /** The previous chunk ended in CR outside quotes; a leading LF in the next chunk belongs to it. */
  private pendingCR = false;
  private first = true;
  private line = 1;

  /** Feed the next chunk; returns the records it completed. */
  push(chunk: string): string[][] {
    const out: string[][] = [];
    let i = 0;
    if (this.first && chunk.length > 0) {
      this.first = false;
      if (chunk.charCodeAt(0) === 0xfeff) i = 1;
    }
    const n = chunk.length;
    while (i < n) {
      const c = chunk[i];
      if (this.pendingCR) {
        this.pendingCR = false;
        if (c === '\n') { i++; continue; }
      }
      if (this.inQuotes) {
        const j = chunk.indexOf('"', i);
        const end = j === -1 ? n : j;
        const part = chunk.slice(i, end);
        this.field += part;
        for (let k = 0; k < part.length; k++) if (part.charCodeAt(k) === 10) this.line++;
        if (j === -1) break;
        this.inQuotes = false;
        this.afterQuote = true;
        i = j + 1;
        continue;
      }
      if (this.afterQuote) {
        this.afterQuote = false;
        if (c === '"') { this.field += '"'; this.inQuotes = true; i++; continue; }
      }
      if (c === ',') {
        this.endField();
        i++;
      } else if (c === '\n' || c === '\r') {
        this.endRecord(out);
        this.line++;
        i++;
        if (c === '\r') {
          if (i < n) { if (chunk[i] === '\n') i++; } else this.pendingCR = true;
        }
      } else if (c === '"' && this.atFieldStart) {
        this.inQuotes = true;
        this.atFieldStart = false;
        i++;
      } else {
        let j = i + 1;
        while (j < n) {
          const d = chunk.charCodeAt(j);
          if (d === 44 /* , */ || d === 10 || d === 13 || d === 34 /* " */) break;
          j++;
        }
        this.field += chunk.slice(i, j);
        this.atFieldStart = false;
        i = j;
      }
    }
    return out;
  }

  /** Signal end of input; returns the final record, if any. Throws on an unterminated quoted field. */
  end(): string[][] {
    if (this.inQuotes) {
      throw new CliError(`CSV: a quoted field is never closed (record ending at line ${this.line}).`, { code: 'bad_csv' });
    }
    this.afterQuote = false;
    const out: string[][] = [];
    this.endRecord(out);
    return out;
  }

  private endField(): void {
    this.record.push(this.field);
    this.field = '';
    this.atFieldStart = true;
  }

  private endRecord(out: string[][]): void {
    const blank = this.record.length === 0 && this.atFieldStart;
    if (!blank) {
      this.endField();
      out.push(this.record);
    }
    this.record = [];
    this.field = '';
    this.atFieldStart = true;
  }
}

/** Parse a whole CSV text into records (arrays of fields). */
export function parseCsv(text: string): string[][] {
  const p = new CsvParser();
  const out = p.push(text);
  for (const r of p.end()) out.push(r);
  return out;
}

/** Turn records into objects keyed by the header (first) record; missing trailing fields become ''. */
function toObject(header: string[], rec: string[]): Record<string, string> {
  const o: Record<string, string> = {};
  for (let k = 0; k < header.length; k++) o[header[k]] = rec[k] ?? '';
  return o;
}

/** Parse a whole CSV text with a header row into one object per record. */
export function parseCsvObjects(text: string): Record<string, string>[] {
  const [header, ...rows] = parseCsv(text);
  return header ? rows.map(r => toObject(header, r)) : [];
}

/** Stream records from chunks of UTF-8 bytes or text (e.g. a fs.ReadStream or a fetch body). */
export async function* csvRecords(source: AsyncIterable<Uint8Array | string>): AsyncGenerator<string[]> {
  const parser = new CsvParser();
  const decoder = new TextDecoder('utf-8');
  for await (const chunk of source) {
    const text = typeof chunk === 'string' ? chunk : decoder.decode(chunk, { stream: true });
    yield* parser.push(text);
  }
  const tail = decoder.decode();
  if (tail) yield* parser.push(tail);
  yield* parser.end();
}

/** Stream objects keyed by the header record. */
export async function* csvObjects(source: AsyncIterable<Uint8Array | string>): AsyncGenerator<Record<string, string>> {
  let header: string[] | undefined;
  for await (const rec of csvRecords(source)) {
    if (!header) { header = rec; continue; }
    yield toObject(header, rec);
  }
}
