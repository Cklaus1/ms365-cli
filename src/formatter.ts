import type { GraphClient, GraphRequestOptions, PagingState } from './graph-client.js';
import { CliError, ExitCode } from './errors.js';
import { defaultColumns } from './columns.js';
import { info } from './utils.js';

export type OutputFormat = 'json' | 'ndjson' | 'table' | 'csv' | 'text' | 'yaml';

// ── Global output controls (set from --fields / --cursor in the preAction hook) ──

let selectedFields: string[] | null = null;
/** Limit output to these (dot-path) fields, e.g. ["id", "from.emailAddress.address"]. */
export function setFields(fields: string[] | null): void { selectedFields = fields && fields.length ? fields : null; }

let pageCursor: string | null = null;
/** Resume a list from an opaque cursor printed by a previous page. */
export function setCursor(cursor: string | null): void { pageCursor = cursor || null; }

/** Opaque, shell-safe cursor for a Graph nextLink. */
export function encodeCursor(nextLink: string): string {
  return Buffer.from(nextLink).toString('base64url');
}

/** Version segment and path of a Graph URL, percent-decoded and case-folded (E3), for comparing a cursor with a command. */
function graphPath(url: URL): string {
  let path = url.pathname;
  try { path = decodeURIComponent(path); } catch { /* a stray % stays as written */ }
  return path.replace(/\/+$/, '').toLowerCase();
}

/**
 * The Graph URL in a --cursor. With `expected` (the command's own endpoint URL) the cursor must name the same API
 * version and path — only the query may differ — else exit 2 `cursor_mismatch` before any request (AH-1(7)), so a
 * cursor can never carry a command to another resource.
 */
export function decodeCursor(cursor: string, expected?: string): string {
  const url = Buffer.from(cursor, 'base64url').toString('utf8');
  let parsed: URL | undefined;
  try { parsed = new URL(url); } catch { /* invalid */ }
  if (parsed?.hostname !== 'graph.microsoft.com') {
    throw new CliError('Invalid --cursor value.', { code: 'usage', exitCode: ExitCode.USAGE, fix: 'pass the nextCursor printed by the previous page' });
  }
  if (expected !== undefined && graphPath(parsed) !== graphPath(new URL(expected))) {
    throw new CliError(`--cursor belongs to ${parsed.pathname}, not to this command (${new URL(expected).pathname}).`, {
      code: 'cursor_mismatch', exitCode: ExitCode.USAGE, fix: 'pass the nextCursor printed by the previous page of this same command and --api-version',
    });
  }
  return url;
}

function getPath(obj: unknown, dotted: string): unknown {
  let cur: unknown = obj;
  for (const part of dotted.split('.')) {
    if (cur === null || typeof cur !== 'object') return undefined;
    cur = (cur as Record<string, unknown>)[part];
  }
  return cur;
}

function setPath(obj: Record<string, unknown>, dotted: string, value: unknown): void {
  const parts = dotted.split('.');
  let cur = obj;
  for (const part of parts.slice(0, -1)) {
    if (typeof cur[part] !== 'object' || cur[part] === null) cur[part] = {};
    cur = cur[part] as Record<string, unknown>;
  }
  cur[parts[parts.length - 1]] = value;
}

/** Keep only the selected fields of one item (nested structure preserved). */
export function projectItem(item: unknown, fields: string[] | null = selectedFields): unknown {
  if (!fields || item === null || typeof item !== 'object' || Array.isArray(item)) return item;
  const out: Record<string, unknown> = {};
  for (const f of fields) {
    const v = getPath(item, f);
    if (v !== undefined) setPath(out, f, v);
  }
  return out;
}

/** Apply --fields to a response body: each item of a collection, or the single object. */
function project(data: unknown): unknown {
  if (!selectedFields) return data;
  if (Array.isArray(data)) return data.map(i => projectItem(i));
  if (data && typeof data === 'object' && Array.isArray((data as Record<string, unknown>).value)) {
    return { ...(data as Record<string, unknown>), value: ((data as Record<string, unknown>).value as unknown[]).map(i => projectItem(i)) };
  }
  return projectItem(data);
}

/**
 * AH-3: whether a list holds everything Graph has. Always in band: JSON/YAML envelope keys, an NDJSON trailer line
 * `{"_meta": …}`. `next_cursor`: one page of several (pass `nextCursor` to `--cursor`); `page_limit`: `--all` stopped
 * at `--page-limit`; `search_cap`: mail `$search` returned its 250-item maximum.
 */
export interface Completeness {
  complete: boolean;
  incompleteReason?: 'page_limit' | 'search_cap' | 'next_cursor';
  nextCursor?: string;
}

/**
 * Stream formatted output page-by-page from an async generator.
 * JSON streams as a valid array; CSV streams rows after a single header;
 * other formats fall back to buffering all items. `meta`, read after the last page, closes the JSON
 * wrapper (`],"complete":…}`) and is the NDJSON trailer line.
 */
export async function streamPages(
  pages: AsyncGenerator<unknown[], void, undefined>,
  format: OutputFormat,
  write: (chunk: string) => void = (s) => process.stdout.write(s),
  options: { jsonWrapper?: boolean; meta?: () => Completeness } = {},
): Promise<void> {
  if (format === 'ndjson') {
    for await (const page of pages) {
      for (const item of page) write(JSON.stringify(projectItem(item)) + '\n');
    }
    if (options.meta) write(JSON.stringify({ _meta: options.meta() }) + '\n');
    return;
  }

  if (format === 'json') {
    // jsonWrapper keeps the same {"value": [...]} shape as a buffered --all response
    const indent = options.jsonWrapper ? '    ' : '';
    let first = true;
    write(options.jsonWrapper ? '{\n  "value": [\n' : '[\n');
    for await (const page of pages) {
      for (const item of page) {
        if (!first) write(',\n');
        write(JSON.stringify(projectItem(item), null, 2).replace(/^/gm, indent));
        first = false;
      }
    }
    if (!options.jsonWrapper) {
      write('\n]\n');
      return;
    }
    const meta = options.meta ? Object.entries(options.meta()).map(([k, v]) => `,\n  ${JSON.stringify(k)}: ${JSON.stringify(v)}`).join('') : '';
    write(`\n  ]${meta}\n}\n`);
    return;
  }

  if (format === 'csv') {
    let headerWritten = false;
    const columnSet = new Set<string>();
    let columns: string[] = [];
    for await (const page of pages) {
      const flatItems = page.map((raw) => {
        const item = projectItem(raw);
        if (typeof item !== 'object' || item === null) return { value: String(item) };
        return flattenObject(item as Record<string, unknown>);
      });
      if (!headerWritten && flatItems.length > 0) {
        for (const item of flatItems) {
          for (const key of Object.keys(item)) columnSet.add(key);
        }
        columns = Array.from(columnSet);
        write(columns.map(csvEscape).join(',') + '\n');
        headerWritten = true;
      } else {
        // Track new columns that weren't in the header (warn but don't break)
        for (const item of flatItems) {
          for (const key of Object.keys(item)) {
            if (!columnSet.has(key)) {
              columnSet.add(key);
              console.error(`Warning: Column "${key}" appeared after CSV header was written and will be omitted.`);
            }
          }
        }
      }
      for (const item of flatItems) {
        write(columns.map(col => csvEscape(item[col] ?? '')).join(',') + '\n');
      }
    }
    return;
  }

  // table/text/yaml: must buffer all items (need full data for column widths / structure)
  const allItems: unknown[] = [];
  for await (const page of pages) {
    allItems.push(...page);
  }
  write(formatOutput({ value: allItems }, format) + '\n');
}

export function formatOutput(rawData: unknown, format: OutputFormat): string {
  const data = project(rawData);
  switch (format) {
    case 'json':
      return JSON.stringify(data, null, 2);
    case 'ndjson':
      return extractItems(data).map(i => JSON.stringify(i)).join('\n');
    case 'table':
      return formatTable(data);
    case 'csv':
      return formatCsv(data);
    case 'text':
      return formatText(data);
    case 'yaml':
      return formatYaml(data);
    default:
      return JSON.stringify(data, null, 2);
  }
}

/**
 * Flatten a nested object into dot-notation keys.
 * { from: { emailAddress: { address: "a@b" } } } → { "from.emailAddress.address": "a@b" }
 */
function flattenObject(obj: Record<string, unknown>, prefix = ''): Record<string, string> {
  const result: Record<string, string> = {};
  for (const [key, val] of Object.entries(obj)) {
    if (key.startsWith('@odata.')) continue; // metadata (type/etag) stays in JSON output only
    const fullKey = prefix ? `${prefix}.${key}` : key;
    if (val === null || val === undefined) {
      result[fullKey] = '';
    } else if (Array.isArray(val)) {
      // Flatten simple arrays to comma-separated, skip complex ones
      if (val.length > 0 && typeof val[0] === 'object') {
        // For arrays of objects (like toRecipients), extract a summary
        const summaries = val.map((item) => {
          if (typeof item === 'object' && item !== null) {
            const flat = flattenObject(item as Record<string, unknown>);
            // Pick the most useful value (address, name, displayName, etc.)
            return flat['emailAddress.address'] || flat['address'] || flat['displayName'] || flat['name'] || Object.values(flat)[0] || '';
          }
          return String(item);
        });
        result[fullKey] = summaries.join(', ');
      } else {
        result[fullKey] = val.join(', ');
      }
    } else if (typeof val === 'object') {
      Object.assign(result, flattenObject(val as Record<string, unknown>, fullKey));
    } else {
      result[fullKey] = String(val);
    }
  }
  return result;
}

/**
 * The detected resource's default columns that some item has, unless --fields is set or
 * MS365_CLI_TABLE_ALL_COLUMNS asks for everything; otherwise every column.
 */
function tableColumns(items: unknown[], allColumns: string[]): string[] {
  const all = process.env.MS365_CLI_TABLE_ALL_COLUMNS;
  if (selectedFields || all === '1' || all === 'true') return allColumns;
  const present = new Set(allColumns);
  const defaults = defaultColumns(items)?.filter(c => present.has(c));
  return defaults?.length ? defaults : allColumns;
}

function formatTable(data: unknown): string {
  const items = extractItems(data);
  if (items.length === 0) return '(no results)';

  // Flatten all items
  const flatItems = items.map((item) => {
    if (typeof item !== 'object' || item === null) return { value: String(item) };
    return flattenObject(item as Record<string, unknown>);
  });

  // Collect all keys preserving insertion order
  const keySet = new Set<string>();
  for (const item of flatItems) {
    for (const key of Object.keys(item)) {
      keySet.add(key);
    }
  }
  const columns = tableColumns(items, Array.from(keySet));
  if (columns.length === 0) return JSON.stringify(data, null, 2);
  // F9: the defaults leave out id and the long tail; say how to get them (stderr, so stdout stays the table).
  if (columns.length < keySet.size) info('(showing default columns; use --fields id,... or -o json for all fields)');

  // Calculate column widths (capped at 50)
  const widths: Record<string, number> = {};
  for (const col of columns) {
    widths[col] = col.length;
  }
  for (const item of flatItems) {
    for (const col of columns) {
      const val = item[col] ?? '';
      widths[col] = Math.min(Math.max(widths[col], val.length), 50);
    }
  }

  const header = columns.map(c => c.padEnd(widths[c])).join('  ');
  const separator = columns.map(c => '-'.repeat(widths[c])).join('  ');

  const rows = flatItems.map(item => {
    return columns.map(col => {
      const val = item[col] ?? '';
      return val.substring(0, 50).padEnd(widths[col]);
    }).join('  ');
  });

  return [header, separator, ...rows].join('\n');
}

function formatCsv(data: unknown): string {
  const items = extractItems(data);
  if (items.length === 0) return '';

  const flatItems = items.map((item) => {
    if (typeof item !== 'object' || item === null) return { value: String(item) };
    return flattenObject(item as Record<string, unknown>);
  });

  const keySet = new Set<string>();
  for (const item of flatItems) {
    for (const key of Object.keys(item)) {
      keySet.add(key);
    }
  }
  const columns = Array.from(keySet);
  const header = columns.map(csvEscape).join(',');
  const rows = flatItems.map(item => {
    return columns.map(col => csvEscape(item[col] ?? '')).join(',');
  });

  return [header, ...rows].join('\n');
}

function formatText(data: unknown): string {
  const items = extractItems(data);
  if (items.length === 0) {
    if (typeof data === 'object' && data !== null) {
      const flat = flattenObject(data as Record<string, unknown>);
      return Object.entries(flat)
        .map(([k, v]) => `${k}: ${v}`)
        .join('\n');
    }
    return String(data);
  }

  return items.map((item, i) => {
    if (typeof item !== 'object' || item === null) return String(item);
    const flat = flattenObject(item as Record<string, unknown>);
    const entries = Object.entries(flat)
      .map(([k, v]) => `  ${k}: ${v}`)
      .join('\n');
    return `[${i + 1}]\n${entries}`;
  }).join('\n\n');
}

function formatYaml(data: unknown): string {
  return toYaml(data, 0);
}

function toYaml(data: unknown, indent: number): string {
  const pad = '  '.repeat(indent);
  if (data === null || data === undefined) return `${pad}null\n`;
  if (typeof data === 'boolean') return `${pad}${data}\n`;
  if (typeof data === 'number') return `${pad}${data}\n`;
  if (typeof data === 'string') {
    if (
      data.includes('\n') || data.includes('"') || data.includes(':') || data.includes('#') ||
      data === '' || data === 'true' || data === 'false' || data === 'null' ||
      data === 'yes' || data === 'no' || data === 'on' || data === 'off' ||
      /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(data)
    ) {
      return `${pad}"${data.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n')}"\n`;
    }
    return `${pad}${data}\n`;
  }
  if (Array.isArray(data)) {
    if (data.length === 0) return `${pad}[]\n`;
    let out = '';
    for (const item of data) {
      if (typeof item === 'object' && item !== null && !Array.isArray(item)) {
        const entries = Object.entries(item as Record<string, unknown>);
        if (entries.length > 0) {
          const [firstKey, firstVal] = entries[0];
          out += `${pad}- ${firstKey}: ${inlineValue(firstVal)}\n`;
          for (let j = 1; j < entries.length; j++) {
            out += `${pad}  ${entries[j][0]}: ${inlineValue(entries[j][1])}\n`;
          }
          continue;
        }
      }
      out += `${pad}- ${inlineValue(item)}\n`;
    }
    return out;
  }
  if (typeof data === 'object') {
    const entries = Object.entries(data as Record<string, unknown>);
    if (entries.length === 0) return `${pad}{}\n`;
    let out = '';
    for (const [key, val] of entries) {
      if (typeof val === 'object' && val !== null && !Array.isArray(val)) {
        out += `${pad}${key}:\n${toYaml(val, indent + 1)}`;
      } else if (Array.isArray(val)) {
        out += `${pad}${key}:\n${toYaml(val, indent + 1)}`;
      } else {
        out += `${pad}${key}: ${inlineValue(val)}\n`;
      }
    }
    return out;
  }
  return `${pad}${String(data)}\n`;
}

function inlineValue(val: unknown): string {
  if (val === null || val === undefined) return 'null';
  if (typeof val === 'boolean' || typeof val === 'number') return String(val);
  if (typeof val === 'string') {
    // Quote strings that YAML would misinterpret as non-string types
    if (
      val.includes('\n') || val.includes('"') || val.includes(':') || val.includes('#') ||
      val === '' || val === 'true' || val === 'false' || val === 'null' ||
      val === 'yes' || val === 'no' || val === 'on' || val === 'off' ||
      /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(val)
    ) {
      return `"${val.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n')}"`;
    }
    return val;
  }
  if (Array.isArray(val)) return JSON.stringify(val);
  if (typeof val === 'object') return JSON.stringify(val);
  return String(val);
}

function extractItems(data: unknown): unknown[] {
  if (Array.isArray(data)) return data;
  if (typeof data === 'object' && data !== null) {
    const obj = data as Record<string, unknown>;
    if (Array.isArray(obj.value)) return obj.value;
  }
  return [data];
}

function csvEscape(val: string): string {
  // Prevent formula injection: prefix with tab if value starts with =, +, -, or @
  let safe = val;
  if (/^[=+\-@]/.test(safe)) {
    safe = '\t' + safe;
  }
  if (safe.includes(',') || safe.includes('"') || safe.includes('\n') || safe !== val) {
    return '"' + safe.replace(/"/g, '""') + '"';
  }
  return safe;
}

export interface ListOptions {
  all?: boolean;
  pageLimit?: string;
  pageDelay?: string;
  output?: string;
}

/**
 * Print a Graph collection: one page, the page named by --cursor, or every page with --all — streamed for
 * json/ndjson/csv. Completeness is in band (AH-3): JSON/YAML get `complete` (+ `incompleteReason`, `nextCursor`),
 * NDJSON ends with a `{"_meta": …}` line; table and csv print `nextCursor: <token>` on stderr.
 * `searchCap`: the most items Graph returns for this query in total (mail `$search`: 250).
 */
export async function printList(
  client: GraphClient,
  endpoint: string,
  options: GraphRequestOptions,
  opts: ListOptions,
  hooks: { transform?: (items: unknown[]) => Promise<unknown[]>; searchCap?: number } = {},
): Promise<void> {
  const transform = hooks.transform ?? (async (items: unknown[]) => items);
  const transformData = async (data: unknown): Promise<unknown> => {
    if (data && typeof data === 'object' && Array.isArray((data as Record<string, unknown>).value)) {
      return { ...(data as Record<string, unknown>), value: await transform((data as Record<string, unknown>).value as unknown[]) };
    }
    return data;
  };
  const format = (opts.output ?? 'json') as OutputFormat;
  const pageLimit = opts.pageLimit ? parseInt(opts.pageLimit, 10) : undefined;
  const pageDelay = opts.pageDelay ? parseInt(opts.pageDelay, 10) : undefined;
  const completeness = (nextLink: string | undefined, reason: 'page_limit' | 'next_cursor', count: number): Completeness => {
    if (nextLink) return { complete: false, incompleteReason: reason, nextCursor: encodeCursor(nextLink) };
    if (hooks.searchCap && count >= hooks.searchCap) return { complete: false, incompleteReason: 'search_cap' };
    return { complete: true };
  };
  const inBand = format === 'json' || format === 'yaml';
  const preview = client.isDryRun();

  if (opts.all && !pageCursor) {
    if (format === 'json' || format === 'ndjson' || format === 'csv') {
      const paging: PagingState = {};
      let count = 0;
      async function* pages() {
        for await (const page of client.requestPagesStreaming(endpoint, options, pageLimit, pageDelay, paging)) {
          count += page.length;
          yield await transform(page);
        }
      }
      const meta = () => completeness(paging.nextLink, 'page_limit', count);
      await streamPages(pages(), format, (chunk) => process.stdout.write(chunk), { jsonWrapper: true, ...(!preview && { meta }) });
      if (format === 'csv' && paging.nextLink) process.stderr.write(`nextCursor: ${meta().nextCursor}\n`);
    } else {
      const response = await client.requestAllPages(endpoint, options, pageLimit, pageDelay);
      const all = response.data;
      const count = all && typeof all === 'object' && 'value' in all && Array.isArray(all.value) ? all.value.length : 0;
      const meta = completeness(response.nextLink, 'page_limit', count);
      const data = await transformData(all);
      console.log(formatOutput(inBand && !preview && data && typeof data === 'object' ? { ...data, ...meta } : data, format));
      if (!inBand && meta.nextCursor) process.stderr.write(`nextCursor: ${meta.nextCursor}\n`);
    }
    return;
  }

  const response = pageCursor
    ? await client.request(decodeCursor(pageCursor, client.buildUrl(endpoint)), { headers: options.headers })
    : await client.request(endpoint, options);
  const pageData = await transformData(response.data);
  const items: unknown = pageData && typeof pageData === 'object' ? Reflect.get(pageData, 'value') : undefined;
  if (preview || !Array.isArray(items)) {
    console.log(formatOutput(pageData, format));
    return;
  }
  const meta = completeness(response.nextLink, 'next_cursor', items.length);
  console.log(formatOutput(inBand && pageData && typeof pageData === 'object' ? { ...pageData, ...meta } : pageData, format));
  if (format === 'ndjson') process.stdout.write(JSON.stringify({ _meta: meta }) + '\n');
  else if (!inBand && meta.nextCursor) process.stderr.write(`nextCursor: ${meta.nextCursor}\n`);
}
