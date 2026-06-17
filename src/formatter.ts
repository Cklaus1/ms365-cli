export type OutputFormat = 'json' | 'table' | 'csv' | 'text' | 'yaml';

/**
 * Stream formatted output page-by-page from an async generator.
 * JSON streams as a valid array; CSV streams rows after a single header;
 * other formats fall back to buffering all items.
 */
export async function streamPages(
  pages: AsyncGenerator<unknown[], void, undefined>,
  format: OutputFormat,
  write: (chunk: string) => void = (s) => process.stdout.write(s),
): Promise<void> {
  if (format === 'json') {
    let first = true;
    write('[\n');
    for await (const page of pages) {
      for (const item of page) {
        if (!first) write(',\n');
        write(JSON.stringify(item, null, 2));
        first = false;
      }
    }
    write('\n]\n');
    return;
  }

  if (format === 'csv') {
    let headerWritten = false;
    const columnSet = new Set<string>();
    let columns: string[] = [];
    for await (const page of pages) {
      const flatItems = page.map((item) => {
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

export function formatOutput(data: unknown, format: OutputFormat): string {
  switch (format) {
    case 'json':
      return JSON.stringify(data, null, 2);
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
  const columns = Array.from(keySet);
  if (columns.length === 0) return JSON.stringify(data, null, 2);

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
