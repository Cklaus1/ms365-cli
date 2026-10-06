import { InvalidArgumentError, Option, type Command } from 'commander';
import type { GraphClient, GraphRequestOptions } from './graph-client.js';
import { formatOutput, streamPages, type OutputFormat } from './formatter.js';

/** Parsed values of the options `paginationOptions()` adds (plus the command's `-o`). */
export interface ListOptions {
  top?: number;
  all?: boolean;
  pageLimit?: number;
  pageDelay?: number;
  output: string;
}

function integer(min: number) {
  return (value: string): number => {
    if (!/^\d+$/.test(value) || Number(value) < min) {
      throw new InvalidArgumentError(`Expected an integer >= ${min}.`);
    }
    return Number(value);
  };
}

/**
 * Add the shared list options to `cmd` and return it (chainable):
 * `--top <n>` (page size / result count, integer >= 1, default `defaultTop` when given),
 * `--all` (follow every nextLink), `--page-limit <n>` (max pages with --all, >= 1; client default 100),
 * `--page-delay <ms>` (pause between pages with --all, >= 0). Values arrive parsed as numbers.
 */
export function paginationOptions(cmd: Command, defaultTop?: number): Command {
  const top = new Option('--top <n>', 'Number of results (page size with --all)').argParser(integer(1));
  if (defaultTop !== undefined) top.default(defaultTop);
  return cmd
    .addOption(top)
    .addOption(new Option('--all', 'Fetch all pages'))
    .addOption(new Option('--page-limit <n>', 'Max pages when using --all (default 100)').argParser(integer(1)))
    .addOption(new Option('--page-delay <ms>', 'Delay between pages in ms when using --all').argParser(integer(0)));
}

export interface ListRequestOptions extends ListOptions {
  /** Keep only matching items (applied to every page, e.g. a client-side filter Graph cannot express). */
  filterItems?: (item: unknown) => boolean;
}

/**
 * Run a list command end to end: GET `endpoint` with `request` (query params, headers, apiVersion
 * kept on every page), `$top` from `--top` unless `request.queryParams.$top` is set, then print.
 * Without `--all`: one page, rendered with `formatOutput`. With `--all`: json/csv stream page by page
 * (json as one array), other formats buffer every page; both honour `--page-limit`/`--page-delay`.
 */
export async function listRequest(
  client: GraphClient,
  endpoint: string,
  request: GraphRequestOptions,
  opts: ListRequestOptions,
  write: (chunk: string) => void = (s) => process.stdout.write(s),
): Promise<void> {
  const queryParams = { ...request.queryParams };
  if (opts.top !== undefined && queryParams.$top === undefined) queryParams.$top = String(opts.top);
  const req: GraphRequestOptions = { ...request, queryParams };
  const format = opts.output as OutputFormat;
  const keep = opts.filterItems;

  if (!opts.all) {
    const res = await client.request(endpoint, req);
    write(formatOutput(keep ? filterValue(res.data, keep) : res.data, format) + '\n');
    return;
  }

  if (format === 'json' || format === 'csv') {
    const pages = client.requestPagesStreaming(endpoint, req, opts.pageLimit, opts.pageDelay);
    await streamPages(keep ? filterPages(pages, keep) : pages, format, write);
    return;
  }

  const res = await client.requestAllPages(endpoint, req, opts.pageLimit, opts.pageDelay);
  write(formatOutput(keep ? filterValue(res.data, keep) : res.data, format) + '\n');
}

function filterValue(data: unknown, keep: (item: unknown) => boolean): unknown {
  if (typeof data === 'object' && data !== null && 'value' in data && Array.isArray(data.value)) {
    return { ...data, value: data.value.filter(keep) };
  }
  return data;
}

async function* filterPages(
  pages: AsyncGenerator<unknown[], void, undefined>,
  keep: (item: unknown) => boolean,
): AsyncGenerator<unknown[], void, undefined> {
  for await (const page of pages) yield page.filter(keep);
}
