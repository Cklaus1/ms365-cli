import { Command, InvalidArgumentError } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { searchQuote } from '../graph-helpers.js';
import { listRequest, paginationOptions, type ListOptions } from '../list.js';
import { outputOption } from '../utils.js';
import { handleErrors } from '../errors.js';

interface QueryOpts { entity: string; top: number; output: string }

/**
 * Entity types one /search/query request may combine (Graph "Known limitations" table):
 * each bucket mixes freely, types from different buckets are rejected by the service.
 */
const ENTITY_BUCKETS: string[][] = [
  ['driveItem', 'drive', 'list', 'listItem', 'site', 'externalItem'],
  ['message'],
  ['event'],
  ['chatMessage'],
  ['person'],
  ['acronym', 'bookmark', 'qna'],
];

function entityTypes(value: string): string[] {
  const types = value.split(',').map((t) => t.trim()).filter((t) => t !== '');
  const unknown = types.filter((t) => !ENTITY_BUCKETS.flat().includes(t));
  if (types.length === 0 || unknown.length > 0) {
    throw new Error(`Unknown entity type(s): ${unknown.join(', ') || '(none)'}. Valid: ${ENTITY_BUCKETS.flat().join(', ')}.`);
  }
  if (!ENTITY_BUCKETS.some((bucket) => types.every((t) => bucket.includes(t)))) {
    throw new Error(
      `${types.join(', ')} cannot be combined in one search. Combinable groups: ${ENTITY_BUCKETS.map((b) => b.join('+')).join(' | ')}. Run one search per group.`,
    );
  }
  return types;
}

function size(value: string): number {
  if (!/^\d+$/.test(value) || Number(value) < 1 || Number(value) > 1000) {
    throw new InvalidArgumentError('Expected an integer from 1 to 1000.');
  }
  return Number(value);
}

export function registerSearchCommands(program: Command, getClient: () => GraphClient): void {
  const search = program.command('search').description('Microsoft 365 search');

  search
    .command('query <query>')
    .description('Search across Microsoft 365 (KQL query string)')
    .option('--entity <types>', `Comma-separated entity types from one group: ${ENTITY_BUCKETS.map((b) => b.join(',')).join(' | ')}`, 'driveItem')
    .option('--top <n>', 'Number of results (1-1000)', size, 10)
    .addOption(outputOption())
    .action(handleErrors(async (query: string, opts: QueryOpts) => {
      const response = await getClient().request('/search/query', {
        method: 'POST',
        body: JSON.stringify({
          requests: [
            {
              entityTypes: entityTypes(opts.entity),
              query: { queryString: query },
              from: 0,
              size: opts.top,
            },
          ],
        }),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  paginationOptions(search
    .command('people [query]')
    .description('Search people relevant to you'), 10)
    .addOption(outputOption())
    .action(handleErrors(async (query: string | undefined, opts: ListOptions) => {
      const queryParams: Record<string, string> = {};
      if (query) queryParams['$search'] = searchQuote(query);
      await listRequest(getClient(), '/me/people', { queryParams }, opts);
    }));
}
