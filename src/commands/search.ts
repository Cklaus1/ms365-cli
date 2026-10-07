import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { searchQuote } from '../graph-helpers.js';
import { integer, listRequest, paginationOptions, type ListOptions } from '../list.js';
import { outputOption } from '../utils.js';
import { handleErrors } from '../errors.js';
import { listChoicesOption, listItems } from '../command-tree.js';

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

/** Split `--entity` (items already checked against the known types by its option parser) and enforce the bucket rule. */
function entityTypes(value: string): string[] {
  const types = listItems(value);
  if (!ENTITY_BUCKETS.some((bucket) => types.every((t) => bucket.includes(t)))) {
    throw new Error(
      `${types.join(', ')} cannot be combined in one search. Combinable groups: ${ENTITY_BUCKETS.map((b) => b.join('+')).join(' | ')}. Run one search per group.`,
    );
  }
  return types;
}

export function registerSearchCommands(program: Command, getClient: () => GraphClient): void {
  const search = program.command('search').description('Microsoft 365 search');

  search
    .command('query <query>')
    .description('Search across Microsoft 365 (KQL query string)')
    .addOption(
      listChoicesOption(
        '--entity <types>',
        `Comma-separated entity types from one group: ${ENTITY_BUCKETS.map((b) => b.join(',')).join(' | ')}`,
        ENTITY_BUCKETS.flat(),
      ).default('driveItem'),
    )
    .option('--top <n>', 'Number of results (1-1000)', integer(1, 1000), 10)
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
