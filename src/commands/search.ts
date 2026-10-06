import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption } from '../utils.js';
import { handleErrors } from '../errors.js';

export function registerSearchCommands(program: Command, getClient: () => GraphClient): void {
  const search = program.command('search').description('Microsoft 365 search');

  search
    .command('query <query>')
    .description('Search across Microsoft 365')
    .option('--entity <types>', 'Comma-separated entity types: message, driveItem, listItem, site, event, chatMessage', 'driveItem')
    .option('--top <n>', 'Number of results per entity type', '10')
    .addOption(outputOption())
    .action(handleErrors(async (query: string, opts) => {
      const client = getClient();
      const entityTypes = opts.entity.split(',').map((e: string) => e.trim());

      const response = await client.request('/search/query', {
        method: 'POST',
        body: JSON.stringify({
          requests: [
            {
              entityTypes,
              query: { queryString: query },
              from: 0,
              size: parseInt(opts.top),
            },
          ],
        }),
      });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  search
    .command('people [query]')
    .description('Search people')
    .option('--top <n>', 'Number of results', '10')
    .addOption(outputOption())
    .action(handleErrors(async (query: string | undefined, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (query) queryParams['$search'] = `"${query}"`;

      const response = await client.request('/me/people', { queryParams });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));
}
