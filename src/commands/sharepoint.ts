import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption } from '../utils.js';
import { handleErrors } from '../errors.js';

export function registerSharePointCommands(program: Command, getClient: () => GraphClient): void {
  const sp = program.command('sharepoint').alias('sp').description('SharePoint (requires --org-mode)');

  sp
    .command('search <query>')
    .description('Search SharePoint sites')
    .addOption(outputOption())
    .action(handleErrors(async (query: string, opts) => {
      const client = getClient();
      const response = await client.request('/sites', {
        queryParams: { search: query },
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  sp
    .command('site <siteId>')
    .description('Get site details')
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/sites/${encodeURIComponent(siteId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  sp
    .command('site-by-path <hostname> <path>')
    .description('Get site by hostname and path (e.g. contoso.sharepoint.com /sites/marketing)')
    .addOption(outputOption())
    .action(handleErrors(async (hostname: string, sitePath: string, opts) => {
      const client = getClient();
      // Graph API format: /sites/{hostname}:/{server-relative-path}
      const normalizedPath = sitePath.startsWith('/') ? sitePath : `/${sitePath}`;
      const response = await client.request(
        `/sites/${encodeURIComponent(hostname)}:${normalizedPath}`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  sp
    .command('drives <siteId>')
    .description('List site drives (document libraries)')
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/sites/${encodeURIComponent(siteId)}/drives`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  sp
    .command('drive <siteId> <driveId>')
    .description('Get a specific site drive')
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, driveId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/sites/${encodeURIComponent(siteId)}/drives/${encodeURIComponent(driveId)}`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  sp
    .command('lists <siteId>')
    .description('List site lists')
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/sites/${encodeURIComponent(siteId)}/lists`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  sp
    .command('list <siteId> <listId>')
    .description('Get a specific SharePoint list')
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, listId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  sp
    .command('list-items <siteId> <listId>')
    .description('List items in a SharePoint list')
    .option('--top <n>', 'Number of items')
    .option('--filter <filter>', 'OData filter')
    .option('--expand <fields>', 'Fields to expand')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, listId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.expand) queryParams['$expand'] = opts.expand;

      const endpoint = `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`;
      const response = opts.all
        ? await client.requestAllPages(endpoint, { queryParams }, opts.pageLimit ? parseInt(opts.pageLimit) : undefined, opts.pageDelay ? parseInt(opts.pageDelay) : undefined)
        : await client.request(endpoint, { queryParams });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  sp
    .command('list-item <siteId> <listId> <itemId>')
    .description('Get a specific list item')
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, listId: string, itemId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  sp
    .command('items <siteId>')
    .description('List site items')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, opts) => {
      const client = getClient();
      const endpoint = `/sites/${encodeURIComponent(siteId)}/items`;
      const response = opts.all
        ? await client.requestAllPages(endpoint, {}, opts.pageLimit ? parseInt(opts.pageLimit) : undefined, opts.pageDelay ? parseInt(opts.pageDelay) : undefined)
        : await client.request(endpoint);

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  sp
    .command('item <siteId> <itemId>')
    .description('Get a specific site item')
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, itemId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/sites/${encodeURIComponent(siteId)}/items/${encodeURIComponent(itemId)}`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  sp
    .command('delta')
    .description('Get changed sites (delta sync)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/sites/delta()');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));
}
