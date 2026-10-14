import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat, printList } from '../formatter.js';
import { addListOptions, outputOption, encodePath } from '../utils.js';
import { handleErrors } from '../errors.js';

export function registerSharePointCommands(program: Command, getClient: () => GraphClient): void {
  const sp = program.command('sharepoint').alias('sp').description('SharePoint (requires --org-mode)');

  addListOptions(sp
    .command('search <query>')
    .description('Search SharePoint sites'))
    .addOption(outputOption())
    .action(handleErrors(async (query: string, opts) => {
      await printList(getClient(), '/sites', { queryParams: { search: query } }, opts);
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
        `/sites/${encodeURIComponent(hostname)}:${encodePath(normalizedPath)}`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  addListOptions(sp
    .command('drives <siteId>')
    .description('List site drives (document libraries)'))
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, opts) => {
      await printList(getClient(), `/sites/${encodeURIComponent(siteId)}/drives`, {}, opts);
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

  addListOptions(sp
    .command('lists <siteId>')
    .description('List site lists'))
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, opts) => {
      await printList(getClient(), `/sites/${encodeURIComponent(siteId)}/lists`, {}, opts);
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

  addListOptions(sp
    .command('list-items <siteId> <listId>')
    .description('List items in a SharePoint list')
    .option('--filter <filter>', 'OData filter')
    .option('--expand <fields>', 'Fields to expand'), { description: 'Number of items' })
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, listId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.expand) queryParams['$expand'] = opts.expand;

      const endpoint = `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`;
      await printList(client, endpoint, { queryParams }, opts);
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

  addListOptions(sp
    .command('items <siteId>')
    .description('List site items'))
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, opts) => {
      const client = getClient();
      const endpoint = `/sites/${encodeURIComponent(siteId)}/items`;
      await printList(client, endpoint, {}, opts);
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
    .description('Get changed sites (delta sync): every page up to the deltaLink')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      let response = await client.request('/sites/delta()');
      if (client.isPreview(response)) {
        console.log(formatOutput(response.data, opts.output as OutputFormat));
        return;
      }
      // A delta feed is only complete at its deltaLink: follow every nextLink before printing.
      const value: unknown[] = [];
      for (;;) {
        value.push(...((response.data as { value?: unknown[] }).value ?? []));
        if (!response.nextLink) break;
        response = await client.request(response.nextLink);
      }
      const data = response.deltaLink ? { value, deltaLink: response.deltaLink } : { value };
      console.log(formatOutput(data, opts.output as OutputFormat));
    }));
}
