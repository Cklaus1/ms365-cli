import { Command } from 'commander';
import { GraphClient, type GraphRequestOptions } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { drivePath } from '../graph-helpers.js';
import { listRequest, paginationOptions, type ListOptions } from '../list.js';
import { info, outputOption } from '../utils.js';
import { handleErrors } from '../errors.js';

interface OutputOpts { output: string }
interface ListItemsOpts extends ListOptions { filter?: string; expand?: string }
interface DeltaOpts extends ListOptions { deltaLink?: string }

/**
 * Split a `@odata.nextLink`/`@odata.deltaLink` from an earlier `sp delta` run into endpoint, query and
 * API version, so it goes through the normal request path. Only Graph `/sites/delta` (or `/sites/delta()`)
 * links on https://graph.microsoft.com are accepted.
 */
function parseDeltaLink(link: string): { endpoint: string; request: GraphRequestOptions } {
  let url: URL;
  try {
    url = new URL(link);
  } catch {
    throw new Error(`Invalid --delta-link: ${link}`);
  }
  const match = /^\/(v1\.0|beta)(\/sites\/delta(?:\(\))?)$/.exec(url.pathname);
  if (url.origin !== 'https://graph.microsoft.com' || !match) {
    throw new Error('--delta-link must be a https://graph.microsoft.com/<version>/sites/delta link from a previous run.');
  }
  return {
    endpoint: match[2],
    request: { queryParams: Object.fromEntries(url.searchParams), apiVersion: match[1] === 'beta' ? 'beta' : 'v1.0' },
  };
}

export function registerSharePointCommands(program: Command, getClient: () => GraphClient): void {
  const sp = program.command('sharepoint').alias('sp').description('SharePoint (requires --org-mode)');
  const sitePath = (siteId: string) => `/sites/${encodeURIComponent(siteId)}`;

  paginationOptions(sp
    .command('search <query>')
    .description('Search SharePoint sites'))
    .addOption(outputOption())
    .action(handleErrors(async (query: string, opts: ListOptions) => {
      await listRequest(getClient(), '/sites', { queryParams: { search: query } }, opts);
    }));

  sp
    .command('site <siteId>')
    .description('Get site details')
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, opts: OutputOpts) => {
      const response = await getClient().request(sitePath(siteId));
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  sp
    .command('site-by-path <hostname> <path>')
    .description('Get site by hostname and path (e.g. contoso.sharepoint.com /sites/marketing)')
    .addOption(outputOption())
    .action(handleErrors(async (hostname: string, serverRelativePath: string, opts: OutputOpts) => {
      // Graph: /sites/{hostname}:/{server-relative-path}; `/` addresses the root site.
      const response = await getClient().request(drivePath(sitePath(hostname), serverRelativePath));
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  paginationOptions(sp
    .command('drives <siteId>')
    .description('List site drives (document libraries)'))
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, opts: ListOptions) => {
      await listRequest(getClient(), `${sitePath(siteId)}/drives`, {}, opts);
    }));

  sp
    .command('drive <siteId> <driveId>')
    .description('Get a specific site drive')
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, driveId: string, opts: OutputOpts) => {
      const response = await getClient().request(`${sitePath(siteId)}/drives/${encodeURIComponent(driveId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  paginationOptions(sp
    .command('lists <siteId>')
    .description('List site lists'))
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, opts: ListOptions) => {
      await listRequest(getClient(), `${sitePath(siteId)}/lists`, {}, opts);
    }));

  sp
    .command('list <siteId> <listId>')
    .description('Get a specific SharePoint list')
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, listId: string, opts: OutputOpts) => {
      const response = await getClient().request(`${sitePath(siteId)}/lists/${encodeURIComponent(listId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  paginationOptions(sp
    .command('list-items <siteId> <listId>')
    .description('List items in a SharePoint list')
    .option('--filter <filter>', 'OData filter')
    .option('--expand <fields>', 'Fields to expand'))
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, listId: string, opts: ListItemsOpts) => {
      const queryParams: Record<string, string> = {};
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.expand) queryParams['$expand'] = opts.expand;
      await listRequest(getClient(), `${sitePath(siteId)}/lists/${encodeURIComponent(listId)}/items`, { queryParams }, opts);
    }));

  sp
    .command('list-item <siteId> <listId> <itemId>')
    .description('Get a specific list item')
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, listId: string, itemId: string, opts: OutputOpts) => {
      const response = await getClient().request(
        `${sitePath(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}`,
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  sp
    .command('item <siteId> <itemId>')
    .description('Get a specific site item')
    .addOption(outputOption())
    .action(handleErrors(async (siteId: string, itemId: string, opts: OutputOpts) => {
      const response = await getClient().request(`${sitePath(siteId)}/items/${encodeURIComponent(itemId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  paginationOptions(sp
    .command('delta')
    .description('Get changed sites (delta sync). Prints { value, nextLink | deltaLink }: pass nextLink back via --delta-link to continue, deltaLink to get later changes. --all drains the round.')
    .option('--delta-link <url>', 'Resume from the nextLink or deltaLink of a previous run'))
    .addOption(outputOption())
    .action(handleErrors(async (opts: DeltaOpts) => {
      const { endpoint, request }: { endpoint: string; request: GraphRequestOptions } = opts.deltaLink
        ? parseDeltaLink(opts.deltaLink)
        : { endpoint: '/sites/delta()', request: {} };
      if (opts.top !== undefined) request.queryParams = { ...request.queryParams, $top: String(opts.top) };
      const client = getClient();
      const response = opts.all
        ? await client.requestAllPages(endpoint, request, opts.pageLimit, opts.pageDelay)
        : await client.request(endpoint, request);
      const data = response.data;
      const value = typeof data === 'object' && data !== null && 'value' in data ? data.value : [data];
      const result: Record<string, unknown> = { value };
      if (response.nextLink) result.nextLink = response.nextLink;
      if (response.deltaLink) result.deltaLink = response.deltaLink;
      if (response.nextLink) info('More changes pending: continue with --delta-link <nextLink>.');
      console.log(formatOutput(result, opts.output as OutputFormat));
    }));
}
