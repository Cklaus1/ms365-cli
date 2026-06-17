import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption } from '../utils.js';
import { handleErrors } from '../errors.js';

export function registerGroupsCommands(program: Command, getClient: () => GraphClient): void {
  const groups = program.command('groups').description('Microsoft 365 Groups and distribution lists (requires --org-mode)');

  groups
    .command('list')
    .description('List groups in the organization')
    .option('--top <n>', 'Number of groups', '25')
    .option('--filter <filter>', 'OData filter')
    .option('--orderby <field>', 'Order results by field')
    .option('--search <query>', 'Search by displayName')
    .option('--select <fields>', 'Fields to return')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.orderby) queryParams['$orderby'] = opts.orderby;
      if (opts.search) queryParams['$search'] = `"displayName:${opts.search}"`;
      if (opts.select) queryParams['$select'] = opts.select;

      // ConsistencyLevel header required for $search
      const headers: Record<string, string> = {};
      if (opts.search) headers['ConsistencyLevel'] = 'eventual';

      const response = opts.all
        ? await client.requestAllPages('/groups', { queryParams, headers })
        : await client.request('/groups', { queryParams, headers });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  groups
    .command('my')
    .description('List groups you are a member of')
    .option('--top <n>', 'Number of groups', '50')
    .option('--select <fields>', 'Fields to return')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.select) queryParams['$select'] = opts.select;

      const response = opts.all
        ? await client.requestAllPages('/me/memberOf/microsoft.graph.group', { queryParams })
        : await client.request('/me/memberOf/microsoft.graph.group', { queryParams });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  groups
    .command('get <groupId>')
    .description('Get group details')
    .addOption(outputOption())
    .action(handleErrors(async (groupId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/groups/${encodeURIComponent(groupId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  groups
    .command('members <groupId>')
    .description('List group members')
    .option('--top <n>', 'Number of members')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (groupId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;

      const endpoint = `/groups/${encodeURIComponent(groupId)}/members`;
      const response = opts.all
        ? await client.requestAllPages(endpoint, { queryParams }, opts.pageLimit ? parseInt(opts.pageLimit) : undefined, opts.pageDelay ? parseInt(opts.pageDelay) : undefined)
        : await client.request(endpoint, { queryParams });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  groups
    .command('owners <groupId>')
    .description('List group owners')
    .addOption(outputOption())
    .action(handleErrors(async (groupId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/groups/${encodeURIComponent(groupId)}/owners`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // Distribution lists use the same /groups endpoint with mailEnabled=true
  groups
    .command('distribution-lists')
    .description('List mail-enabled distribution groups')
    .option('--top <n>', 'Number of results', '25')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {
        '$filter': "mailEnabled eq true and securityEnabled eq false",
        '$select': 'id,displayName,mail,description,membershipRule',
      };
      if (opts.top) queryParams['$top'] = opts.top;

      const response = opts.all
        ? await client.requestAllPages('/groups', { queryParams })
        : await client.request('/groups', { queryParams });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  groups
    .command('teams')
    .description('List groups that have a Team associated')
    .option('--top <n>', 'Number of results', '25')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {
        '$filter': "resourceProvisioningOptions/Any(x:x eq 'Team')",
        '$select': 'id,displayName,mail,description',
      };
      if (opts.top) queryParams['$top'] = opts.top;

      const response = opts.all
        ? await client.requestAllPages('/groups', { queryParams })
        : await client.request('/groups', { queryParams });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));
}
