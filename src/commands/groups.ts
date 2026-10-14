import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat, printList } from '../formatter.js';
import { addListOptions, outputOption, kqlQuote } from '../utils.js';
import { handleErrors } from '../errors.js';

export function registerGroupsCommands(program: Command, getClient: () => GraphClient): void {
  const groups = program.command('groups').description('Microsoft 365 Groups and distribution lists (requires --org-mode)');

  addListOptions(groups
    .command('list')
    .description('List groups in the organization')
    .option('--filter <filter>', 'OData filter')
    .option('--orderby <field>', 'Order results by field')
    .option('--search <query>', 'Search by displayName')
    .option('--select <fields>', 'Fields to return'), { description: 'Number of groups', default: '25' })
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.orderby) queryParams['$orderby'] = opts.orderby;
      if (opts.search) queryParams['$search'] = kqlQuote(`displayName:${opts.search}`);
      if (opts.select) queryParams['$select'] = opts.select;

      // ConsistencyLevel header required for $search
      const headers: Record<string, string> = {};
      if (opts.search) headers['ConsistencyLevel'] = 'eventual';

      await printList(client, '/groups', { queryParams, headers }, opts);
    }));

  addListOptions(groups
    .command('my')
    .description('List groups you are a member of')
    .option('--select <fields>', 'Fields to return'), { description: 'Number of groups', default: '50' })
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.select) queryParams['$select'] = opts.select;

      await printList(client, '/me/memberOf/microsoft.graph.group', { queryParams }, opts);
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

  addListOptions(groups
    .command('members <groupId>')
    .description('List group members'), { description: 'Number of members' })
    .addOption(outputOption())
    .action(handleErrors(async (groupId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;

      const endpoint = `/groups/${encodeURIComponent(groupId)}/members`;
      await printList(client, endpoint, { queryParams }, opts);
    }));

  addListOptions(groups
    .command('owners <groupId>')
    .description('List group owners'))
    .addOption(outputOption())
    .action(handleErrors(async (groupId: string, opts) => {
      const client = getClient();
      await printList(client, `/groups/${encodeURIComponent(groupId)}/owners`, {}, opts);
    }));

  // Distribution lists use the same /groups endpoint with mailEnabled=true
  addListOptions(groups
    .command('distribution-lists')
    .description('List mail-enabled distribution groups'), { description: 'Number of results', default: '25' })
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {
        '$filter': "mailEnabled eq true and securityEnabled eq false",
        '$select': 'id,displayName,mail,description,membershipRule',
      };
      if (opts.top) queryParams['$top'] = opts.top;

      await printList(client, '/groups', { queryParams }, opts);
    }));

  addListOptions(groups
    .command('teams')
    .description('List groups that have a Team associated'), { description: 'Number of results', default: '25' })
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {
        '$filter': "resourceProvisioningOptions/Any(x:x eq 'Team')",
        '$select': 'id,displayName,mail,description',
      };
      if (opts.top) queryParams['$top'] = opts.top;

      await printList(client, '/groups', { queryParams }, opts);
    }));
}
