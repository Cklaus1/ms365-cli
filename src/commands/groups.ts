import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption } from '../utils.js';
import { handleErrors } from '../errors.js';
import { searchQuote } from '../graph-helpers.js';
import { listRequest, paginationOptions, type ListOptions } from '../list.js';

/** Microsoft 365 groups are mail-enabled too; only non-`Unified` groups are distribution lists. */
function isNotUnified(item: unknown): boolean {
  if (typeof item !== 'object' || item === null || !('groupTypes' in item)) return true;
  const types = item.groupTypes;
  return !(Array.isArray(types) && types.includes('Unified'));
}

export function registerGroupsCommands(program: Command, getClient: () => GraphClient): void {
  const groups = program.command('groups').description('Microsoft 365 Groups and distribution lists (requires --org-mode)');

  paginationOptions(
    groups
      .command('list')
      .description('List groups in the organization')
      .option('--filter <filter>', 'OData filter')
      .option('--orderby <field>', 'Order results by field')
      .option('--search <query>', 'Search by displayName')
      .option('--select <fields>', 'Fields to return'),
    25,
  )
    .addOption(outputOption())
    .action(handleErrors(async (opts: ListOptions & { filter?: string; orderby?: string; search?: string; select?: string }) => {
      const queryParams: Record<string, string> = {};
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.orderby) queryParams['$orderby'] = opts.orderby;
      if (opts.search) queryParams['$search'] = searchQuote(`displayName:${opts.search}`);
      if (opts.select) queryParams['$select'] = opts.select;

      // ConsistencyLevel header required for $search
      const headers: Record<string, string> = {};
      if (opts.search) headers['ConsistencyLevel'] = 'eventual';

      await listRequest(getClient(), '/groups', { queryParams, headers }, opts);
    }));

  paginationOptions(
    groups
      .command('my')
      .description('List groups you are a member of')
      .option('--select <fields>', 'Fields to return'),
    50,
  )
    .addOption(outputOption())
    .action(handleErrors(async (opts: ListOptions & { select?: string }) => {
      const queryParams: Record<string, string> = {};
      if (opts.select) queryParams['$select'] = opts.select;
      await listRequest(getClient(), '/me/memberOf/microsoft.graph.group', { queryParams }, opts);
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

  paginationOptions(
    groups
      .command('members <groupId>')
      .description('List group members'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (groupId: string, opts: ListOptions) => {
      await listRequest(getClient(), `/groups/${encodeURIComponent(groupId)}/members`, {}, opts);
    }));

  paginationOptions(
    groups
      .command('owners <groupId>')
      .description('List group owners'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (groupId: string, opts: ListOptions) => {
      await listRequest(getClient(), `/groups/${encodeURIComponent(groupId)}/owners`, {}, opts);
    }));

  // Distribution lists use the same /groups endpoint with mailEnabled=true; M365 (Unified) groups are dropped client-side.
  paginationOptions(
    groups
      .command('distribution-lists')
      .description('List mail-enabled distribution groups (excludes Microsoft 365 groups)'),
    25,
  )
    .addOption(outputOption())
    .action(handleErrors(async (opts: ListOptions) => {
      const queryParams = {
        '$filter': 'mailEnabled eq true and securityEnabled eq false',
        '$select': 'id,displayName,mail,description,membershipRule,groupTypes',
      };
      await listRequest(getClient(), '/groups', { queryParams }, { ...opts, filterItems: isNotUnified });
    }));

  paginationOptions(
    groups
      .command('teams')
      .description('List groups that have a Team associated'),
    25,
  )
    .addOption(outputOption())
    .action(handleErrors(async (opts: ListOptions) => {
      const queryParams = {
        '$filter': "resourceProvisioningOptions/Any(x:x eq 'Team')",
        '$select': 'id,displayName,mail,description',
      };
      await listRequest(getClient(), '/groups', { queryParams }, opts);
    }));
}
