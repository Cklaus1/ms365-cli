import { Command, Argument } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { handleErrors } from '../errors.js';
import { outputOption } from '../utils.js';
import { resolve, resolveId, RESOLVE_TYPES, type ResolveType } from '../resolve.js';

export function registerResolveCommands(program: Command, getClient: () => GraphClient): void {
  program
    .command('resolve')
    .description('Resolve a name to its ID (exact, then unique prefix/substring match; mail-folder: exact name or path like Inbox/Projects)')
    .addArgument(new Argument('<type>', 'Resource type').choices(RESOLVE_TYPES))
    .argument('<name>', 'Name (or ID, returned as-is)')
    .option('--team <team>', 'Team ID or name (for type channel)')
    .option('--plan <plan>', 'Plan ID or name (for type bucket)')
    .option('--user <user>', "Mailbox user ID or UPN (for type mail-folder; default: your own)")
    .addOption(outputOption())
    .action(handleErrors(async (type: ResolveType, name: string, opts) => {
      const client = getClient();
      let parent: string | undefined;
      if (type === 'channel' && opts.team) parent = await resolve(client, 'team', opts.team);
      if (type === 'bucket' && opts.plan) parent = await resolve(client, 'plan', opts.plan);
      if (type === 'mail-folder') parent = opts.user;
      const resolved = await resolveId(client, type, name, parent);
      console.log(formatOutput(resolved, opts.output as OutputFormat));
    }));
}
