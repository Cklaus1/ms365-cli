import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption } from '../utils.js';
import { handleErrors } from '../errors.js';

export function registerTeamsCommands(program: Command, getClient: () => GraphClient): void {
  const teams = program.command('teams').description('Microsoft Teams (requires --org-mode)');

  teams
    .command('list')
    .description('List joined teams')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/joinedTeams');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  teams
    .command('get <teamId>')
    .description('Get team details')
    .addOption(outputOption())
    .action(handleErrors(async (teamId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/teams/${encodeURIComponent(teamId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  teams
    .command('members <teamId>')
    .description('List team members')
    .addOption(outputOption())
    .action(handleErrors(async (teamId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/teams/${encodeURIComponent(teamId)}/members`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  teams
    .command('channels <teamId>')
    .description('List team channels')
    .addOption(outputOption())
    .action(handleErrors(async (teamId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/teams/${encodeURIComponent(teamId)}/channels`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  teams
    .command('channel <teamId> <channelId>')
    .description('Get channel details')
    .addOption(outputOption())
    .action(handleErrors(async (teamId: string, channelId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  teams
    .command('channel-messages <teamId> <channelId>')
    .description('List channel messages')
    .option('--top <n>', 'Number of messages')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (teamId: string, channelId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;

      const endpoint = `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`;
      const response = opts.all
        ? await client.requestAllPages(endpoint, { queryParams }, opts.pageLimit ? parseInt(opts.pageLimit) : undefined, opts.pageDelay ? parseInt(opts.pageDelay) : undefined)
        : await client.request(endpoint, { queryParams });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  teams
    .command('send-channel-message <teamId> <channelId>')
    .description('Send a message to a channel')
    .requiredOption('--body <body>', 'Message body')
    .option('--html', 'Body is HTML')
    .action(handleErrors(async (teamId: string, channelId: string, opts) => {
      const client = getClient();
      await client.request(
        `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`,
        {
          method: 'POST',
          body: JSON.stringify({
            body: { contentType: opts.html ? 'html' : 'text', content: opts.body },
          }),
        }
      );
      info('Channel message sent.');
    }));

  teams
    .command('reply-channel-message <teamId> <channelId> <messageId>')
    .description('Reply to a channel message (threaded)')
    .requiredOption('--body <body>', 'Reply body')
    .option('--html', 'Body is HTML')
    .action(handleErrors(async (teamId: string, channelId: string, messageId: string, opts) => {
      const client = getClient();
      await client.request(
        `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/${encodeURIComponent(messageId)}/replies`,
        {
          method: 'POST',
          body: JSON.stringify({
            body: { contentType: opts.html ? 'html' : 'text', content: opts.body },
          }),
        }
      );
      info('Reply sent to channel thread.');
    }));

  // Chats
  const chat = program.command('chat').description('Teams chat (requires --org-mode)');

  chat
    .command('list')
    .description('List chats')
    .option('--top <n>', 'Number of chats')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;

      const response = opts.all
        ? await client.requestAllPages('/me/chats', { queryParams })
        : await client.request('/me/chats', { queryParams });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  chat
    .command('get <chatId>')
    .description('Get chat details')
    .addOption(outputOption())
    .action(handleErrors(async (chatId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/chats/${encodeURIComponent(chatId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  chat
    .command('messages <chatId>')
    .description('List chat messages')
    .option('--top <n>', 'Number of messages')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (chatId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;

      const endpoint = `/chats/${encodeURIComponent(chatId)}/messages`;
      const response = opts.all
        ? await client.requestAllPages(endpoint, { queryParams }, opts.pageLimit ? parseInt(opts.pageLimit) : undefined, opts.pageDelay ? parseInt(opts.pageDelay) : undefined)
        : await client.request(endpoint, { queryParams });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  chat
    .command('send <chatId>')
    .description('Send a chat message')
    .requiredOption('--body <body>', 'Message body')
    .option('--html', 'Body is HTML')
    .action(handleErrors(async (chatId: string, opts) => {
      const client = getClient();
      await client.request(`/chats/${encodeURIComponent(chatId)}/messages`, {
        method: 'POST',
        body: JSON.stringify({
          body: { contentType: opts.html ? 'html' : 'text', content: opts.body },
        }),
      });
      info('Message sent.');
    }));

  chat
    .command('reply <chatId> <messageId>')
    .description('Reply to a chat message')
    .requiredOption('--body <body>', 'Reply body')
    .option('--html', 'Body is HTML')
    .action(handleErrors(async (chatId: string, messageId: string, opts) => {
      const client = getClient();
      await client.request(
        `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}/replies`,
        {
          method: 'POST',
          body: JSON.stringify({
            body: { contentType: opts.html ? 'html' : 'text', content: opts.body },
          }),
        }
      );
      info('Reply sent.');
    }));
}
