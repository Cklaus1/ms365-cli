import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat, printList } from '../formatter.js';
import { outputOption, info } from '../utils.js';
import { handleErrors } from '../errors.js';
import { resolve } from '../resolve.js';
import { addContentOptions, contentOptions, htmlToText, messageTransform } from '../content.js';

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
    .command('get')
    .argument('<teamId>', 'Team ID or name')
    .description('Get team details')
    .addOption(outputOption())
    .action(handleErrors(async (teamArg: string, opts) => {
      const client = getClient();
      const teamId = await resolve(client, 'team', teamArg);
      const response = await client.request(`/teams/${encodeURIComponent(teamId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  teams
    .command('members')
    .argument('<teamId>', 'Team ID or name')
    .description('List team members')
    .addOption(outputOption())
    .action(handleErrors(async (teamArg: string, opts) => {
      const client = getClient();
      const teamId = await resolve(client, 'team', teamArg);
      const response = await client.request(`/teams/${encodeURIComponent(teamId)}/members`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  teams
    .command('channels')
    .argument('<teamId>', 'Team ID or name')
    .description('List team channels')
    .addOption(outputOption())
    .action(handleErrors(async (teamArg: string, opts) => {
      const client = getClient();
      const teamId = await resolve(client, 'team', teamArg);
      const response = await client.request(`/teams/${encodeURIComponent(teamId)}/channels`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  teams
    .command('channel')
    .argument('<teamId>', 'Team ID or name')
    .argument('<channelId>', 'Channel ID or name')
    .description('Get channel details')
    .addOption(outputOption())
    .action(handleErrors(async (teamArg: string, channelArg: string, opts) => {
      const client = getClient();
      const teamId = await resolve(client, 'team', teamArg);
      const channelId = await resolve(client, 'channel', channelArg, teamId);
      const response = await client.request(
        `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  addContentOptions(teams
    .command('channel-messages')
    .argument('<teamId>', 'Team ID or name')
    .argument('<channelId>', 'Channel ID or name')
    .description('List channel messages (marked untrusted; see --body, --wrap-untrusted)')
    .option('--top <n>', 'Number of messages')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms'))
    .addOption(outputOption())
    .action(handleErrors(async (teamArg: string, channelArg: string, opts) => {
      const client = getClient();
      const teamId = await resolve(client, 'team', teamArg);
      const channelId = await resolve(client, 'channel', channelArg, teamId);
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;

      const endpoint = `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`;
      await printList(client, endpoint, { queryParams }, opts, { transform: messageTransform(client, 'teams', contentOptions(opts)) });
    }));

  teams
    .command('send-channel-message')
    .argument('<teamId>', 'Team ID or name')
    .argument('<channelId>', 'Channel ID or name')
    .description('Send a message to a channel')
    .requiredOption('--body <body>', 'Message body')
    .option('--html', 'Body is HTML')
    .action(handleErrors(async (teamArg: string, channelArg: string, opts) => {
      const client = getClient();
      const teamId = await resolve(client, 'team', teamArg);
      const channelId = await resolve(client, 'channel', channelArg, teamId);
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
    .command('reply-channel-message')
    .argument('<teamId>', 'Team ID or name')
    .argument('<channelId>', 'Channel ID or name')
    .argument('<messageId>', 'Message ID')
    .description('Reply to a channel message (threaded)')
    .requiredOption('--body <body>', 'Reply body')
    .option('--html', 'Body is HTML')
    .action(handleErrors(async (teamArg: string, channelArg: string, messageId: string, opts) => {
      const client = getClient();
      const teamId = await resolve(client, 'team', teamArg);
      const channelId = await resolve(client, 'channel', channelArg, teamId);
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

      await printList(client, '/me/chats', { queryParams }, opts);
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

  addContentOptions(chat
    .command('messages <chatId>')
    .description('List chat messages (marked untrusted; see --body, --wrap-untrusted)')
    .option('--top <n>', 'Number of messages')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms'))
    .addOption(outputOption())
    .action(handleErrors(async (chatId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;

      const endpoint = `/chats/${encodeURIComponent(chatId)}/messages`;
      await printList(client, endpoint, { queryParams }, opts, { transform: messageTransform(client, 'teams', contentOptions(opts)) });
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
      // Chats have no /replies endpoint; Graph's documented "reply with quote" is a new
      // message carrying a messageReference attachment to the original.
      const messagesEndpoint = `/chats/${encodeURIComponent(chatId)}/messages`;
      const original = (await client.request(`${messagesEndpoint}/${encodeURIComponent(messageId)}`)).data as {
        body?: { content?: string };
        from?: unknown;
      };
      const replyHtml = opts.html ? opts.body : escapeHtml(opts.body);
      await client.request(messagesEndpoint, {
        method: 'POST',
        body: JSON.stringify({
          body: { contentType: 'html', content: `<attachment id="${escapeHtml(messageId)}"></attachment>${replyHtml}` },
          attachments: [{
            id: messageId,
            contentType: 'messageReference',
            content: JSON.stringify({
              messageId,
              messagePreview: htmlToText(original.body?.content ?? ''),
              messageSender: original.from,
            }),
          }],
        }),
      });
      info('Reply sent.');
    }));
}

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
