import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption, done } from '../utils.js';
import { handleErrors } from '../errors.js';
import { listRequest, paginationOptions, type ListOptions } from '../list.js';

function messageBody(opts: { body: string; html?: boolean }): { contentType: string; content: string } {
  return { contentType: opts.html ? 'html' : 'text', content: opts.body };
}

/** POST a message-like body; the success note is skipped in dry-run (status 0), whose preview is on stderr. */
async function post(client: GraphClient, endpoint: string, body: unknown, message: string): Promise<void> {
  done(await client.request(endpoint, { method: 'POST', body: JSON.stringify(body) }), message);
}

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

  paginationOptions(
    teams
      .command('channel-messages <teamId> <channelId>')
      .description('List channel messages'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (teamId: string, channelId: string, opts: ListOptions) => {
      const endpoint = `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`;
      await listRequest(getClient(), endpoint, {}, opts);
    }));

  teams
    .command('send-channel-message <teamId> <channelId>')
    .description('Send a message to a channel')
    .requiredOption('--body <body>', 'Message body')
    .option('--html', 'Body is HTML')
    .action(handleErrors(async (teamId: string, channelId: string, opts: { body: string; html?: boolean }) => {
      await post(
        getClient(),
        `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`,
        { body: messageBody(opts) },
        'Channel message sent.',
      );
    }));

  teams
    .command('reply-channel-message <teamId> <channelId> <messageId>')
    .description('Reply to a channel message (threaded)')
    .requiredOption('--body <body>', 'Reply body')
    .option('--html', 'Body is HTML')
    .action(handleErrors(async (teamId: string, channelId: string, messageId: string, opts: { body: string; html?: boolean }) => {
      await post(
        getClient(),
        `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/${encodeURIComponent(messageId)}/replies`,
        { body: messageBody(opts) },
        'Reply sent to channel thread.',
      );
    }));

  // Chats
  const chat = program.command('chat').description('Teams chat (requires --org-mode)');

  paginationOptions(
    chat
      .command('list')
      .description('List chats'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (opts: ListOptions) => {
      await listRequest(getClient(), '/me/chats', {}, opts);
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

  paginationOptions(
    chat
      .command('messages <chatId>')
      .description('List chat messages'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (chatId: string, opts: ListOptions) => {
      await listRequest(getClient(), `/chats/${encodeURIComponent(chatId)}/messages`, {}, opts);
    }));

  chat
    .command('send <chatId>')
    .description('Send a chat message')
    .requiredOption('--body <body>', 'Message body')
    .option('--html', 'Body is HTML')
    .action(handleErrors(async (chatId: string, opts: { body: string; html?: boolean }) => {
      await post(getClient(), `/chats/${encodeURIComponent(chatId)}/messages`, { body: messageBody(opts) }, 'Message sent.');
    }));

  chat
    .command('reply <chatId> <messageId>')
    .description('Reply to a chat message (posts a new message quoting it)')
    .requiredOption('--body <body>', 'Reply body')
    .option('--html', 'Body is HTML')
    .action(handleErrors(async (chatId: string, messageId: string, opts: { body: string; html?: boolean }) => {
      // Chats have no threaded /replies (channels only); Graph's chat reply is replyWithQuote.
      await post(
        getClient(),
        `/chats/${encodeURIComponent(chatId)}/messages/replyWithQuote`,
        { messageIds: [messageId], replyMessage: { body: messageBody(opts) } },
        'Reply sent.',
      );
    }));
}
