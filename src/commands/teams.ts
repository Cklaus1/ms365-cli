import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption, done } from '../utils.js';
import { handleErrors } from '../errors.js';
import { listRequest, paginationOptions, type ListOptions } from '../list.js';
import { graph } from '../command-tree.js';

function messageBody(opts: { body: string; html?: boolean }): { contentType: string; content: string } {
  return { contentType: opts.html ? 'html' : 'text', content: opts.body };
}

interface SendOptions { body: string; html?: boolean; output: string }

/**
 * POST a message-like body and print the created message (the dry-run object under --dry-run) on stdout;
 * the confirmation goes to stderr and is skipped in dry-run, whose preview is already there.
 */
async function post(client: GraphClient, endpoint: string, body: unknown, message: string, output: string): Promise<void> {
  const res = await client.request(endpoint, { method: 'POST', body: JSON.stringify(body) });
  console.log(formatOutput(res.data, output as OutputFormat));
  done(res, message);
}

export function registerTeamsCommands(program: Command, getClient: () => GraphClient): void {
  const teams = program.command('teams').description('Microsoft Teams (requires --org-mode)');

  // joinedTeams and channels take no $top; drain every nextLink into one result.
  graph(teams.command('list'), ['GET /me/joinedTeams'])
    .description('List joined teams')
    .addOption(outputOption())
    .action(handleErrors(async (opts: { output: string }) => {
      await listRequest(getClient(), '/me/joinedTeams', {}, { output: opts.output, all: true });
    }));

  graph(teams.command('get <teamId>'), ['GET /teams/{team-id}'])
    .description('Get team details')
    .addOption(outputOption())
    .action(handleErrors(async (teamId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/teams/${encodeURIComponent(teamId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  paginationOptions(
    graph(teams.command('members <teamId>'), ['GET /teams/{team-id}/members'])
      .description('List team members'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (teamId: string, opts: ListOptions) => {
      await listRequest(getClient(), `/teams/${encodeURIComponent(teamId)}/members`, {}, opts);
    }));

  graph(teams.command('channels <teamId>'), ['GET /teams/{team-id}/channels'])
    .description('List team channels')
    .addOption(outputOption())
    .action(handleErrors(async (teamId: string, opts: { output: string }) => {
      await listRequest(getClient(), `/teams/${encodeURIComponent(teamId)}/channels`, {}, { output: opts.output, all: true });
    }));

  graph(teams.command('channel <teamId> <channelId>'), ['GET /teams/{team-id}/channels/{channel-id}'])
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
    graph(
      teams.command('channel-messages <teamId> <channelId>'),
      ['GET /teams/{team-id}/channels/{channel-id}/messages'],
    )
      .description('List channel messages'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (teamId: string, channelId: string, opts: ListOptions) => {
      const endpoint = `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`;
      await listRequest(getClient(), endpoint, {}, opts);
    }));

  graph(
    teams.command('send-channel-message <teamId> <channelId>'),
    ['POST /teams/{team-id}/channels/{channel-id}/messages'],
  )
    .description('Send a message to a channel')
    .requiredOption('--body <body>', 'Message body')
    .option('--html', 'Body is HTML')
    .addOption(outputOption())
    .action(handleErrors(async (teamId: string, channelId: string, opts: SendOptions) => {
      await post(
        getClient(),
        `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`,
        { body: messageBody(opts) },
        'Channel message sent.',
        opts.output,
      );
    }));

  graph(
    teams.command('reply-channel-message <teamId> <channelId> <messageId>'),
    ['POST /teams/{team-id}/channels/{channel-id}/messages/{message-id}/replies'],
  )
    .description('Reply to a channel message (threaded)')
    .requiredOption('--body <body>', 'Reply body')
    .option('--html', 'Body is HTML')
    .addOption(outputOption())
    .action(handleErrors(async (teamId: string, channelId: string, messageId: string, opts: SendOptions) => {
      await post(
        getClient(),
        `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/${encodeURIComponent(messageId)}/replies`,
        { body: messageBody(opts) },
        'Reply sent to channel thread.',
        opts.output,
      );
    }));

  // Chats
  const chat = program.command('chat').description('Teams chat (requires --org-mode)');

  paginationOptions(
    graph(chat.command('list'), ['GET /me/chats'])
      .description('List chats'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (opts: ListOptions) => {
      await listRequest(getClient(), '/me/chats', {}, opts);
    }));

  graph(chat.command('get <chatId>'), ['GET /chats/{chat-id}'])
    .description('Get chat details')
    .addOption(outputOption())
    .action(handleErrors(async (chatId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/chats/${encodeURIComponent(chatId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  paginationOptions(
    graph(chat.command('messages <chatId>'), ['GET /chats/{chat-id}/messages'])
      .description('List chat messages'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (chatId: string, opts: ListOptions) => {
      await listRequest(getClient(), `/chats/${encodeURIComponent(chatId)}/messages`, {}, opts);
    }));

  graph(chat.command('send <chatId>'), ['POST /chats/{chat-id}/messages'])
    .description('Send a chat message')
    .requiredOption('--body <body>', 'Message body')
    .option('--html', 'Body is HTML')
    .addOption(outputOption())
    .action(handleErrors(async (chatId: string, opts: SendOptions) => {
      await post(getClient(), `/chats/${encodeURIComponent(chatId)}/messages`, { body: messageBody(opts) }, 'Message sent.', opts.output);
    }));

  graph(chat.command('reply <chatId> <messageId>'), ['POST /chats/{chat-id}/messages/replyWithQuote'])
    .description('Reply to a chat message (posts a new message quoting it)')
    .requiredOption('--body <body>', 'Reply body')
    .option('--html', 'Body is HTML')
    .addOption(outputOption())
    .action(handleErrors(async (chatId: string, messageId: string, opts: SendOptions) => {
      // Chats have no threaded /replies (channels only); Graph's chat reply is replyWithQuote.
      await post(
        getClient(),
        `/chats/${encodeURIComponent(chatId)}/messages/replyWithQuote`,
        { messageIds: [messageId], replyMessage: { body: messageBody(opts) } },
        'Reply sent.',
        opts.output,
      );
    }));
}
