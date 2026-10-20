import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat, printList } from '../formatter.js';
import { addListOptions, outputOption, success, odataQuote, validateEmail, addBodyOptions, bodyText } from '../utils.js';
import { idempotentSend, IDEMPOTENCY_HELP } from '../idempotency.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import { resolve } from '../resolve.js';
import { addContentOptions, chatTransform, contentOptions, htmlToText, messageTransform } from '../content.js';

export function registerTeamsCommands(program: Command, getClient: () => GraphClient): void {
  const teams = program.command('teams').description('Microsoft Teams (requires --org-mode)');

  addListOptions(teams
    .command('list')
    .description('List joined teams'))
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      await printList(getClient(), '/me/joinedTeams', {}, opts);
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

  addListOptions(teams
    .command('members')
    .argument('<teamId>', 'Team ID or name')
    .description('List team members'))
    .addOption(outputOption())
    .action(handleErrors(async (teamArg: string, opts) => {
      const client = getClient();
      const teamId = await resolve(client, 'team', teamArg);
      await printList(client, `/teams/${encodeURIComponent(teamId)}/members`, {}, opts);
    }));

  addListOptions(teams
    .command('channels')
    .argument('<teamId>', 'Team ID or name')
    .description('List team channels'))
    .addOption(outputOption())
    .action(handleErrors(async (teamArg: string, opts) => {
      const client = getClient();
      const teamId = await resolve(client, 'team', teamArg);
      await printList(client, `/teams/${encodeURIComponent(teamId)}/channels`, {}, opts);
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

  addListOptions(addContentOptions(teams
    .command('channel-messages')
    .argument('<teamId>', 'Team ID or name')
    .argument('<channelId>', 'Channel ID or name')
    .description('List channel messages (marked untrusted; see --body, --wrap-untrusted)')), { description: 'Number of messages' })
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

  addBodyOptions(teams
    .command('send-channel-message')
    .argument('<teamId>', 'Team ID or name')
    .argument('<channelId>', 'Channel ID or name')
    .description('Send a message to a channel'), 'Message body', true)
    .option('--html', 'Body is HTML')
    .option('--mention <upn>', '@mention a user by UPN/email (repeatable)', collect, [])
    .option('--idempotency-key <key>', IDEMPOTENCY_HELP)
    .action(handleErrors(async (teamArg: string, channelArg: string, opts) => {
      const content = bodyText(opts) ?? '';
      const upns = mentionUpns(opts.mention);
      const client = getClient();
      const teamId = await resolve(client, 'team', teamArg);
      const channelId = await resolve(client, 'channel', channelArg, teamId);
      const payload = await messagePayload(client, content, !!opts.html, upns);
      const request = { team: teamArg, channel: channelArg, payload };
      const sent = await idempotentSend(client, opts.idempotencyKey, 'teams send-channel-message', request, async () => {
        await client.request(
          `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`,
          { method: 'POST', body: JSON.stringify(payload) }
        );
      });
      if (sent) success('Channel message sent.');
    }));

  addBodyOptions(teams
    .command('reply-channel-message')
    .argument('<teamId>', 'Team ID or name')
    .argument('<channelId>', 'Channel ID or name')
    .argument('<messageId>', 'Message ID')
    .description('Reply to a channel message (threaded)'), 'Reply body', true)
    .option('--html', 'Body is HTML')
    .option('--idempotency-key <key>', IDEMPOTENCY_HELP)
    .action(handleErrors(async (teamArg: string, channelArg: string, messageId: string, opts) => {
      const content = bodyText(opts) ?? '';
      const client = getClient();
      const teamId = await resolve(client, 'team', teamArg);
      const channelId = await resolve(client, 'channel', channelArg, teamId);
      const payload = { body: { contentType: opts.html ? 'html' : 'text', content } };
      const request = { team: teamArg, channel: channelArg, messageId, payload };
      const sent = await idempotentSend(client, opts.idempotencyKey, 'teams reply-channel-message', request, async () => {
        await client.request(
          `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/${encodeURIComponent(messageId)}/replies`,
          { method: 'POST', body: JSON.stringify(payload) }
        );
      });
      if (sent) success('Reply sent to channel thread.');
    }));

  // Chats
  const chat = program.command('chat').description('Teams chat (requires --org-mode)');

  addListOptions(addContentOptions(chat
    .command('list')
    .description('List chats (marked untrusted; --recent: newest activity first, with the last message preview)')
    .option('--recent', 'Order by the last message, newest first, and include its preview (lastMessagePreview)')), { description: 'Number of chats' })
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      // topic is null for 1:1 chats; the table's members column tells them apart (X-13).
      const expand = opts.output === 'table' ? ['members($select=displayName)'] : [];
      // TMS-1: Graph orders /me/chats only by lastMessagePreview/createdDateTime desc.
      if (opts.recent) expand.push('lastMessagePreview');
      if (expand.length) queryParams['$expand'] = expand.join(',');
      if (opts.recent) queryParams['$orderby'] = 'lastMessagePreview/createdDateTime desc';

      await printList(client, '/me/chats', { queryParams }, opts, { transform: chatTransform(client, contentOptions(opts)) });
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

  addListOptions(addContentOptions(chat
    .command('messages <chatId>')
    .description('List chat messages (marked untrusted; see --body, --wrap-untrusted)')), { description: 'Number of messages' })
    .addOption(outputOption())
    .action(handleErrors(async (chatId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;

      const endpoint = `/chats/${encodeURIComponent(chatId)}/messages`;
      await printList(client, endpoint, { queryParams }, opts, { transform: messageTransform(client, 'teams', contentOptions(opts)) });
    }));

  chat
    .command('create')
    .description('Create a 1:1 chat (one member) or group chat (two or more)')
    .requiredOption('--members <emails>', 'Comma-separated UPNs/emails of the other members')
    .option('--topic <topic>', 'Group chat topic')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const upns = String(opts.members).split(',').filter(e => e.trim()).map(validateEmail);
      if (upns.length === 0) throw usage('--members needs at least one email');
      const chatType = upns.length === 1 ? 'oneOnOne' : 'group';
      if (opts.topic && chatType === 'oneOnOne') throw usage('--topic applies to group chats only (two or more members)');
      const client = getClient();
      // The signed-in user must be listed as a member too.
      const me = (await client.request('/me', { queryParams: { $select: 'id' } })).data as { id?: string };
      const member = (user: string) => ({
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${odataQuote(user)}')`,
      });
      const chatBody: Record<string, unknown> = { chatType };
      if (opts.topic) chatBody.topic = opts.topic;
      chatBody.members = [member(me.id ?? '<me>'), ...upns.map(member)];
      const response = await client.request('/chats', { method: 'POST', body: JSON.stringify(chatBody) });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  addBodyOptions(chat
    .command('send <chatId>')
    .description('Send a chat message'), 'Message body', true)
    .option('--html', 'Body is HTML')
    .option('--mention <upn>', '@mention a user by UPN/email (repeatable)', collect, [])
    .option('--idempotency-key <key>', IDEMPOTENCY_HELP)
    .action(handleErrors(async (chatId: string, opts) => {
      const content = bodyText(opts) ?? '';
      const upns = mentionUpns(opts.mention);
      const client = getClient();
      const payload = await messagePayload(client, content, !!opts.html, upns);
      const sent = await idempotentSend(client, opts.idempotencyKey, 'chat send', { chat: chatId, payload }, async () => {
        await client.request(`/chats/${encodeURIComponent(chatId)}/messages`, { method: 'POST', body: JSON.stringify(payload) });
      });
      if (sent) success('Message sent.');
    }));

  addBodyOptions(chat
    .command('reply <chatId> <messageId>')
    .description('Reply to a chat message'), 'Reply body', true)
    .option('--html', 'Body is HTML')
    .option('--idempotency-key <key>', IDEMPOTENCY_HELP)
    .action(handleErrors(async (chatId: string, messageId: string, opts) => {
      const content = bodyText(opts) ?? '';
      const client = getClient();
      // Chats have no /replies endpoint; Graph's documented "reply with quote" is a new
      // message carrying a messageReference attachment to the original. The key covers the
      // caller's input, not the quoted preview read from Graph.
      const request = { chat: chatId, messageId, body: content, html: !!opts.html };
      const sent = await idempotentSend(client, opts.idempotencyKey, 'chat reply', request, async () => {
        const messagesEndpoint = `/chats/${encodeURIComponent(chatId)}/messages`;
        const original = (await client.request(`${messagesEndpoint}/${encodeURIComponent(messageId)}`)).data as {
          body?: { content?: string };
          from?: unknown;
        };
        const replyHtml = opts.html ? content : escapeHtml(content);
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
      });
      if (sent) success('Reply sent.');
    }));
}

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const usage = (msg: string) => new CliError(msg, { code: 'usage', exitCode: ExitCode.USAGE });

function collect(value: string, previous: string[]): string[] {
  return [...previous, value];
}

function mentionUpns(values: string[] | undefined): string[] {
  return (values ?? []).map(validateEmail);
}

/**
 * chatMessage body with @mentions: each mentioned user becomes `<at id="N">displayName</at>`
 * in an HTML body (replacing the first "@displayName" in the text, else appended) plus a
 * matching `mentions[N]` entry. Without mentions the body is sent as given.
 */
async function messagePayload(client: GraphClient, text: string, html: boolean, upns: string[]): Promise<Record<string, unknown>> {
  if (upns.length === 0) return { body: { contentType: html ? 'html' : 'text', content: text } };
  let content = html ? text : escapeHtml(text);
  const mentions = [];
  for (const [n, upn] of upns.entries()) {
    const user = (await client.request(`/users/${encodeURIComponent(upn)}`, { queryParams: { $select: 'id,displayName' } })).data as
      { id?: string; displayName?: string };
    if (!user.id && !client.isDryRun()) {
      throw new CliError(`No user found for ${upn}`, { code: 'not_found', exitCode: ExitCode.NOT_FOUND });
    }
    const displayName = user.displayName ?? upn;
    const tag = `<at id="${n}">${escapeHtml(displayName)}</at>`;
    const needle = `@${escapeHtml(displayName)}`;
    const at = content.indexOf(needle);
    content = at >= 0 ? content.slice(0, at) + tag + content.slice(at + needle.length) : `${content} ${tag}`;
    mentions.push({
      id: n,
      mentionText: displayName,
      mentioned: { user: { id: user.id ?? upn, displayName, userIdentityType: 'aadUser' } },
    });
  }
  return { body: { contentType: 'html', content }, mentions };
}
