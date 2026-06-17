import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, streamPages, type OutputFormat } from '../formatter.js';
import { handleErrors } from '../errors.js';
import { parseDate, truncate, parseRecipients, formatSize, outputOption, confirm, info } from '../utils.js';

// Well-known folder names that Graph API supports directly
const WELL_KNOWN_FOLDERS = [
  'inbox', 'drafts', 'sentitems', 'deleteditems', 'archive',
  'junkemail', 'outbox', 'clutter', 'conflicts',
  'conversationhistory', 'localfailures', 'msgfolderroot',
  'recoverableitemsdeletions', 'scheduled', 'searchfolders', 'serverfailures',
];

/**
 * Resolve a folder display name to its ID.
 * First tries well-known folder names (case-insensitive), then searches by displayName.
 */
async function resolveFolderName(client: GraphClient, name: string): Promise<string | null> {
  // Check if it's a well-known name
  const normalized = name.toLowerCase().replace(/\s+/g, '');
  const wellKnown = WELL_KNOWN_FOLDERS.find(f => f === normalized);
  if (wellKnown) {
    // Graph API accepts well-known names directly as folder IDs
    return wellKnown;
  }

  // Search by displayName — validate input to prevent OData filter injection
  const safeName = name.replace(/[^a-zA-Z0-9 _\-().]/g, '').trim();
  if (!safeName) return null;
  const response = await client.request('/me/mailFolders', {
    queryParams: {
      '$filter': `displayName eq '${safeName.replace(/'/g, "''")}'`,
      '$select': 'id,displayName',
      '$top': '1',
    },
  });
  const data = response.data as Record<string, unknown>;
  const folders = (data.value as Array<Record<string, unknown>>) || [];
  if (folders.length > 0) {
    return folders[0].id as string;
  }

  // Fallback: try case-insensitive match on all folders
  const allResponse = await client.request('/me/mailFolders', {
    queryParams: { '$select': 'id,displayName', '$top': '100' },
  });
  const allData = allResponse.data as Record<string, unknown>;
  const allFolders = (allData.value as Array<Record<string, unknown>>) || [];
  const match = allFolders.find(
    f => (f.displayName as string || '').toLowerCase() === name.toLowerCase()
  );
  return match ? (match.id as string) : null;
}

export function registerMailCommands(program: Command, getClient: () => GraphClient): void {
  const mail = program.command('mail').description('Outlook mail operations');

  mail
    .command('list')
    .description('List inbox messages')
    .option('--top <n>', 'Number of messages to return', '10')
    .option('--filter <filter>', 'OData filter expression')
    .option('--select <fields>', 'Comma-separated fields to return')
    .option('--search <query>', 'Search query')
    .option('--folder <name>', 'Folder name or ID (e.g. "Drafts", "Sent Items", "Archive", or a folder ID)')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.select) queryParams['$select'] = opts.select;
      if (opts.search) queryParams['$search'] = `"${opts.search}"`;

      let endpoint = '/me/messages';
      if (opts.folder) {
        // Try as well-known name first, then resolve
        const folderId = await resolveFolderName(client, opts.folder) || opts.folder;
        endpoint = `/me/mailFolders/${encodeURIComponent(folderId)}/messages`;
      }

      if (opts.all && (opts.output === 'json' || opts.output === 'csv')) {
        // Stream page-by-page for json/csv to avoid buffering all results in memory
        const pages = client.requestPagesStreaming(
          endpoint,
          { queryParams },
          opts.pageLimit ? parseInt(opts.pageLimit) : undefined,
          opts.pageDelay ? parseInt(opts.pageDelay) : undefined,
        );
        await streamPages(pages, opts.output as OutputFormat);
      } else {
        const response = opts.all
          ? await client.requestAllPages(
              endpoint,
              { queryParams },
              opts.pageLimit ? parseInt(opts.pageLimit) : undefined,
              opts.pageDelay ? parseInt(opts.pageDelay) : undefined,
            )
          : await client.request(endpoint, { queryParams });
        console.log(formatOutput(response.data, opts.output as OutputFormat));
      }
    }));

  mail
    .command('get <messageId>')
    .description('Get a specific message (raw API response)')
    .option('--select <fields>', 'Comma-separated fields to return')
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.select) queryParams['$select'] = opts.select;

      const response = await client.request(`/me/messages/${encodeURIComponent(messageId)}`, { queryParams });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('read <messageId>')
    .description('Read a message with full body, attachments info, and metadata')
    .option('--text', 'Prefer plain text body over HTML')
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {
        '$select': 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,hasAttachments,importance,categories,flag,isRead,conversationId,internetMessageId',
      };
      const headers: Record<string, string> = {};
      if (opts.text) {
        headers['Prefer'] = 'outlook.body-content-type="text"';
      }

      const response = await client.request(
        `/me/messages/${encodeURIComponent(messageId)}`,
        { queryParams, headers }
      );

      const msg = response.data as Record<string, unknown>;

      if (opts.output === 'json') {
        console.log(formatOutput(response.data, 'json'));
      } else {
        // Human-readable output
        const from = (msg.from as Record<string, unknown>)?.emailAddress as Record<string, unknown>;
        const to = (msg.toRecipients as Array<Record<string, unknown>>)?.map(
          (r) => ((r.emailAddress as Record<string, unknown>)?.address || '') as string
        ).join(', ');
        const cc = (msg.ccRecipients as Array<Record<string, unknown>>)?.map(
          (r) => ((r.emailAddress as Record<string, unknown>)?.address || '') as string
        ).join(', ');
        const flag = msg.flag as Record<string, unknown>;
        const categories = msg.categories as string[] || [];
        const body = msg.body as Record<string, unknown>;

        const lines: string[] = [];
        lines.push(`Subject:    ${msg.subject || '(no subject)'}`);
        lines.push(`From:       ${from?.name || ''} <${from?.address || ''}>`);
        lines.push(`To:         ${to}`);
        if (cc) lines.push(`CC:         ${cc}`);
        lines.push(`Date:       ${msg.receivedDateTime}`);
        lines.push(`Read:       ${msg.isRead ? 'Yes' : 'No'}`);
        lines.push(`Importance: ${msg.importance || 'normal'}`);
        if (msg.hasAttachments) lines.push(`Attachments: Yes`);
        if (categories.length > 0) lines.push(`Categories: ${categories.join(', ')}`);
        if (flag?.flagStatus && flag.flagStatus !== 'notFlagged') {
          lines.push(`Flag:       ${flag.flagStatus}`);
        }
        lines.push(`---`);
        if (body?.content) {
          let content = body.content as string;
          // Strip HTML tags for text output if not --text mode
          if (body.contentType === 'html' && !opts.text) {
            content = content.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/\n\s*\n/g, '\n').trim();
          }
          lines.push(content);
        }
        console.log(lines.join('\n'));
      }

      // Also show attachment list if present
      if (msg.hasAttachments) {
        const attResponse = await client.request(
          `/me/messages/${encodeURIComponent(messageId)}/attachments`,
          { queryParams: { '$select': 'id,name,contentType,size' } }
        );
        const attData = attResponse.data as Record<string, unknown>;
        const attachments = attData.value as Array<Record<string, unknown>>;
        if (attachments && attachments.length > 0 && opts.output !== 'json') {
          console.log(`\nAttachments (${attachments.length}):`);
          for (const att of attachments) {
            const sizeKb = Math.round((att.size as number || 0) / 1024);
            console.log(`  - ${att.name} (${att.contentType}, ${sizeKb} KB) [${att.id}]`);
          }
        }
      }
    }));

  mail
    .command('triage')
    .description('Show inbox summary — unread messages with key metadata')
    .option('--top <n>', 'Number of messages', '20')
    .option('--filter <filter>', 'OData filter (default: unread)', 'isRead eq false')
    .option('--folder <folderId>', 'Mail folder ID')
    .option('--include-read', 'Include read messages too')
    .addOption(outputOption('table'))
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {
        '$top': opts.top,
        '$select': 'id,subject,from,receivedDateTime,hasAttachments,importance,categories,flag,isRead,bodyPreview',
        '$orderby': 'receivedDateTime desc',
      };
      if (!opts.includeRead) {
        queryParams['$filter'] = opts.filter;
      }

      const endpoint = opts.folder
        ? `/me/mailFolders/${encodeURIComponent(opts.folder)}/messages`
        : '/me/messages';

      const response = await client.request(endpoint, { queryParams });
      const data = response.data as Record<string, unknown>;
      const messages = (data.value as Array<Record<string, unknown>>) || [];

      if (messages.length === 0) {
        console.log('No messages found.');
        return;
      }

      if (opts.output === 'json') {
        console.log(formatOutput(response.data, 'json'));
        return;
      }

      // Build a clean summary
      const rows = messages.map((msg) => {
        const from = (msg.from as Record<string, unknown>)?.emailAddress as Record<string, unknown>;
        const flag = msg.flag as Record<string, unknown>;
        const categories = (msg.categories as string[]) || [];
        const date = new Date(msg.receivedDateTime as string);
        const dateStr = date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });

        const indicators: string[] = [];
        if (msg.hasAttachments) indicators.push('[A]');
        if (msg.importance === 'high') indicators.push('[!]');
        if (flag?.flagStatus === 'flagged') indicators.push('[F]');
        if (categories.length > 0) indicators.push(`[${categories[0]}]`);

        return {
          date: dateStr,
          from: truncate(from?.name as string || from?.address as string || '', 25),
          subject: truncate(msg.subject as string || '(no subject)', 50),
          flags: indicators.join(' '),
          id: (msg.id as string || '').substring(0, 12) + '...',
        };
      });

      if (opts.output === 'csv') {
        console.log(formatOutput(rows, 'csv'));
      } else {
        // Table output
        console.log(`Inbox: ${messages.length} message(s)  [A]=attachment [!]=high priority [F]=flagged\n`);
        console.log(formatOutput(rows, 'table'));
      }
    }));

  mail
    .command('send')
    .description('Send an email')
    .requiredOption('--to <addresses>', 'Comma-separated recipient email addresses')
    .option('--cc <addresses>', 'Comma-separated CC addresses')
    .option('--bcc <addresses>', 'Comma-separated BCC addresses')
    .requiredOption('--subject <subject>', 'Email subject')
    .requiredOption('--body <body>', 'Email body')
    .option('--html', 'Send body as HTML')
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const toRecipients = parseRecipients(opts.to);

      const message: Record<string, unknown> = {
        subject: opts.subject,
        body: {
          contentType: opts.html ? 'HTML' : 'Text',
          content: opts.body,
        },
        toRecipients,
      };

      if (opts.cc) {
        message.ccRecipients = parseRecipients(opts.cc);
      }

      if (opts.bcc) {
        message.bccRecipients = parseRecipients(opts.bcc);
      }

      await client.request('/me/sendMail', {
        method: 'POST',
        body: JSON.stringify({ message, saveToSentItems: true }),
      });

      info('Email sent successfully.');
    }));

  mail
    .command('delete <messageId>')
    .description('Delete a message')
    .action(handleErrors(async (messageId: string) => {
      if (!await confirm(`Delete message ${messageId.substring(0, 16)}...?`)) return;
      const client = getClient();
      await client.request(`/me/messages/${encodeURIComponent(messageId)}`, { method: 'DELETE' });
      info('Message deleted.');
    }));

  mail
    .command('move <messageId>')
    .description('Move a message to a folder')
    .requiredOption('--folder <folderId>', 'Destination folder ID')
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      await client.request(`/me/messages/${encodeURIComponent(messageId)}/move`, {
        method: 'POST',
        body: JSON.stringify({ destinationId: opts.folder }),
      });
      info('Message moved.');
    }));

  mail
    .command('folders')
    .description('List mail folders')
    .option('--parent <folderId>', 'List child folders of this folder')
    .option('--all', 'Fetch all pages (including nested child folders)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const endpoint = opts.parent
        ? `/me/mailFolders/${encodeURIComponent(opts.parent)}/childFolders`
        : '/me/mailFolders';
      const queryParams: Record<string, string> = {
        '$select': 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount',
        '$top': '100',
      };
      const response = opts.all
        ? await client.requestAllPages(endpoint, { queryParams })
        : await client.request(endpoint, { queryParams });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('folder <name>')
    .description('Find a folder by display name (e.g. "Drafts", "Sent Items", "Archive")')
    .addOption(outputOption())
    .action(handleErrors(async (folderName: string, opts) => {
      const client = getClient();
      const folderId = await resolveFolderName(client, folderName);
      if (!folderId) {
        throw new Error(`Folder "${folderName}" not found. Run: ms365 mail folders -o table`);
      }
      const response = await client.request(`/me/mailFolders/${encodeURIComponent(folderId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('create-folder')
    .description('Create a mail folder')
    .requiredOption('--name <name>', 'Folder display name')
    .option('--parent <parentFolderId>', 'Parent folder ID (default: root)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const endpoint = opts.parent
        ? `/me/mailFolders/${encodeURIComponent(opts.parent)}/childFolders`
        : '/me/mailFolders';
      const response = await client.request(endpoint, {
        method: 'POST',
        body: JSON.stringify({ displayName: opts.name }),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // ── Drafts ────────────────────────────────────────────────────
  mail
    .command('draft')
    .description('Create a draft email')
    .requiredOption('--to <addresses>', 'Comma-separated recipient email addresses')
    .requiredOption('--subject <subject>', 'Email subject')
    .requiredOption('--body <body>', 'Email body')
    .option('--cc <addresses>', 'Comma-separated CC addresses')
    .option('--html', 'Send body as HTML')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const toRecipients = parseRecipients(opts.to);

      const message: Record<string, unknown> = {
        subject: opts.subject,
        body: {
          contentType: opts.html ? 'HTML' : 'Text',
          content: opts.body,
        },
        toRecipients,
      };

      if (opts.cc) {
        message.ccRecipients = parseRecipients(opts.cc);
      }

      const response = await client.request('/me/messages', {
        method: 'POST',
        body: JSON.stringify(message),
      });

      const created = response.data as Record<string, unknown>;
      if (opts.output !== 'json') {
        console.log(`Draft created. ID: ${created.id}`);
      } else {
        console.log(formatOutput(response.data, 'json'));
      }
    }));

  mail
    .command('drafts')
    .description('List draft emails')
    .option('--top <n>', 'Number of drafts', '10')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      // Drafts folder has well-known name "Drafts"
      const folderId = await resolveFolderName(client, 'Drafts');
      if (!folderId) {
        throw new Error('Drafts folder not found.');
      }
      const response = await client.request(
        `/me/mailFolders/${encodeURIComponent(folderId)}/messages`,
        {
          queryParams: {
            '$top': opts.top,
            '$select': 'id,subject,toRecipients,createdDateTime,lastModifiedDateTime,bodyPreview,hasAttachments',
            '$orderby': 'lastModifiedDateTime desc',
          },
        }
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('update-draft <messageId>')
    .description('Update an existing draft')
    .option('--to <addresses>', 'Comma-separated recipient emails')
    .option('--cc <addresses>', 'Comma-separated CC addresses')
    .option('--subject <subject>', 'Email subject')
    .option('--body <body>', 'Email body')
    .option('--html', 'Body is HTML')
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      const patch: Record<string, unknown> = {};

      if (opts.subject) patch.subject = opts.subject;
      if (opts.body) {
        patch.body = { contentType: opts.html ? 'HTML' : 'Text', content: opts.body };
      }
      if (opts.to) {
        patch.toRecipients = parseRecipients(opts.to);
      }
      if (opts.cc) {
        patch.ccRecipients = parseRecipients(opts.cc);
      }

      const response = await client.request(`/me/messages/${encodeURIComponent(messageId)}`, {
        method: 'PATCH',
        body: JSON.stringify(patch),
      });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('send-draft <messageId>')
    .description('Send an existing draft')
    .action(handleErrors(async (messageId: string) => {
      const client = getClient();
      await client.request(`/me/messages/${encodeURIComponent(messageId)}/send`, {
        method: 'POST',
      });
      info('Draft sent.');
    }));

  mail
    .command('attachments <messageId>')
    .description('List attachments for a message')
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/me/messages/${encodeURIComponent(messageId)}/attachments`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('attachment <messageId> <attachmentId>')
    .description('Get a specific attachment')
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, attachmentId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('add-attachment <messageId>')
    .description('Add an attachment to a draft message')
    .requiredOption('--name <name>', 'Attachment file name')
    .requiredOption('--content-type <type>', 'MIME content type')
    .requiredOption('--content <base64>', 'Base64-encoded content')
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/me/messages/${encodeURIComponent(messageId)}/attachments`,
        {
          method: 'POST',
          body: JSON.stringify({
            '@odata.type': '#microsoft.graph.fileAttachment',
            name: opts.name,
            contentType: opts.contentType,
            contentBytes: opts.content,
          }),
        }
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('delete-attachment <messageId> <attachmentId>')
    .description('Delete an attachment')
    .action(handleErrors(async (messageId: string, attachmentId: string) => {
      const client = getClient();
      await client.request(
        `/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
        { method: 'DELETE' }
      );
      info('Attachment deleted.');
    }));

  mail
    .command('download-attachment <messageId> <attachmentId>')
    .description('Download an attachment to a file')
    .option('--out <path>', 'Output file path (default: attachment name)')
    .action(handleErrors(async (messageId: string, attachmentId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`
      );
      const data = response.data as Record<string, unknown>;
      const name = data.name as string || 'attachment';
      const contentBytes = data.contentBytes as string;
      if (!contentBytes) {
        throw new Error('No content found. This may be a reference attachment or item attachment.');
      }
      const outPath = (await import('path')).resolve(opts.out || name);
      (await import('fs')).writeFileSync(outPath, Buffer.from(contentBytes, 'base64'));
      info(`Saved to ${outPath} (${formatSize(Buffer.from(contentBytes, 'base64').length)})`);
    }));

  mail
    .command('reply <messageId>')
    .description('Reply to a message')
    .requiredOption('--body <body>', 'Reply body')
    .option('--html', 'Send body as HTML')
    .option('--reply-all', 'Reply to all recipients')
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      const action = opts.replyAll ? 'replyAll' : 'reply';
      await client.request(`/me/messages/${encodeURIComponent(messageId)}/${action}`, {
        method: 'POST',
        body: JSON.stringify({
          comment: opts.body,
        }),
      });
      info('Reply sent.');
    }));

  // ── Forward ──────────────────────────────────────────────────
  mail
    .command('forward <messageId>')
    .description('Forward a message')
    .requiredOption('--to <addresses>', 'Comma-separated recipient emails')
    .option('--comment <text>', 'Comment to include with forward')
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      const toRecipients = parseRecipients(opts.to);
      await client.request(`/me/messages/${encodeURIComponent(messageId)}/forward`, {
        method: 'POST',
        body: JSON.stringify({ comment: opts.comment || '', toRecipients }),
      });
      info('Message forwarded.');
    }));

  // ── Mark read/unread ────────────────────────────────────────
  mail
    .command('mark-read <messageId>')
    .description('Mark a message as read')
    .action(handleErrors(async (messageId: string) => {
      const client = getClient();
      await client.request(`/me/messages/${encodeURIComponent(messageId)}`, {
        method: 'PATCH',
        body: JSON.stringify({ isRead: true }),
      });
      info('Marked as read.');
    }));

  mail
    .command('mark-unread <messageId>')
    .description('Mark a message as unread')
    .action(handleErrors(async (messageId: string) => {
      const client = getClient();
      await client.request(`/me/messages/${encodeURIComponent(messageId)}`, {
        method: 'PATCH',
        body: JSON.stringify({ isRead: false }),
      });
      info('Marked as unread.');
    }));

  // ── Flag / unflag ───────────────────────────────────────────
  mail
    .command('flag <messageId>')
    .description('Set follow-up flag on a message')
    .option('--due <date>', 'Due date (YYYY-MM-DD)')
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      const flag: Record<string, unknown> = { flagStatus: 'flagged' };
      if (opts.due) {
        flag.dueDateTime = { dateTime: parseDate(opts.due) + 'T00:00:00', timeZone: 'UTC' };
      }
      await client.request(`/me/messages/${encodeURIComponent(messageId)}`, {
        method: 'PATCH',
        body: JSON.stringify({ flag }),
      });
      info('Message flagged.');
    }));

  mail
    .command('unflag <messageId>')
    .description('Remove follow-up flag from a message')
    .action(handleErrors(async (messageId: string) => {
      const client = getClient();
      await client.request(`/me/messages/${encodeURIComponent(messageId)}`, {
        method: 'PATCH',
        body: JSON.stringify({ flag: { flagStatus: 'notFlagged' } }),
      });
      info('Flag removed.');
    }));

  mail
    .command('complete <messageId>')
    .description('Mark a flagged message as complete')
    .action(handleErrors(async (messageId: string) => {
      const client = getClient();
      await client.request(`/me/messages/${encodeURIComponent(messageId)}`, {
        method: 'PATCH',
        body: JSON.stringify({ flag: { flagStatus: 'complete' } }),
      });
      info('Flag marked complete.');
    }));

  // ── Categories ──────────────────────────────────────────────
  mail
    .command('categorize <messageId>')
    .description('Set categories on a message')
    .requiredOption('--categories <names>', 'Comma-separated category names (e.g. "Red category,Blue category")')
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      const categories = opts.categories.split(',').map((c: string) => c.trim());
      await client.request(`/me/messages/${encodeURIComponent(messageId)}`, {
        method: 'PATCH',
        body: JSON.stringify({ categories }),
      });
      console.log(`Categories set: ${categories.join(', ')}`);
    }));

  // Shared mailbox commands (require --org-mode)
  mail
    .command('shared-list <userId>')
    .description('List shared mailbox messages (requires --org-mode)')
    .option('--top <n>', 'Number of messages', '10')
    .option('--filter <filter>', 'OData filter')
    .option('--select <fields>', 'Fields to return')
    .option('--folder <folderId>', 'Mail folder ID')
    .option('--all', 'Fetch all pages')
    .addOption(outputOption())
    .action(handleErrors(async (userId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.select) queryParams['$select'] = opts.select;

      const endpoint = opts.folder
        ? `/users/${encodeURIComponent(userId)}/mailFolders/${encodeURIComponent(opts.folder)}/messages`
        : `/users/${encodeURIComponent(userId)}/messages`;

      const response = opts.all
        ? await client.requestAllPages(endpoint, { queryParams })
        : await client.request(endpoint, { queryParams });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('shared-get <userId> <messageId>')
    .description('Get a shared mailbox message (requires --org-mode)')
    .option('--select <fields>', 'Fields to return')
    .addOption(outputOption())
    .action(handleErrors(async (userId: string, messageId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.select) queryParams['$select'] = opts.select;

      const response = await client.request(
        `/users/${encodeURIComponent(userId)}/messages/${encodeURIComponent(messageId)}`,
        { queryParams }
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('shared-send <userId>')
    .description('Send email from a shared mailbox (requires --org-mode)')
    .requiredOption('--to <addresses>', 'Comma-separated recipient emails')
    .requiredOption('--subject <subject>', 'Email subject')
    .requiredOption('--body <body>', 'Email body')
    .option('--html', 'Body as HTML')
    .action(handleErrors(async (userId: string, opts) => {
      const client = getClient();
      const toRecipients = parseRecipients(opts.to);

      await client.request(`/users/${encodeURIComponent(userId)}/sendMail`, {
        method: 'POST',
        body: JSON.stringify({
          message: {
            subject: opts.subject,
            body: { contentType: opts.html ? 'HTML' : 'Text', content: opts.body },
            toRecipients,
          },
          saveToSentItems: true,
        }),
      });
      console.log('Email sent from shared mailbox.');
    }));

  // ── Inbox rules ─────────────────────────────────────────────
  mail
    .command('rules')
    .description('List inbox rules')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/mailFolders/inbox/messageRules');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('rule <ruleId>')
    .description('Get a specific inbox rule')
    .addOption(outputOption())
    .action(handleErrors(async (ruleId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('create-rule')
    .description('Create an inbox rule')
    .requiredOption('--name <name>', 'Rule display name')
    .option('--from <addresses>', 'Comma-separated sender addresses to match')
    .option('--subject-contains <words>', 'Comma-separated words to match in subject')
    .option('--move-to <folderId>', 'Move matching messages to folder ID')
    .option('--forward-to <addresses>', 'Comma-separated addresses to forward to')
    .option('--delete', 'Delete matching messages')
    .option('--mark-read', 'Mark matching messages as read')
    .option('--importance <level>', 'Set importance: low, normal, high')
    .option('--stop-processing', 'Stop processing subsequent rules')
    .option('--enabled', 'Enable rule (default: true)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();

      const conditions: Record<string, unknown> = {};
      if (opts.from) {
        conditions.senderContains = opts.from.split(',').map((s: string) => s.trim());
      }
      if (opts.subjectContains) {
        conditions.subjectContains = opts.subjectContains.split(',').map((s: string) => s.trim());
      }

      const actions: Record<string, unknown> = {};
      if (opts.moveTo) {
        actions.moveToFolder = opts.moveTo;
      }
      if (opts.forwardTo) {
        actions.forwardTo = parseRecipients(opts.forwardTo);
      }
      if (opts.delete) {
        actions.delete = true;
      }
      if (opts.markRead) {
        actions.markAsRead = true;
      }
      if (opts.importance) {
        actions.markImportance = opts.importance;
      }
      if (opts.stopProcessing) {
        actions.stopProcessingRules = true;
      }

      const rule: Record<string, unknown> = {
        displayName: opts.name,
        conditions,
        actions,
        isEnabled: opts.enabled !== false,
      };

      const response = await client.request('/me/mailFolders/inbox/messageRules', {
        method: 'POST',
        body: JSON.stringify(rule),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('update-rule <ruleId>')
    .description('Update an inbox rule')
    .option('--name <name>', 'Rule display name')
    .option('--enabled <bool>', 'Enable/disable rule (true/false)')
    .addOption(outputOption())
    .action(handleErrors(async (ruleId: string, opts) => {
      const client = getClient();
      const patch: Record<string, unknown> = {};
      if (opts.name) patch.displayName = opts.name;
      if (opts.enabled !== undefined) patch.isEnabled = opts.enabled === 'true';

      const response = await client.request(
        `/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`,
        { method: 'PATCH', body: JSON.stringify(patch) }
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));


  // ── Batch operations ────────────────────────────────────────
  mail
    .command('batch-read <messageIds>')
    .description('Mark multiple messages as read (comma-separated IDs)')
    .action(handleErrors(async (messageIds: string) => {
      const client = getClient();
      const ids = messageIds.split(',').map(id => id.trim()).filter(Boolean);
      let succeeded = 0;
      const failed: Array<{ id: string; error: string }> = [];
      for (const id of ids) {
        try {
          await client.request(`/me/messages/${encodeURIComponent(id)}`, {
            method: 'PATCH',
            body: JSON.stringify({ isRead: true }),
          });
          succeeded++;
        } catch (err) {
          failed.push({ id, error: (err as Error).message });
        }
      }
      console.log(`Marked ${succeeded} message(s) as read.`);
      if (failed.length > 0) {
        console.error(`Failed ${failed.length}: ${failed.map(f => f.id.substring(0, 12) + '...').join(', ')}`);
      }
    }));

  mail
    .command('batch-move <messageIds>')
    .description('Move multiple messages to a folder (comma-separated IDs)')
    .requiredOption('--folder <folderId>', 'Destination folder name or ID')
    .action(handleErrors(async (messageIds: string, opts) => {
      const client = getClient();
      const folderId = await resolveFolderName(client, opts.folder) || opts.folder;
      const ids = messageIds.split(',').map(id => id.trim()).filter(Boolean);
      let succeeded = 0;
      const failed: Array<{ id: string; error: string }> = [];
      for (const id of ids) {
        try {
          await client.request(`/me/messages/${encodeURIComponent(id)}/move`, {
            method: 'POST',
            body: JSON.stringify({ destinationId: folderId }),
          });
          succeeded++;
        } catch (err) {
          failed.push({ id, error: (err as Error).message });
        }
      }
      console.log(`Moved ${succeeded} message(s).`);
      if (failed.length > 0) {
        console.error(`Failed ${failed.length}: ${failed.map(f => f.id.substring(0, 12) + '...').join(', ')}`);
      }
    }));

}
