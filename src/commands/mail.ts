import fs from 'node:fs';
import path from 'node:path';
import { Command, InvalidArgumentError } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { handleErrors } from '../errors.js';
import { saveOrStream } from '../files.js';
import { calendarDate, odataQuote, searchQuote } from '../graph-helpers.js';
import { listRequest, paginationOptions } from '../list.js';
import { truncate, parseRecipients, formatSize, outputOption, confirm, info } from '../utils.js';

// Well-known folder names that Graph API supports directly
const WELL_KNOWN_FOLDERS = [
  'inbox', 'drafts', 'sentitems', 'deleteditems', 'archive',
  'junkemail', 'outbox', 'clutter', 'conflicts',
  'conversationhistory', 'localfailures', 'msgfolderroot',
  'recoverableitemsdeletions', 'scheduled', 'searchfolders', 'serverfailures',
];

/** Graph mail folder IDs are long base64 strings; display names of that shape are not looked up. */
const FOLDER_ID = /^[A-Za-z0-9+/=_-]{32,}$/;

const RULES = '/me/mailFolders/inbox/messageRules';

interface MailFolder {
  id: string;
  displayName?: string;
  childFolderCount?: number;
}

/**
 * Resolve a folder name or ID in `mailbox` (`/me` or `/users/{id}`) to a value usable in
 * `mailFolders/{x}` paths and as a move `destinationId`. Well-known names ("Sent Items" -> `sentitems`)
 * and ID-shaped values pass through. Anything else is a display name: an exact OData match among the
 * top-level folders, else a breadth-first, case-insensitive walk of every childFolders level (the
 * shallowest match wins; two matches at that depth are ambiguous). Throws when nothing matches.
 * Dry-run cannot look names up and returns the name unchanged.
 */
async function resolveFolder(client: GraphClient, name: string, mailbox = '/me'): Promise<string> {
  const wellKnown = name.toLowerCase().replace(/\s+/g, '');
  if (WELL_KNOWN_FOLDERS.includes(wellKnown)) return wellKnown;
  if (FOLDER_ID.test(name)) return name;

  const $select = 'id,displayName,childFolderCount';
  const exact = await client.request(`${mailbox}/mailFolders`, {
    queryParams: { $filter: `displayName eq ${odataQuote(name)}`, $select },
  });
  if (exact.status === 0) return name;
  const hit = (exact.data as { value?: MailFolder[] }).value?.[0];
  if (hit) return hit.id;

  const wanted = name.toLowerCase();
  let level = [`${mailbox}/mailFolders`];
  while (level.length > 0) {
    const matches: string[] = [];
    const next: string[] = [];
    for (const endpoint of level) {
      const res = await client.requestAllPages(endpoint, { queryParams: { $select, $top: '100' } });
      for (const folder of (res.data as { value?: MailFolder[] }).value ?? []) {
        if ((folder.displayName ?? '').toLowerCase() === wanted) matches.push(folder.id);
        if ((folder.childFolderCount ?? 0) > 0) {
          next.push(`${mailbox}/mailFolders/${encodeURIComponent(folder.id)}/childFolders`);
        }
      }
    }
    if (matches.length === 1) return matches[0];
    if (matches.length > 1) {
      throw new Error(`Folder name "${name}" is ambiguous (${matches.length} folders). Pass the folder ID from: ms365 mail folders`);
    }
    level = next;
  }
  throw new Error(`Folder "${name}" not found. Run: ms365 mail folders -o table`);
}

/** After a batch: print every failure (full ID + Graph error) and fail the command if there was any. */
function failOnBatchErrors(failed: Array<{ id: string; error: string }>, total: number): void {
  if (failed.length === 0) return;
  for (const f of failed) console.error(`Failed: ${f.id}: ${f.error}`);
  throw new Error(`${failed.length} of ${total} message(s) failed.`);
}

export function registerMailCommands(program: Command, getClient: () => GraphClient): void {
  const mail = program.command('mail').description('Outlook mail operations');

  paginationOptions(
    mail
      .command('list')
      .description('List messages across all folders (use --folder inbox for the inbox only)'),
    10,
  )
    .option('--filter <filter>', 'OData filter expression')
    .option('--select <fields>', 'Comma-separated fields to return')
    .option('--search <query>', 'Search query (KQL, sent as one quoted $search phrase)')
    .option('--folder <name>', 'Folder name or ID (e.g. "Inbox", "Sent Items", "Archive", a nested folder name, or a folder ID)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.select) queryParams['$select'] = opts.select;
      if (opts.search) queryParams['$search'] = searchQuote(opts.search);

      const endpoint = opts.folder
        ? `/me/mailFolders/${encodeURIComponent(await resolveFolder(client, opts.folder))}/messages`
        : '/me/messages';
      await listRequest(client, endpoint, { queryParams }, opts);
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
      let attachments: Array<Record<string, unknown>> = [];
      if (msg.hasAttachments) {
        const attResponse = await client.request(
          `/me/messages/${encodeURIComponent(messageId)}/attachments`,
          { queryParams: { '$select': 'id,name,contentType,size' } }
        );
        attachments = (attResponse.data as { value?: Array<Record<string, unknown>> }).value ?? [];
      }

      if (opts.output === 'json' || opts.output === 'yaml') {
        console.log(formatOutput({ ...msg, attachments }, opts.output));
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
        if (attachments.length > 0) {
          lines.push('', `Attachments (${attachments.length}):`);
          for (const att of attachments) {
            const sizeKb = Math.round((att.size as number || 0) / 1024);
            lines.push(`  - ${att.name} (${att.contentType}, ${sizeKb} KB) [${att.id}]`);
          }
        }
        console.log(lines.join('\n'));
      }
    }));

  paginationOptions(
    mail
      .command('triage')
      .description('Show inbox summary — unread inbox messages with key metadata'),
    20,
  )
    .option('--filter <filter>', 'OData filter (default: isRead eq false)')
    .option('--folder <name>', 'Folder name or ID (default: inbox)')
    .option('--include-read', 'Include read messages too (drops the default unread filter, keeps --filter)')
    .addOption(outputOption('table'))
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const filter: string | undefined = opts.filter ?? (opts.includeRead ? undefined : 'isRead eq false');
      // Graph rejects $orderby on a property that does not lead the $filter (InefficientFilter).
      const queryParams: Record<string, string> = {
        '$select': 'id,subject,from,receivedDateTime,hasAttachments,importance,categories,flag,isRead,bodyPreview',
        '$orderby': 'receivedDateTime desc',
        '$filter': `receivedDateTime ge 1900-01-01T00:00:00Z${filter ? ` and (${filter})` : ''}`,
      };
      const folder = opts.folder ? await resolveFolder(client, opts.folder) : 'inbox';
      const endpoint = `/me/mailFolders/${encodeURIComponent(folder)}/messages`;

      if (opts.output === 'json') {
        await listRequest(client, endpoint, { queryParams }, opts);
        return;
      }

      const request = { queryParams: { ...queryParams, '$top': String(opts.top) } };
      const response = opts.all
        ? await client.requestAllPages(endpoint, request, opts.pageLimit, opts.pageDelay)
        : await client.request(endpoint, request);
      const messages = (response.data as { value?: Array<Record<string, unknown>> }).value ?? [];
      const table = opts.output === 'table';

      // Build a clean summary; only the table shortens dates and IDs for width.
      const rows = messages.map((msg) => {
        const from = (msg.from as Record<string, unknown>)?.emailAddress as Record<string, unknown>;
        const flag = msg.flag as Record<string, unknown>;
        const categories = (msg.categories as string[]) || [];
        const id = msg.id as string || '';

        const indicators: string[] = [];
        if (msg.hasAttachments) indicators.push('[A]');
        if (msg.importance === 'high') indicators.push('[!]');
        if (flag?.flagStatus === 'flagged') indicators.push('[F]');
        if (categories.length > 0) indicators.push(`[${categories[0]}]`);

        return {
          date: table
            ? new Date(msg.receivedDateTime as string).toLocaleDateString('en-US', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })
            : msg.receivedDateTime,
          from: truncate(from?.name as string || from?.address as string || '', 25),
          subject: truncate(msg.subject as string || '(no subject)', 50),
          flags: indicators.join(' '),
          id: table ? id.substring(0, 12) + '...' : id,
        };
      });

      if (table && rows.length > 0) {
        info(`${messages.length} message(s)  [A]=attachment [!]=high priority [F]=flagged`);
      }
      console.log(formatOutput(rows, opts.output as OutputFormat));
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
      await confirm(`Delete message ${messageId.substring(0, 16)}...?`);
      const client = getClient();
      await client.request(`/me/messages/${encodeURIComponent(messageId)}`, { method: 'DELETE' });
      info('Message deleted.');
    }));

  mail
    .command('move <messageId>')
    .description('Move a message to a folder')
    .requiredOption('--folder <name>', 'Destination folder name or ID')
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      const destinationId = await resolveFolder(client, opts.folder);
      await client.request(`/me/messages/${encodeURIComponent(messageId)}/move`, {
        method: 'POST',
        body: JSON.stringify({ destinationId }),
      });
      info('Message moved.');
    }));

  paginationOptions(mail.command('folders').description('List mail folders'), 100)
    .option('--parent <folderId>', 'List child folders of this folder')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const endpoint = opts.parent
        ? `/me/mailFolders/${encodeURIComponent(opts.parent)}/childFolders`
        : '/me/mailFolders';
      await listRequest(client, endpoint, {
        queryParams: { '$select': 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount' },
      }, opts);
    }));

  mail
    .command('folder <name>')
    .description('Find a folder by display name (e.g. "Drafts", "Sent Items", "Archive")')
    .addOption(outputOption())
    .action(handleErrors(async (folderName: string, opts) => {
      const client = getClient();
      const folderId = await resolveFolder(client, folderName);
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

      info('Draft created.');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  paginationOptions(mail.command('drafts').description('List draft emails'), 10)
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      await listRequest(getClient(), '/me/mailFolders/drafts/messages', {
        queryParams: {
          '$select': 'id,subject,toRecipients,createdDateTime,lastModifiedDateTime,bodyPreview,hasAttachments',
          '$orderby': 'lastModifiedDateTime desc',
        },
      }, opts);
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

  paginationOptions(mail.command('attachments <messageId>').description('List attachments for a message'))
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts) => {
      await listRequest(getClient(), `/me/messages/${encodeURIComponent(messageId)}/attachments`, {}, opts);
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
      await confirm(`Delete attachment ${attachmentId.substring(0, 16)}...?`);
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
    .option('--out <path>', 'Output file, or an existing directory (default: the attachment name in the current directory)')
    .option('--force', 'Overwrite an existing file')
    .action(handleErrors(async (messageId: string, attachmentId: string, opts) => {
      const client = getClient();
      const endpoint = `/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`;
      const meta = (await client.request(endpoint, { queryParams: { '$select': 'id,name,contentType,size' } })).data as Record<string, unknown>;
      if (meta['@odata.type'] === '#microsoft.graph.referenceAttachment') {
        throw new Error('This is a reference attachment (a link to a cloud file); it has no content to download.');
      }

      let out: string | undefined = opts.out;
      const intoDir = out !== undefined && fs.statSync(out, { throwIfNoEntry: false })?.isDirectory() === true;
      if (out === undefined || intoDir) {
        // The name comes from the sender: keep only its last path segment so it stays in the target directory.
        const name = path.basename(String(meta.name ?? 'attachment').replace(/\\/g, '/'));
        if (name === '' || name === '.' || name === '..') {
          throw new Error(`Unsafe attachment name "${meta.name}". Choose a file name with --out <file>.`);
        }
        out = out === undefined ? name : path.join(out, name);
      }

      const response = await client.requestBinary(`${endpoint}/$value`);
      const saved = await saveOrStream(response.body, { out, force: opts.force });
      if (saved.path) info(`Saved to ${path.resolve(saved.path)} (${formatSize(saved.bytes)})`);
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
      // Graph takes either a plain-text comment or a message body, never both.
      const payload = opts.html
        ? { message: { body: { contentType: 'HTML', content: opts.body } } }
        : { comment: opts.body };
      await client.request(`/me/messages/${encodeURIComponent(messageId)}/${action}`, {
        method: 'POST',
        body: JSON.stringify(payload),
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
    .option('--start <date>', 'Start date (YYYY-MM-DD; with --due default: today, or the due date if earlier)')
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      const flag: Record<string, unknown> = { flagStatus: 'flagged' };
      const due = opts.due ? calendarDate(opts.due) : undefined;
      // Graph rejects a dueDateTime without a startDateTime (400).
      const today = new Date().toISOString().slice(0, 10);
      const start = opts.start ? calendarDate(opts.start) : due && (due < today ? due : today);
      if (start) flag.startDateTime = { dateTime: `${start}T00:00:00`, timeZone: 'UTC' };
      if (due) flag.dueDateTime = { dateTime: `${due}T00:00:00`, timeZone: 'UTC' };
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
      info(`Categories set: ${categories.join(', ')}`);
    }));

  // Shared mailbox commands (require --org-mode)
  paginationOptions(
    mail
      .command('shared-list <userId>')
      .description('List shared mailbox messages (requires --org-mode)'),
    10,
  )
    .option('--filter <filter>', 'OData filter')
    .option('--select <fields>', 'Fields to return')
    .option('--folder <name>', 'Folder name or ID')
    .addOption(outputOption())
    .action(handleErrors(async (userId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.select) queryParams['$select'] = opts.select;

      const mailbox = `/users/${encodeURIComponent(userId)}`;
      const endpoint = opts.folder
        ? `${mailbox}/mailFolders/${encodeURIComponent(await resolveFolder(client, opts.folder, mailbox))}/messages`
        : `${mailbox}/messages`;
      await listRequest(client, endpoint, { queryParams }, opts);
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
      info('Email sent from shared mailbox.');
    }));

  // ── Inbox rules ─────────────────────────────────────────────
  paginationOptions(mail.command('rules').description('List inbox rules'))
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      await listRequest(getClient(), RULES, {}, opts);
    }));

  mail
    .command('rule <ruleId>')
    .description('Get a specific inbox rule')
    .addOption(outputOption())
    .action(handleErrors(async (ruleId: string, opts) => {
      const client = getClient();
      const response = await client.request(`${RULES}/${encodeURIComponent(ruleId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('create-rule')
    .description('Create an inbox rule')
    .requiredOption('--name <name>', 'Rule display name')
    .option('--from <addresses>', 'Comma-separated sender addresses to match (exact addresses)')
    .option('--subject-contains <words>', 'Comma-separated words to match in subject')
    .option('--move-to <folderId>', 'Move matching messages to folder ID')
    .option('--forward-to <addresses>', 'Comma-separated addresses to forward to')
    .option('--delete', 'Delete matching messages')
    .option('--mark-read', 'Mark matching messages as read')
    .option('--importance <level>', 'Set importance: low, normal, high')
    .option('--stop-processing', 'Stop processing subsequent rules')
    .option('--sequence <n>', 'Rule order, lowest runs first (default: after the last existing rule)', (value: string) => {
      if (!/^\d+$/.test(value)) throw new InvalidArgumentError('Expected a non-negative integer.');
      return Number(value);
    })
    .option('--enabled', 'Create the rule enabled (default)')
    .option('--no-enabled', 'Create the rule disabled')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();

      const conditions: Record<string, unknown> = {};
      if (opts.from) {
        conditions.fromAddresses = parseRecipients(opts.from);
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
      if (Object.keys(actions).length === 0) {
        throw new Error('Specify at least one action: --move-to, --forward-to, --delete, --mark-read, --importance or --stop-processing.');
      }
      if (Object.keys(conditions).length === 0) {
        await confirm(`Rule "${opts.name}" has no conditions and applies to every incoming message. Create it?`);
      }

      let sequence: number | undefined = opts.sequence;
      if (sequence === undefined) {
        const existing = await client.requestAllPages(RULES);
        const rules = (existing.data as { value?: Array<{ sequence?: number }> }).value ?? [];
        sequence = Math.max(0, ...rules.map((r) => r.sequence ?? 0)) + 1;
      }

      const rule: Record<string, unknown> = {
        displayName: opts.name,
        sequence,
        conditions,
        actions,
        isEnabled: opts.enabled !== false,
      };

      const response = await client.request(RULES, {
        method: 'POST',
        body: JSON.stringify(rule),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('update-rule <ruleId>')
    .description('Update an inbox rule')
    .option('--name <name>', 'Rule display name')
    .option('--enabled <bool>', 'Enable or disable the rule: true/false (also yes/no, on/off, 1/0)', (value: string) => {
      const v = value.toLowerCase();
      if (['true', 'yes', 'on', '1'].includes(v)) return true;
      if (['false', 'no', 'off', '0'].includes(v)) return false;
      throw new InvalidArgumentError('Expected true or false.');
    })
    .addOption(outputOption())
    .action(handleErrors(async (ruleId: string, opts) => {
      const client = getClient();
      const patch: Record<string, unknown> = {};
      if (opts.name) patch.displayName = opts.name;
      if (opts.enabled !== undefined) patch.isEnabled = opts.enabled;

      const response = await client.request(
        `${RULES}/${encodeURIComponent(ruleId)}`,
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
      info(`Marked ${succeeded} message(s) as read.`);
      failOnBatchErrors(failed, ids.length);
    }));

  mail
    .command('batch-move <messageIds>')
    .description('Move multiple messages to a folder (comma-separated IDs)')
    .requiredOption('--folder <name>', 'Destination folder name or ID')
    .action(handleErrors(async (messageIds: string, opts) => {
      const client = getClient();
      const folderId = await resolveFolder(client, opts.folder);
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
      info(`Moved ${succeeded} message(s).`);
      failOnBatchErrors(failed, ids.length);
    }));
}
