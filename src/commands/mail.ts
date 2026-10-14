import { Command } from 'commander';
import fs from 'fs';
import path from 'path';
import { GraphClient, type BatchResult } from '../graph-client.js';
import { formatOutput, type OutputFormat, printList } from '../formatter.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import {
  truncate, parseRecipients, formatSize, outputOption, confirm, info,
  odataQuote, kqlQuote, safeOutputPath, resolveTimeZone, parseTime, toGraphDateTime,
} from '../utils.js';
import { detectMimeType } from '../mime.js';
import { addContentOptions, contentOptions, messageTransform, transformData, type BodyMode } from '../content.js';

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

  // Search by displayName; odataQuote keeps the literal intact (including non-ASCII names)
  const trimmed = name.trim();
  if (!trimmed) return null;
  const response = await client.request('/me/mailFolders', {
    queryParams: {
      '$filter': `displayName eq '${odataQuote(trimmed)}'`,
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

// Inline (base64) attachments are limited to 3 MB; larger files need an upload session.
const INLINE_ATTACHMENT_LIMIT = 3 * 1024 * 1024;
const ATTACHMENT_CHUNK_SIZE = 3 * 1024 * 1024;

/**
 * Upload a large file attachment via POST …/attachments/createUploadSession, then PUT
 * the bytes to the pre-authorized uploadUrl in chunks (no bearer token is sent there).
 */
async function uploadLargeAttachment(
  client: GraphClient, messageId: string, filePath: string, name: string, contentType: string, size: number,
): Promise<unknown> {
  const session = await client.request(`/me/messages/${encodeURIComponent(messageId)}/attachments/createUploadSession`, {
    method: 'POST',
    body: JSON.stringify({ AttachmentItem: { attachmentType: 'file', name, size, contentType } }),
  });
  const sessionData = session.data as { uploadUrl?: string; dryRun?: boolean };
  if (sessionData.dryRun) {
    console.log(`[dry-run] Then PUT ${Math.ceil(size / ATTACHMENT_CHUNK_SIZE)} chunk(s) to the session uploadUrl`);
    return session.data;
  }
  if (!sessionData.uploadUrl) {
    throw new CliError('Upload session response did not include an uploadUrl.', { code: 'bad_response' });
  }

  const fd = fs.openSync(filePath, 'r');
  try {
    let offset = 0;
    while (offset < size) {
      const length = Math.min(ATTACHMENT_CHUNK_SIZE, size - offset);
      const buffer = Buffer.alloc(length);
      fs.readSync(fd, buffer, 0, length, offset);
      const res = await fetch(sessionData.uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/octet-stream',
          'Content-Length': String(length),
          'Content-Range': `bytes ${offset}-${offset + length - 1}/${size}`,
        },
        body: buffer,
      });
      if (!res.ok) {
        throw new CliError(`Attachment chunk upload failed at byte ${offset} (${res.status}): ${await res.text()}`,
          { code: 'upload_failed', status: res.status, exitCode: res.status === 429 ? ExitCode.THROTTLED : ExitCode.GENERIC });
      }
      offset += length;
      info(`Uploading... ${Math.round((offset / size) * 100)}%`);
      if (res.status === 201) break;
    }
  } finally {
    fs.closeSync(fd);
  }
  return { message: 'Attachment uploaded', name, size };
}

interface FileToAttach { filePath: string; name: string; contentType: string; size: number }

const TEXT_BODY_HEADER = { Prefer: 'outlook.body-content-type="text"' };
const textBodyHeaders = (body: BodyMode | undefined): Record<string, string> => (body === 'text' ? TEXT_BODY_HEADER : {});

// Default $select for `mail get --unique` (uniqueBody is only returned when selected).
const UNIQUE_SELECT = 'id,subject,from,sender,toRecipients,ccRecipients,bccRecipients,replyTo,receivedDateTime,sentDateTime,' +
  'bodyPreview,uniqueBody,hasAttachments,importance,categories,flag,isRead,isDraft,conversationId,internetMessageId,parentFolderId,webLink';

const collect = (value: string, previous: string[]): string[] => [...previous, value];

/** Stat every --attach path up front so a bad path fails (exit 2) before any request is sent. */
function prepareAttachments(paths: string[]): FileToAttach[] {
  return paths.map((filePath) => {
    let stat: fs.Stats;
    try {
      stat = fs.statSync(filePath);
    } catch {
      throw new CliError(`Attachment not found: ${filePath}`, { code: 'usage', exitCode: ExitCode.USAGE });
    }
    if (!stat.isFile()) {
      throw new CliError(`Attachment is not a file: ${filePath}`, { code: 'usage', exitCode: ExitCode.USAGE });
    }
    const name = path.basename(filePath);
    return { filePath, name, contentType: detectMimeType(name), size: stat.size };
  });
}

/** Any file over 3 MB, or more than 3 MB in total, can't go inline in one request. */
function needsUploadSession(files: FileToAttach[]): boolean {
  return files.reduce((total, f) => total + f.size, 0) > INLINE_ATTACHMENT_LIMIT;
}

function inlineAttachment(file: FileToAttach): Record<string, unknown> {
  return {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: file.name,
    contentType: file.contentType,
    contentBytes: fs.readFileSync(file.filePath).toString('base64'),
  };
}

/** Add files to an existing draft: inline POSTs for small ones, an upload session for large ones. */
async function addAttachments(client: GraphClient, messageId: string, files: FileToAttach[]): Promise<void> {
  for (const file of files) {
    if (file.size > INLINE_ATTACHMENT_LIMIT) {
      await uploadLargeAttachment(client, messageId, file.filePath, file.name, file.contentType, file.size);
    } else {
      await client.request(`/me/messages/${encodeURIComponent(messageId)}/attachments`, {
        method: 'POST',
        body: JSON.stringify(inlineAttachment(file)),
      });
    }
  }
}

/** POST a draft-creating endpoint and return the new draft's id (a placeholder under --dry-run). */
async function createDraft(client: GraphClient, endpoint: string, body: unknown): Promise<Record<string, unknown>> {
  const response = await client.request(endpoint, { method: 'POST', body: JSON.stringify(body) });
  const draft = response.data as Record<string, unknown>;
  return draft.dryRun ? { ...draft, id: 'DRAFT_ID' } : draft;
}

/** Create a draft via `endpoint`, attach the files, then send it. */
async function sendViaDraft(client: GraphClient, endpoint: string, body: unknown, files: FileToAttach[]): Promise<void> {
  const draft = await createDraft(client, endpoint, body);
  await addAttachments(client, draft.id as string, files);
  await client.request(`/me/messages/${encodeURIComponent(draft.id as string)}/send`, { method: 'POST' });
}

export function registerMailCommands(program: Command, getClient: () => GraphClient): void {
  const mail = program.command('mail').description('Outlook mail operations');

  addContentOptions(mail
    .command('list')
    .description('List inbox messages')
    .option('--top <n>', 'Number of messages to return', '10')
    .option('--filter <filter>', 'OData filter expression')
    .option('--select <fields>', 'Comma-separated fields to return')
    .option('--search <query>', 'Search query (cannot be combined with --filter; Graph returns at most 250 results)')
    .option('--folder <name>', 'Folder name or ID (e.g. "Drafts", "Sent Items", "Archive", or a folder ID)')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms'))
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.select) queryParams['$select'] = opts.select;
      if (opts.search) {
        if (opts.filter) {
          throw new CliError('--search cannot be combined with --filter: Outlook $search does not support $filter or $orderby.',
            { code: 'usage', exitCode: ExitCode.USAGE, fix: 'drop --filter, or filter the --search results client-side (e.g. with jq)' });
        }
        queryParams['$search'] = kqlQuote(opts.search);
        if (opts.all || parseInt(opts.top) > 250) {
          info('Warning: mail search returns at most 250 results; --all cannot page beyond that.');
        }
      }

      let endpoint = '/me/messages';
      if (opts.folder) {
        // Try as well-known name first, then resolve
        const folderId = await resolveFolderName(client, opts.folder) || opts.folder;
        endpoint = `/me/mailFolders/${encodeURIComponent(folderId)}/messages`;
      }

      await printList(client, endpoint, { queryParams, headers: textBodyHeaders(opts.body) }, opts,
        { transform: messageTransform(client, 'mail', contentOptions(opts)) });
    }));

  addContentOptions(mail
    .command('get <messageId>')
    .description('Get a specific message (marked untrusted; see --body, --unique, --wrap-untrusted)')
    .option('--select <fields>', 'Comma-separated fields to return')
    .option('--unique', 'Return uniqueBody (the new part of a thread reply, without quoted history) instead of body'))
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.unique) {
        queryParams['$select'] = opts.select
          ? [...new Set(opts.select.split(',').map((f: string) => f.trim()).filter((f: string) => f && f !== 'body').concat('uniqueBody'))].join(',')
          : UNIQUE_SELECT;
      } else if (opts.select) {
        queryParams['$select'] = opts.select;
      }

      const response = await client.request(`/me/messages/${encodeURIComponent(messageId)}`, {
        queryParams, headers: textBodyHeaders(opts.body),
      });
      const data = await transformData(response.data, messageTransform(client, 'mail', contentOptions(opts)));
      console.log(formatOutput(data, opts.output as OutputFormat));
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
        console.log(formatOutput(await transformData(response.data, messageTransform(client, 'mail', {})), 'json'));
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

  addContentOptions(mail
    .command('triage')
    .description('Show inbox summary — unread messages with key metadata')
    .option('--top <n>', 'Number of messages', '20')
    .option('--filter <filter>', 'OData filter (default: unread)', 'isRead eq false')
    .option('--folder <folderId>', 'Mail folder ID')
    .option('--include-read', 'Include read messages too'), false)
    .addOption(outputOption('table'))
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {
        '$top': opts.top,
        '$select': 'id,subject,from,receivedDateTime,hasAttachments,importance,categories,flag,isRead,bodyPreview',
        '$orderby': 'receivedDateTime desc',
      };
      if (!opts.includeRead) {
        // Graph requires the $orderby property to appear first in $filter (else 400 InefficientFilter).
        queryParams['$filter'] = `receivedDateTime ge 1900-01-01T00:00:00Z and (${opts.filter})`;
      }

      const endpoint = opts.folder
        ? `/me/mailFolders/${encodeURIComponent(opts.folder)}/messages`
        : '/me/messages';

      const response = await client.request(endpoint, { queryParams });
      const data = await transformData(response.data, messageTransform(client, 'mail', contentOptions(opts))) as Record<string, unknown>;
      const messages = (data.value as Array<Record<string, unknown>>) || [];

      if (opts.output === 'json') {
        console.log(formatOutput(data, 'json'));
        return;
      }

      if (messages.length === 0) {
        info('No messages found.');
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
    .option('--attach <path>', 'Attach a file (repeatable; over 3 MB total goes via a draft + upload session)', collect, [])
    .action(handleErrors(async (opts) => {
      const files = prepareAttachments(opts.attach);
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

      if (needsUploadSession(files)) {
        await sendViaDraft(client, '/me/messages', message, files);
      } else {
        if (files.length > 0) message.attachments = files.map(inlineAttachment);
        await client.request('/me/sendMail', {
          method: 'POST',
          body: JSON.stringify({ message, saveToSentItems: true }),
        });
      }

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
      await printList(client, endpoint, { queryParams }, opts);
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
    .option('--attach <path>', 'Attach a file (repeatable; over 3 MB total goes via a draft + upload session)', collect, [])
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const files = prepareAttachments(opts.attach);
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

      const upload = needsUploadSession(files);
      if (!upload && files.length > 0) message.attachments = files.map(inlineAttachment);
      const created = await createDraft(client, '/me/messages', message);
      if (upload) await addAttachments(client, created.id as string, files);

      if (opts.output !== 'json') {
        info(`Draft created. ID: ${created.id}`);
      } else {
        console.log(formatOutput(created, 'json'));
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
      console.log(formatOutput(await transformData(response.data, messageTransform(client, 'mail', {})), opts.output as OutputFormat));
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
      // Editing subject/body/recipients of a sent/received message needs Mail-Advanced.ReadWrite
      // (admin consent) from 2026-12-31; refuse up front instead of surfacing an unclear 403.
      const current = await client.request(`/me/messages/${encodeURIComponent(messageId)}`, {
        queryParams: { '$select': 'isDraft' },
      });
      if ((current.data as Record<string, unknown>).isDraft === false) {
        throw new CliError(`Message ${messageId.substring(0, 16)}... is not a draft. Changing the subject, body or recipients ` +
          'of a non-draft message requires the Mail-Advanced.ReadWrite permission (admin consent) from 2026-12-31.',
          { code: 'not_a_draft', fix: 'create a new draft with `ms365 mail draft`, or use `mail forward`/`mail reply`' });
      }
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
    .option('--name <name>', 'Attachment file name (default with --file: the file\'s base name)')
    .requiredOption('--content-type <type>', 'MIME content type')
    .option('--content <base64>', 'Base64-encoded content (small attachments only)')
    .option('--file <path>', 'Read the attachment from a file (over 3 MB uses an upload session)')
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts) => {
      if (!!opts.content === !!opts.file) {
        throw new CliError('Pass exactly one of --content or --file.', { code: 'usage', exitCode: ExitCode.USAGE });
      }
      const name: string | undefined = opts.name || (opts.file && path.basename(opts.file));
      if (!name) {
        throw new CliError('--name is required with --content.', { code: 'usage', exitCode: ExitCode.USAGE });
      }
      const client = getClient();
      let contentBytes: string = opts.content;
      if (opts.file) {
        const size = fs.statSync(opts.file).size;
        if (size > INLINE_ATTACHMENT_LIMIT) {
          const result = await uploadLargeAttachment(client, messageId, opts.file, name, opts.contentType, size);
          console.log(formatOutput(result, opts.output as OutputFormat));
          return;
        }
        contentBytes = fs.readFileSync(opts.file).toString('base64');
      }
      const response = await client.request(
        `/me/messages/${encodeURIComponent(messageId)}/attachments`,
        {
          method: 'POST',
          body: JSON.stringify({
            '@odata.type': '#microsoft.graph.fileAttachment',
            name,
            contentType: opts.contentType,
            contentBytes,
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
    .option('--out <path>', 'Output file path (default: attachment name in the current directory)')
    .option('--force', 'Overwrite an existing file')
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
      const outPath = safeOutputPath(name, { out: opts.out, force: opts.force, fallback: 'attachment' });
      const bytes = Buffer.from(contentBytes, 'base64');
      fs.writeFileSync(outPath, bytes);
      info(`Saved to ${outPath} (${formatSize(bytes.length)})`);
    }));

  mail
    .command('reply <messageId>')
    .description('Reply to a message')
    .requiredOption('--body <body>', 'Reply body')
    .option('--html', 'Send body as HTML')
    .option('--reply-all', 'Reply to all recipients')
    .option('--attach <path>', 'Attach a file (repeatable; over 3 MB total goes via a draft + upload session)', collect, [])
    .action(handleErrors(async (messageId: string, opts) => {
      const files = prepareAttachments(opts.attach);
      const client = getClient();
      const base = `/me/messages/${encodeURIComponent(messageId)}`;
      if (needsUploadSession(files)) {
        await sendViaDraft(client, `${base}/${opts.replyAll ? 'createReplyAll' : 'createReply'}`, { comment: opts.body }, files);
      } else {
        const body: Record<string, unknown> = { comment: opts.body };
        if (files.length > 0) body.message = { attachments: files.map(inlineAttachment) };
        await client.request(`${base}/${opts.replyAll ? 'replyAll' : 'reply'}`, { method: 'POST', body: JSON.stringify(body) });
      }
      info('Reply sent.');
    }));

  // ── Forward ──────────────────────────────────────────────────
  mail
    .command('forward <messageId>')
    .description('Forward a message')
    .requiredOption('--to <addresses>', 'Comma-separated recipient emails')
    .option('--comment <text>', 'Comment to include with forward')
    .option('--attach <path>', 'Attach a file (repeatable; over 3 MB total goes via a draft + upload session)', collect, [])
    .action(handleErrors(async (messageId: string, opts) => {
      const files = prepareAttachments(opts.attach);
      const client = getClient();
      const toRecipients = parseRecipients(opts.to);
      const base = `/me/messages/${encodeURIComponent(messageId)}`;
      const body: Record<string, unknown> = { comment: opts.comment || '', toRecipients };
      if (needsUploadSession(files)) {
        await sendViaDraft(client, `${base}/createForward`, body, files);
      } else {
        if (files.length > 0) body.message = { attachments: files.map(inlineAttachment) };
        await client.request(`${base}/forward`, { method: 'POST', body: JSON.stringify(body) });
      }
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
    .option('--due <when>', 'Due date/time (YYYY-MM-DD, ISO 8601, "tomorrow", "next fri 17:00", "+3d")')
    .option('--timezone <tz>', 'IANA time zone for --due (default: MS365_CLI_TIMEZONE or local)')
    .action(handleErrors(async (messageId: string, opts) => {
      const flag: Record<string, unknown> = { flagStatus: 'flagged' };
      if (opts.due) {
        // Graph rejects a dueDateTime without a startDateTime; start now (or at the due time if it is past).
        const tz = resolveTimeZone(opts.timezone);
        const due = parseTime(opts.due, { tz });
        const start = new Date(Math.min(Date.now(), due.date.getTime()));
        flag.startDateTime = { dateTime: toGraphDateTime(start, tz), timeZone: tz };
        flag.dueDateTime = { dateTime: due.dateTime, timeZone: tz };
      }
      const client = getClient();
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

      await printList(client, endpoint, { queryParams }, opts, { transform: messageTransform(client, 'mail', {}) });
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
      console.log(formatOutput(await transformData(response.data, messageTransform(client, 'mail', {})), opts.output as OutputFormat));
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
    .option('--all-messages', 'Apply the rule to every incoming message (required when no condition is given)')
    .option('--disabled', 'Create the rule disabled')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const conditions: Record<string, unknown> = {};
      if (opts.from) {
        conditions.senderContains = opts.from.split(',').map((s: string) => s.trim());
      }
      if (opts.subjectContains) {
        conditions.subjectContains = opts.subjectContains.split(',').map((s: string) => s.trim());
      }

      if (Object.keys(conditions).length === 0 && !opts.allMessages) {
        throw new CliError('Refusing to create a rule with no conditions: it would act on every incoming message.',
          { code: 'usage', exitCode: ExitCode.USAGE, fix: 'add --from or --subject-contains, or pass --all-messages to confirm' });
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
        isEnabled: !opts.disabled,
      };

      const client = getClient();
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
    .description('Mark multiple messages as read (comma-separated IDs; sent via JSON batching)')
    .action(handleErrors(async (messageIds: string) => {
      const ids = messageIds.split(',').map(id => id.trim()).filter(Boolean);
      const results = await getClient().batch(ids.map(id => ({
        method: 'PATCH', url: `/me/messages/${encodeURIComponent(id)}`, body: { isRead: true },
      })));
      reportBatch(ids, results, 'Marked', 'as read');
    }));

  mail
    .command('batch-move <messageIds>')
    .description('Move multiple messages to a folder (comma-separated IDs; sent via JSON batching)')
    .requiredOption('--folder <folderId>', 'Destination folder name or ID')
    .action(handleErrors(async (messageIds: string, opts) => {
      const client = getClient();
      const folderId = await resolveFolderName(client, opts.folder) || opts.folder;
      const ids = messageIds.split(',').map(id => id.trim()).filter(Boolean);
      const results = await client.batch(ids.map(id => ({
        method: 'POST', url: `/me/messages/${encodeURIComponent(id)}/move`, body: { destinationId: folderId },
      })));
      reportBatch(ids, results, 'Moved', '');
    }));

}

/** Summarize batch results on stderr; fail (exit 1) when any item failed. */
function reportBatch(ids: string[], results: BatchResult[], verb: string, suffix: string): void {
  const failed = ids
    .map((id, i) => ({ id, r: results[i] }))
    .filter(({ r }) => !(r.status === 0 && (r.body as { dryRun?: boolean } | null)?.dryRun) && (r.status < 200 || r.status >= 300));
  info(`${verb} ${ids.length - failed.length} message(s)${suffix ? ' ' + suffix : ''}.`);
  if (failed.length > 0) {
    const detail = failed.map(({ id, r }) => {
      const msg = (r.body as { error?: { message?: string } } | null)?.error?.message ?? `status ${r.status}`;
      return `${id.substring(0, 12)}...: ${msg}`;
    }).join('; ');
    throw new CliError(`Failed ${failed.length} of ${ids.length}: ${detail}`, { code: 'partial_failure' });
  }
}
