import { Command, Option } from 'commander';
import fs from 'fs';
import path from 'path';
import { GraphClient, type BatchResult, type GraphResponse } from '../graph-client.js';
import { formatOutput, type OutputFormat, printList } from '../formatter.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import {
  truncate, parseRecipients, formatSize, addListOptions, outputOption, confirm, info, success,
  kqlQuote, odataQuote, safeOutputPath, resolveTimeZone, parseTime, toGraphDateTime, fromGet, addBodyOptions, bodyText, mailboxOption, userPath, intParser,
} from '../utils.js';
import { detectMimeType } from '../mime.js';
import {
  fingerprint, idempotentSend, withIdempotency, IDEMPOTENCY_HELP, listEntries, settleKey, type Entry, type IdempotencyResult,
} from '../idempotency.js';
import { attachmentDigests, ifMatchFor, planning, readVersion, setPlanPayload, setPlanTarget, usage } from '../plan.js';
import { resolve } from '../resolve.js';
import { addContentOptions, contentOptions, messageTransform, transformData, type BodyMode } from '../content.js';

// Inline (base64) attachments are limited to 3 MB; larger files need an upload session.
const INLINE_ATTACHMENT_LIMIT = 3 * 1024 * 1024;
// Outlook caps upload-session chunks at 4 MB; 3 MiB is a multiple of 320 KiB.
const ATTACHMENT_CHUNK_SIZE = 3 * 1024 * 1024;

/** Upload a large file attachment through an Outlook upload session (POST …/attachments/createUploadSession). */
function uploadLargeAttachment(
  client: GraphClient, messageId: string, filePath: string, name: string, contentType: string, size: number,
): Promise<GraphResponse> {
  const attachments = `/me/messages/${encodeURIComponent(messageId)}/attachments`;
  return client.uploadSession(`${attachments}/createUploadSession`,
    { AttachmentItem: { attachmentType: 'file', name, size, contentType } }, filePath, size,
    { chunkSize: ATTACHMENT_CHUNK_SIZE, auditAs: attachments });
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
  return client.isPreview(response) && client.isDryRun() ? { ...draft, id: 'DRAFT_ID' } : draft;
}

/**
 * GET a message, refuse a non-draft, and return its version for If-Match. Editing the subject, body or recipients of a
 * sent/received message needs Mail-Advanced.ReadWrite (admin consent) from 2026-12-31; refusing up front beats an
 * unclear 403.
 */
async function draftVersion(client: GraphClient, messageId: string): Promise<string | undefined> {
  const { item, version } = await readVersion(client, `/me/messages/${encodeURIComponent(messageId)}`, 'isDraft');
  if (item && typeof item === 'object' && Reflect.get(item, 'isDraft') === false) {
    throw new CliError(`Message ${messageId.substring(0, 16)}... is not a draft. Changing the subject, body or recipients ` +
      'of a non-draft message requires the Mail-Advanced.ReadWrite permission (admin consent) from 2026-12-31.',
      { code: 'not_a_draft', fix: 'create a new draft with `ms365 mail draft`, or use `mail forward`/`mail reply`' });
  }
  return version;
}

/**
 * Put `html` at the top of a draft's body, keeping the quoted thread that createReply/createForward wrote
 * below it: GET the body, insert right after `<body…>`, PATCH it back (F18). Under --dry-run the GET has no
 * body, so the preview shows `<from GET>` where the quoted thread goes.
 */
async function prependHtml(client: GraphClient, draftId: string, html: string): Promise<void> {
  if (planning()) {
    throw usage('A reply or forward body (plain text or --html) goes above the new draft\'s quoted thread, which does not exist yet, so it cannot be planned. Plan it with an empty body (forward: no --comment).');
  }
  const endpoint = `/me/messages/${encodeURIComponent(draftId)}`;
  const current = await client.request(endpoint, {
    queryParams: { '$select': 'body' }, headers: { Prefer: 'outlook.body-content-type="html"' },
  });
  const existing = client.isPreview(current)
    ? fromGet()
    : String((current.data as { body?: { content?: string } }).body?.content ?? '');
  const open = /<body[^>]*>/i.exec(existing);
  const content = open ? existing.slice(0, open.index + open[0].length) + html + existing.slice(open.index + open[0].length) : html + existing;
  await client.request(endpoint, { method: 'PATCH', body: JSON.stringify({ body: { contentType: 'HTML', content } }) });
}

const HTML_ENTITIES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };

/**
 * A plain-text reply/forward body as HTML: escaped, line breaks as `<br>`. Graph's `comment` drops the line breaks
 * (the text renders as one line), so plain text goes through the same draft path as --html (E36).
 */
function plainTextHtml(text: string): string {
  return `<div>${text.replace(/[&<>"']/g, c => HTML_ENTITIES[c]).replace(/\r\n|\r|\n/g, '<br>')}</div>`;
}

/**
 * The HTML a reply/forward puts above the quoted thread: --html as given, plain text escaped; none for an empty body,
 * which goes as `comment` without a draft, so only a non-empty body needs Mail.ReadWrite (catalog, E37).
 */
function replyHtml(text: string, html: boolean): string | undefined {
  if (!text) return undefined;
  return html ? text : plainTextHtml(text);
}

/**
 * The sequence that puts a new inbox rule after every existing one: 1 + the highest `sequence` over all pages, 1 when
 * there are none. Graph rejects a rule without a sequence (0) as `InvalidValue` (E38). Under --dry-run the listing is
 * not read, so the preview shows `<from GET>`.
 */
async function nextRuleSequence(client: GraphClient): Promise<number | string> {
  const endpoint = '/me/mailFolders/inbox/messageRules';
  let highest = 0;
  let next: string | undefined = endpoint;
  while (next) {
    const response: GraphResponse = await client.request(next, next === endpoint ? { queryParams: { '$select': 'sequence' } } : {});
    if (client.isPreview(response)) return fromGet();
    const data: unknown = response.data;
    const value: unknown = data && typeof data === 'object' ? Reflect.get(data, 'value') : undefined;
    for (const item of Array.isArray(value) ? value : []) {
      const sequence: unknown = item && typeof item === 'object' ? Reflect.get(item, 'sequence') : undefined;
      if (typeof sequence === 'number' && sequence > highest) highest = sequence;
    }
    next = response.nextLink;
  }
  return highest + 1;
}

/**
 * Under --plan, a reply or forward binds the message it answers (id and version, for the reviewer) and its payload.
 * Apply sends no If-Match here: the original is not modified, and its changeKey moves whenever it is read or flagged
 * (E19).
 */
async function planTargetMessage(client: GraphClient, messageId: string, request: { attachments: Array<{ name: string; size: number; sha256: string }> }): Promise<void> {
  if (!planning()) return;
  const { version } = await readVersion(client, `/me/messages/${encodeURIComponent(messageId)}`, 'id');
  setPlanTarget({ id: messageId, version: version ?? '' });
  setPlanPayload({ ...request, attachments: request.attachments.map(({ name, size, sha256 }) => ({ name, size, sha256 })) });
}

/** The handles of a draft-path send (AH-6): the draft's immutable id and its internetMessageId (null under preview). */
export interface SendCorrelation { draftId: string; internetMessageId: string | null }

/** The internetMessageId a ledger entry's send recorded (its result's or its unknown outcome's correlation). */
function sentMessageId(entry: Entry): string | undefined {
  const correlation: unknown = entry.result?.correlation;
  const id: unknown = correlation && typeof correlation === 'object' ? Reflect.get(correlation, 'internetMessageId') : undefined;
  return typeof id === 'string' ? id : undefined;
}

/**
 * Create a draft via `endpoint`, put `html` on top of its body (if given) and attach the files; return the created
 * draft. If a step after the create fails, the draft is deleted (best effort) so retries don't pile up copies in
 * Drafts (R-8).
 */
async function prepareDraft(
  client: GraphClient, endpoint: string, body: unknown, files: FileToAttach[], html?: string,
): Promise<Record<string, unknown>> {
  const draft = await createDraft(client, endpoint, body);
  const draftId = String(draft.id);
  try {
    if (html !== undefined) await prependHtml(client, draftId, html);
    await addAttachments(client, draftId, files);
  } catch (err) {
    const deleted = await client.request(`/me/messages/${encodeURIComponent(draftId)}`, { method: 'DELETE' }).then(() => true, () => false);
    if (err instanceof Error) {
      err.message += deleted
        ? ' The unsent draft was deleted.'
        : ` The unsent draft ${draftId} could not be deleted; remove it from Drafts.`;
    }
    // A deleted draft can no longer be sent, so the request cannot have taken effect.
    if (deleted && err instanceof CliError) err.maybeProcessed = undefined;
    throw err;
  }
  return draft;
}

/** What `mail reply --draft`/`mail forward --draft` print and their --idempotency-key stores: the draft's handles, no message content. */
function draftHandles(draft: Record<string, unknown>): IdempotencyResult {
  const handle = (k: string) => (typeof draft[k] === 'string' ? draft[k] : null);
  return {
    id: handle('id'), conversationId: handle('conversationId'), webLink: handle('webLink'),
    internetMessageId: handle('internetMessageId'), createdDateTime: handle('createdDateTime'),
  };
}

/**
 * Create a draft under --idempotency-key (D23) and return its handles. A repeat with the same key and request returns
 * the first draft's handles and calls Graph for nothing, so a retry never leaves a second draft. Dry-run and --plan
 * never touch the ledger.
 */
async function idempotentDraft(
  client: GraphClient, key: string | undefined, command: string, request: unknown, create: () => Promise<Record<string, unknown>>,
): Promise<IdempotencyResult> {
  const run = async () => draftHandles(await create());
  if (!key || client.isDryRun() || planning()) return run();
  const outcome = await withIdempotency(key, command, fingerprint(request), run);
  if (outcome.status === 'replay') info(`Draft already created at ${outcome.ts}; not created again.`);
  return outcome.result;
}

/**
 * Create and fill a draft via `endpoint` (prepareDraft), then send it, and return its correlation. Nothing is deleted
 * once /send was dispatched: a lost /send response is exit 1 `outcome_unknown` with `maybeProcessed` and the
 * correlation (AH-6(3)).
 */
async function sendViaDraft(
  client: GraphClient, endpoint: string, body: unknown, files: FileToAttach[], html?: string,
): Promise<SendCorrelation> {
  const draft = await prepareDraft(client, endpoint, body, files, html);
  const draftId = String(draft.id);
  const correlation: SendCorrelation = {
    draftId, internetMessageId: typeof draft.internetMessageId === 'string' ? draft.internetMessageId : null,
  };
  const draftPath = `/me/messages/${encodeURIComponent(draftId)}`;
  try {
    await client.request(`${draftPath}/send`, { method: 'POST' });
  } catch (err) {
    if (!(err instanceof CliError) || !err.maybeProcessed) throw err;
    const ref = correlation.internetMessageId ? `--internet-message-id ${JSON.stringify(correlation.internetMessageId)}` : '--idempotency-key <key>';
    throw new CliError(`The send of draft ${draftId} may or may not have been accepted (${err.message}).`, {
      code: 'outcome_unknown', exitCode: ExitCode.GENERIC, maybeProcessed: true, correlation: { ...correlation },
      status: err.status, requestId: err.requestId,
      fix: `do not resend; run \`ms365 mail reconcile ${ref}\` to see whether it reached Sent Items`,
    });
  }
  return correlation;
}

export function registerMailCommands(program: Command, getClient: () => GraphClient): void {
  const mail = program.command('mail').description('Outlook mail operations');

  addListOptions(addContentOptions(mail
    .command('list')
    .description('List inbox messages')
    .option('--filter <filter>', 'OData filter expression')
    .option('--select <fields>', 'Comma-separated fields to return')
    .option('--search <query>', 'Search query (cannot be combined with --filter; Graph returns at most 250 results)')
    .option('--folder <name>', 'Folder name, path or ID (e.g. "Drafts", "Sent Items", "Inbox/Projects", or a folder ID)')), { description: 'Number of messages to return', default: '10' })
    .addOption(mailboxOption('Mail.Read.Shared'))
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

      const owner = userPath(opts.mailbox);
      const endpoint = opts.folder
        ? `${owner}/mailFolders/${encodeURIComponent(await resolve(client, 'mail-folder', opts.folder, opts.mailbox))}/messages`
        : `${owner}/messages`;

      // AH-3: a $search that returns its 250-item maximum says so in band (incompleteReason "search_cap").
      await printList(client, endpoint, { queryParams, headers: textBodyHeaders(opts.body) }, opts,
        { transform: messageTransform(client, 'mail', contentOptions(opts)), searchCap: opts.search ? 250 : undefined });
    }));

  addContentOptions(mail
    .command('get <messageId>')
    .description('Get a specific message (marked untrusted; see --body, --unique, --wrap-untrusted)')
    .option('--select <fields>', 'Comma-separated fields to return')
    .option('--unique', 'Return uniqueBody (the new part of a thread reply, without quoted history) instead of body'))
    .addOption(mailboxOption('Mail.Read.Shared'))
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

      const response = await client.request(`${userPath(opts.mailbox)}/messages/${encodeURIComponent(messageId)}`, {
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
    .option('--folder <name>', 'Folder name, path (Inbox/Projects) or ID')
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
        ? `/me/mailFolders/${encodeURIComponent(await resolve(client, 'mail-folder', opts.folder))}/messages`
        : '/me/messages';

      const response = await client.request(endpoint, { queryParams });
      if (client.isPreview(response)) {
        // The trace is the whole preview: no report built from the placeholder (DR-12).
        if (opts.output === 'json') console.log(formatOutput({ dryRun: true }, 'json'));
        return;
      }
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

  addBodyOptions(mail
    .command('send')
    .description('Send an email (via a draft, so the result carries its id and internetMessageId; accepted, not delivered)')
    .requiredOption('--to <addresses>', 'Comma-separated recipient email addresses')
    .option('--cc <addresses>', 'Comma-separated CC addresses')
    .option('--bcc <addresses>', 'Comma-separated BCC addresses')
    .requiredOption('--subject <subject>', 'Email subject'), 'Email body', true)
    .option('--html', 'Send body as HTML')
    .option('--attach <path>', 'Attach a file (repeatable; over 3 MB goes via an upload session)', collect, [])
    .option('--idempotency-key <key>', IDEMPOTENCY_HELP)
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const content = bodyText(opts) ?? '';
      const files = prepareAttachments(opts.attach);
      const client = getClient();
      const toRecipients = parseRecipients(opts.to);

      const message: Record<string, unknown> = {
        subject: opts.subject,
        body: {
          contentType: opts.html ? 'HTML' : 'Text',
          content,
        },
        toRecipients,
      };

      if (opts.cc) {
        message.ccRecipients = parseRecipients(opts.cc);
      }

      if (opts.bcc) {
        message.bccRecipients = parseRecipients(opts.bcc);
      }

      const attachments = attachmentDigests(files);
      setPlanPayload({ ...message, attachments: attachments.map(({ name, size, sha256 }) => ({ name, size, sha256 })) });
      const key: string | undefined = opts.idempotencyKey;
      // AH-6(1): always the draft path, so the send has a handle (draft id, internetMessageId) before the 202.
      const accepted = await idempotentSend(client, key, 'mail send', { message, attachments }, async () => {
        const upload = needsUploadSession(files);
        const draft = !upload && files.length > 0 ? { ...message, attachments: files.map(inlineAttachment) } : message;
        let correlation: SendCorrelation;
        try {
          correlation = await sendViaDraft(client, '/me/messages', draft, upload ? files : []);
        } catch (err) {
          if (err instanceof CliError && err.correlation) err.correlation.idempotencyKey = key ?? null;
          throw err;
        }
        return {
          outcome: 'ACCEPTED', acceptedAt: new Date().toISOString(),
          correlation: { ...correlation, idempotencyKey: key ?? null }, delivery: 'unknown',
        };
      });
      if (!accepted || planning() || client.isDryRun()) return;
      if (opts.output === 'json') {
        console.log(formatOutput(accepted, 'json'));
        return;
      }
      const correlation: unknown = accepted.correlation;
      const imid = correlation && typeof correlation === 'object' ? Reflect.get(correlation, 'internetMessageId') : null;
      success(`Accepted by Exchange for delivery${typeof imid === 'string' ? ` (internetMessageId ${imid})` : ''}; delivery is not confirmed.`);
    }));

  mail
    .command('delete <messageId>')
    .description('Delete a message')
    .action(handleErrors(async (messageId: string) => {
      if (!await confirm(`Delete message ${messageId.substring(0, 16)}...?`)) return;
      const client = getClient();
      await client.request(`/me/messages/${encodeURIComponent(messageId)}`, { method: 'DELETE' });
      success('Message deleted.');
    }));

  mail
    .command('move <messageId>')
    .description('Move a message to a folder')
    .requiredOption('--folder <name>', 'Destination folder name, path (Inbox/Projects) or ID')
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      const destinationId = await resolve(client, 'mail-folder', opts.folder);
      await client.request(`/me/messages/${encodeURIComponent(messageId)}/move`, {
        method: 'POST',
        body: JSON.stringify({ destinationId }),
      });
      success('Message moved.');
    }));

  addListOptions(mail
    .command('folders')
    .description('List mail folders')
    .option('--parent <name>', 'List child folders of this folder (name, path (inbox/projects) or id)'))
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const endpoint = opts.parent
        ? `/me/mailFolders/${encodeURIComponent(await resolve(client, 'mail-folder', opts.parent))}/childFolders`
        : '/me/mailFolders';
      const queryParams: Record<string, string> = {
        '$select': 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount',
        '$top': '100',
      };
      await printList(client, endpoint, { queryParams }, opts);
    }));

  mail
    .command('folder <name>')
    .description('Find a folder by display name or path (e.g. "Drafts", "Sent Items", "Inbox/Projects")')
    .addOption(outputOption())
    .action(handleErrors(async (folderName: string, opts) => {
      const client = getClient();
      const folderId = await resolve(client, 'mail-folder', folderName);
      const response = await client.request(`/me/mailFolders/${encodeURIComponent(folderId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('create-folder')
    .description('Create a mail folder')
    .requiredOption('--name <name>', 'Folder display name')
    .option('--parent <name>', 'Parent folder name, path or ID (default: root)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const endpoint = opts.parent
        ? `/me/mailFolders/${encodeURIComponent(await resolve(client, 'mail-folder', opts.parent))}/childFolders`
        : '/me/mailFolders';
      const response = await client.request(endpoint, {
        method: 'POST',
        body: JSON.stringify({ displayName: opts.name }),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // ── Drafts ────────────────────────────────────────────────────
  addBodyOptions(mail
    .command('draft')
    .description('Create a draft email')
    .requiredOption('--to <addresses>', 'Comma-separated recipient email addresses')
    .requiredOption('--subject <subject>', 'Email subject'), 'Email body', true)
    .option('--cc <addresses>', 'Comma-separated CC addresses')
    .option('--bcc <addresses>', 'Comma-separated BCC addresses')
    .option('--html', 'Send body as HTML')
    .option('--attach <path>', 'Attach a file (repeatable; over 3 MB total goes via a draft + upload session)', collect, [])
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const content = bodyText(opts) ?? '';
      const files = prepareAttachments(opts.attach);
      const client = getClient();
      const toRecipients = parseRecipients(opts.to);

      const message: Record<string, unknown> = {
        subject: opts.subject,
        body: {
          contentType: opts.html ? 'HTML' : 'Text',
          content,
        },
        toRecipients,
      };

      if (opts.cc) {
        message.ccRecipients = parseRecipients(opts.cc);
      }
      if (opts.bcc) {
        message.bccRecipients = parseRecipients(opts.bcc);
      }
      setPlanPayload({ ...message, attachments: attachmentDigests(files).map(({ name, size, sha256 }) => ({ name, size, sha256 })) });

      const upload = needsUploadSession(files);
      if (!upload && files.length > 0) message.attachments = files.map(inlineAttachment);
      const created = await createDraft(client, '/me/messages', message);
      if (upload) await addAttachments(client, created.id as string, files);
      if (planning()) return;

      if (opts.output !== 'json') {
        success(`Draft created. ID: ${created.id}`);
      } else {
        console.log(formatOutput(created, 'json'));
      }
    }));

  addListOptions(mail
    .command('drafts')
    .description('List draft emails'), { description: 'Number of drafts', default: '10' })
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      await printList(client, '/me/mailFolders/drafts/messages', {
        queryParams: {
          '$top': opts.top,
          '$select': 'id,subject,toRecipients,createdDateTime,lastModifiedDateTime,bodyPreview,hasAttachments',
          '$orderby': 'lastModifiedDateTime desc',
        },
      }, opts, { transform: messageTransform(client, 'mail', {}) });
    }));

  addBodyOptions(mail
    .command('update-draft <messageId>')
    .description('Update an existing draft')
    .option('--to <addresses>', 'Comma-separated recipient emails')
    .option('--cc <addresses>', 'Comma-separated CC addresses')
    .option('--bcc <addresses>', 'Comma-separated BCC addresses')
    .option('--subject <subject>', 'Email subject'), 'Email body')
    .option('--html', 'Body is HTML')
    .option('--if-match <etag>', 'Only update if the draft is unchanged: an @odata.etag, or "auto" to read it first (412 → exit 4 version_changed)')
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts) => {
      const content = bodyText(opts);
      const client = getClient();
      const version = await draftVersion(client, messageId);
      const patch: Record<string, unknown> = {};

      if (opts.subject) patch.subject = opts.subject;
      if (content !== undefined) {
        patch.body = { contentType: opts.html ? 'HTML' : 'Text', content };
      }
      if (opts.to) {
        patch.toRecipients = parseRecipients(opts.to);
      }
      if (opts.cc) {
        patch.ccRecipients = parseRecipients(opts.cc);
      }
      if (opts.bcc) {
        patch.bccRecipients = parseRecipients(opts.bcc);
      }
      setPlanPayload(patch);
      const ifMatch = await ifMatchFor(opts.ifMatch, async () => version);
      if (planning() && ifMatch) setPlanTarget({ id: messageId, version: ifMatch });

      const response = await client.request(`/me/messages/${encodeURIComponent(messageId)}`, {
        method: 'PATCH',
        body: JSON.stringify(patch),
        ...(ifMatch && { headers: { 'If-Match': ifMatch } }),
      });
      if (planning()) return;

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  mail
    .command('send-draft <messageId>')
    .description('Send an existing draft')
    .option('--idempotency-key <key>', IDEMPOTENCY_HELP)
    .action(handleErrors(async (messageId: string, opts) => {
      const client = getClient();
      // Under --plan the draft's version is bound: apply sends If-Match, so an edited draft is not sent (AH-5).
      const ifMatch = await ifMatchFor(undefined, () => draftVersion(client, messageId));
      if (ifMatch) {
        setPlanTarget({ id: messageId, version: ifMatch });
        setPlanPayload({ messageId });
      }
      const sent = await idempotentSend(client, opts.idempotencyKey, 'mail send-draft', { messageId }, async () => {
        await client.request(`/me/messages/${encodeURIComponent(messageId)}/send`, {
          method: 'POST',
          ...(ifMatch && { headers: { 'If-Match': ifMatch } }),
        });
      });
      if (sent && !planning()) success('Draft sent.');
    }));

  // ── Reconcile a send whose outcome is unknown (AH-6(4)) ─────
  mail
    .command('reconcile')
    .description('Find out what became of a send: SENT (in Sent Items), STILL_DRAFT or NOT_FOUND_YET (never "not sent")')
    .addOption(new Option('--idempotency-key <key>', 'The send\'s --idempotency-key (its ledger entry holds the internetMessageId)').conflicts('internetMessageId'))
    .option('--internet-message-id <id>', 'The internetMessageId from the send result or its outcome_unknown error')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const key: string | undefined = opts.idempotencyKey;
      const entries = listEntries();
      const byKey = key ? entries.find(e => e.key === key) : undefined;
      if (key && !byKey) throw usage(`No idempotency ledger entry for key "${key}" (entries are kept 24 h). Pass --internet-message-id instead.`);
      const id: string | undefined = opts.internetMessageId ?? (byKey && sentMessageId(byKey));
      if (!id) {
        throw usage(key
          ? `The ledger entry for "${key}" (${byKey?.command}) has no internetMessageId to look for. Pass --internet-message-id.`
          : 'Pass --idempotency-key or --internet-message-id.');
      }
      const client = getClient();
      let found: { folder: string; item: Record<string, unknown> } | undefined;
      for (const folder of ['sentitems', 'outbox', 'drafts']) {
        const response = await client.request(`/me/mailFolders/${folder}/messages`, {
          queryParams: { '$filter': `internetMessageId eq '${odataQuote(id)}'`, '$select': 'id,sentDateTime,webLink,internetMessageId', '$top': '1' },
        });
        if (client.isPreview(response)) continue;
        const value: unknown = Reflect.get(Object(response.data), 'value');
        const item: unknown = Array.isArray(value) ? value[0] : undefined;
        if (item && typeof item === 'object') {
          found = { folder, item: { ...item } };
          break;
        }
      }
      if (client.isDryRun()) return;
      const evidence = found && {
        id: found.item.id ?? null, folder: found.folder, sentDateTime: found.item.sentDateTime ?? null,
        webLink: found.item.webLink ?? null, internetMessageId: found.item.internetMessageId ?? id,
      };
      // An Outbox copy is accepted but not yet sent, so it is still NOT_FOUND_YET (with its evidence) (E18).
      const state = found?.folder === 'sentitems' ? 'SENT' : found?.folder === 'drafts' ? 'STILL_DRAFT' : 'NOT_FOUND_YET';
      const checkedAt = new Date().toISOString();
      const entry = byKey ?? entries.find(e => sentMessageId(e) === id);
      if (state === 'SENT' && evidence && entry) settleKey(entry.key, { ...evidence, checkedAt });
      console.log(formatOutput({ state, evidence: evidence ?? null, checkedAt }, opts.output as OutputFormat));
    }));

  addListOptions(mail
    .command('attachments <messageId>')
    .description('List attachments for a message'))
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts) => {
      await printList(getClient(), `/me/messages/${encodeURIComponent(messageId)}/attachments`, {}, opts);
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
          const res = await uploadLargeAttachment(client, messageId, opts.file, name, opts.contentType, size);
          const result = client.isPreview(res) ? res.data : { message: 'Attachment uploaded', name, size };
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
      await confirm(`Delete attachment ${attachmentId.substring(0, 16)}... from message ${messageId.substring(0, 16)}...?`);
      const client = getClient();
      await client.request(
        `/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
        { method: 'DELETE' }
      );
      success('Attachment deleted.');
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
      if (client.isPreview(response)) {
        // --dry-run: nothing to save (DR-13)
        console.log(formatOutput(response.data, 'json'));
        return;
      }
      const data = response.data as Record<string, unknown>;
      const name = data.name as string || 'attachment';
      const contentBytes = data.contentBytes as string;
      if (!contentBytes) {
        throw new CliError('No content found. This may be a reference attachment or item attachment.',
          { code: 'unsupported_attachment', fix: 'run `ms365 mail attachment <messageId> <attachmentId>` to see its type' });
      }
      const outPath = safeOutputPath(name, { out: opts.out, force: opts.force, fallback: 'attachment' });
      const bytes = Buffer.from(contentBytes, 'base64');
      fs.writeFileSync(outPath, bytes);
      success(`Saved to ${outPath} (${formatSize(bytes.length)})`);
    }));

  addBodyOptions(mail
    .command('reply <messageId>')
    .description('Reply to a message (--draft: create the reply draft and do not send it)'), 'Reply body', true)
    .option('--html', 'Body is HTML (otherwise plain text, sent escaped with its line breaks kept)')
    .option('--reply-all', 'Reply to all recipients')
    .option('--attach <path>', 'Attach a file (repeatable; over 3 MB total goes via a draft + upload session)', collect, [])
    .option('--draft', 'Create the reply draft (threaded, with --body, --html and --attach) and print it; never send. Needs Mail.ReadWrite, not Mail.Send')
    .option('--idempotency-key <key>', IDEMPOTENCY_HELP)
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts) => {
      const comment = bodyText(opts) ?? '';
      const files = prepareAttachments(opts.attach);
      const client = getClient();
      const base = `/me/messages/${encodeURIComponent(messageId)}`;
      const create = `${base}/${opts.replyAll ? 'createReplyAll' : 'createReply'}`;
      const html = !!opts.html;
      const request = { messageId, replyAll: !!opts.replyAll, comment, html, attachments: attachmentDigests(files), ...(opts.draft && { draft: true }) };
      // Graph's `comment` loses line breaks and is rejected together with message.body, so a body (plain text escaped,
      // or --html) is patched above the draft's quoted thread; only an empty plain body still goes as `comment`.
      const content = replyHtml(comment, html);
      await planTargetMessage(client, messageId, request);
      if (opts.draft) {
        const created = await idempotentDraft(client, opts.idempotencyKey, 'mail reply', request,
          () => prepareDraft(client, create, content === undefined ? { comment } : {}, files, content));
        if (created && !planning() && !client.isDryRun()) console.log(formatOutput(created, opts.output as OutputFormat));
        return;
      }
      const sent = await idempotentSend(client, opts.idempotencyKey, 'mail reply', request, async () => {
        if (content !== undefined) {
          await sendViaDraft(client, create, {}, files, content);
        } else if (needsUploadSession(files)) {
          await sendViaDraft(client, create, { comment }, files);
        } else {
          const body: Record<string, unknown> = { comment };
          if (files.length > 0) body.message = { attachments: files.map(inlineAttachment) };
          await client.request(`${base}/${opts.replyAll ? 'replyAll' : 'reply'}`, { method: 'POST', body: JSON.stringify(body) });
        }
      });
      if (sent && !planning()) success('Reply sent.');
    }));

  // ── Forward ──────────────────────────────────────────────────
  mail
    .command('forward <messageId>')
    .description('Forward a message (--draft: create the forward draft and do not send it)')
    .requiredOption('--to <addresses>', 'Comma-separated recipient emails')
    .option('--comment <text>', 'Comment to include with forward')
    .option('--html', 'Send --comment as HTML (otherwise plain text, sent escaped with its line breaks kept)')
    .option('--attach <path>', 'Attach a file (repeatable; over 3 MB total goes via a draft + upload session)', collect, [])
    .option('--draft', 'Create the forward draft (addressed to --to, with --comment, --html and --attach) and print it; never send. Needs Mail.ReadWrite, not Mail.Send')
    .option('--idempotency-key <key>', IDEMPOTENCY_HELP)
    .addOption(outputOption())
    .action(handleErrors(async (messageId: string, opts) => {
      const files = prepareAttachments(opts.attach);
      const client = getClient();
      const toRecipients = parseRecipients(opts.to);
      const base = `/me/messages/${encodeURIComponent(messageId)}`;
      const html = !!opts.html;
      const body: Record<string, unknown> = { comment: opts.comment || '', toRecipients };
      const request = { messageId, ...body, html, attachments: attachmentDigests(files), ...(opts.draft && { draft: true }) };
      const content = replyHtml(opts.comment || '', html);
      await planTargetMessage(client, messageId, request);
      if (opts.draft) {
        const created = await idempotentDraft(client, opts.idempotencyKey, 'mail forward', request,
          () => prepareDraft(client, `${base}/createForward`, content === undefined ? body : { toRecipients }, files, content));
        if (created && !planning() && !client.isDryRun()) console.log(formatOutput(created, opts.output as OutputFormat));
        return;
      }
      const sent = await idempotentSend(client, opts.idempotencyKey, 'mail forward', request, async () => {
        if (content !== undefined) {
          await sendViaDraft(client, `${base}/createForward`, { toRecipients }, files, content);
        } else if (needsUploadSession(files)) {
          await sendViaDraft(client, `${base}/createForward`, body, files);
        } else {
          if (files.length > 0) body.message = { attachments: files.map(inlineAttachment) };
          await client.request(`${base}/forward`, { method: 'POST', body: JSON.stringify(body) });
        }
      });
      if (sent && !planning()) success('Message forwarded.');
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
      success('Marked as read.');
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
      success('Marked as unread.');
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
      success('Message flagged.');
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
      success('Flag removed.');
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
      success('Flag marked complete.');
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
      success(`Categories set: ${categories.join(', ')}`);
    }));

  // Shared mailbox send (requires --org-mode); delegate reads are `mail list/get --mailbox` (AH-10).
  addBodyOptions(mail
    .command('shared-send <userId>')
    .description('Send email from a shared mailbox (requires --org-mode)')
    .requiredOption('--to <addresses>', 'Comma-separated recipient emails')
    .requiredOption('--subject <subject>', 'Email subject'), 'Email body', true)
    .option('--html', 'Body as HTML')
    .option('--idempotency-key <key>', IDEMPOTENCY_HELP)
    .action(handleErrors(async (userId: string, opts) => {
      const content = bodyText(opts) ?? '';
      const client = getClient();
      const message = {
        subject: opts.subject,
        body: { contentType: opts.html ? 'HTML' : 'Text', content },
        toRecipients: parseRecipients(opts.to),
      };
      const sent = await idempotentSend(client, opts.idempotencyKey, 'mail shared-send', { userId, message }, async () => {
        await client.request(`/users/${encodeURIComponent(userId)}/sendMail`, {
          method: 'POST',
          body: JSON.stringify({ message, saveToSentItems: true }),
        });
      });
      if (sent) success('Email sent from shared mailbox.');
    }));

  // ── Inbox rules ─────────────────────────────────────────────
  addListOptions(mail
    .command('rules')
    .description('List inbox rules'))
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      await printList(getClient(), '/me/mailFolders/inbox/messageRules', {}, opts);
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
    .option('--from <addresses>', 'Comma-separated strings to match in the sender (Graph senderContains, substring)')
    .option('--from-address <addresses>', 'Comma-separated exact sender addresses to match (Graph fromAddresses); excludes --from')
    .option('--subject-contains <words>', 'Comma-separated words to match in subject')
    .option('--move-to <name>', 'Move matching messages to this folder (name, path or ID)')
    .option('--forward-to <addresses>', 'Comma-separated addresses to forward to')
    .option('--delete', 'Delete matching messages')
    .option('--mark-read', 'Mark matching messages as read')
    .addOption(new Option('--importance <level>', 'Set importance').choices(['low', 'normal', 'high']))
    .option('--stop-processing', 'Stop processing subsequent rules')
    .option('--all-messages', 'Apply the rule to every incoming message (required when no condition is given)')
    .option('--disabled', 'Create the rule disabled')
    .option('--sequence <n>', 'Order among the inbox rules (1 runs first; default: after every existing rule)', intParser(1, 2_147_483_647))
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      if (opts.from && opts.fromAddress) {
        throw new CliError('--from and --from-address are mutually exclusive.',
          { code: 'usage', exitCode: ExitCode.USAGE, fix: 'use --from-address for exact senders or --from for substring matches, not both' });
      }
      const conditions: Record<string, unknown> = {};
      if (opts.fromAddress) {
        conditions.fromAddresses = opts.fromAddress.split(',').map((s: string) => s.trim()).filter(Boolean)
          .map((address: string) => ({ emailAddress: { address } }));
      }
      if (opts.from) {
        conditions.senderContains = opts.from.split(',').map((s: string) => s.trim());
      }
      if (opts.subjectContains) {
        conditions.subjectContains = opts.subjectContains.split(',').map((s: string) => s.trim());
      }

      if (Object.keys(conditions).length === 0 && !opts.allMessages) {
        throw new CliError('Refusing to create a rule with no conditions: it would act on every incoming message.',
          { code: 'usage', exitCode: ExitCode.USAGE, fix: 'add --from-address, --from or --subject-contains, or pass --all-messages to confirm' });
      }

      const actions: Record<string, unknown> = {};
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

      const client = getClient();
      if (opts.moveTo) {
        actions.moveToFolder = await resolve(client, 'mail-folder', opts.moveTo);
      }

      const rule: Record<string, unknown> = {
        displayName: opts.name,
        sequence: opts.sequence === undefined ? await nextRuleSequence(client) : Number(opts.sequence),
        conditions,
        actions,
        isEnabled: !opts.disabled,
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
    .addOption(new Option('--enabled <bool>', 'Enable/disable the rule').choices(['true', 'false']))
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
      const client = getClient();
      const results = await client.batch(ids.map(id => ({
        method: 'PATCH', url: `/me/messages/${encodeURIComponent(id)}`, body: { isRead: true },
      })));
      reportBatch(client, ids, results, 'Marked', 'as read');
    }));

  mail
    .command('batch-move <messageIds>')
    .description('Move multiple messages to a folder (comma-separated IDs; sent via JSON batching)')
    .requiredOption('--folder <name>', 'Destination folder name, path or ID')
    .action(handleErrors(async (messageIds: string, opts) => {
      const client = getClient();
      const folderId = await resolve(client, 'mail-folder', opts.folder);
      const ids = messageIds.split(',').map(id => id.trim()).filter(Boolean);
      const results = await client.batch(ids.map(id => ({
        method: 'POST', url: `/me/messages/${encodeURIComponent(id)}/move`, body: { destinationId: folderId },
      })));
      reportBatch(client, ids, results, 'Moved', '');
    }));

}

/** Summarize batch results on stderr; fail (exit 1) when any item failed. */
function reportBatch(client: GraphClient, ids: string[], results: BatchResult[], verb: string, suffix: string): void {
  const failed = ids
    .map((id, i) => ({ id, r: results[i] }))
    .filter(({ r }) => !client.isPreview(r) && (r.status < 200 || r.status >= 300));
  success(`${verb} ${ids.length - failed.length} message(s)${suffix ? ' ' + suffix : ''}.`);
  if (failed.length > 0) {
    const detail = failed.map(({ id, r }) => {
      const msg = (r.body as { error?: { message?: string } } | null)?.error?.message ?? `status ${r.status}`;
      return `${id.substring(0, 12)}...: ${msg}`;
    }).join('; ');
    throw new CliError(`Failed ${failed.length} of ${ids.length}: ${detail}`, { code: 'partial_failure' });
  }
}
