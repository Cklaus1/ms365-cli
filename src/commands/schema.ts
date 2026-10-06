import { Command } from 'commander';
import { handleErrors } from '../errors.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption } from '../utils.js';

/**
 * Endpoint catalog with parameter schemas for introspection.
 * This is a curated, static catalog — not fetched from OpenAPI at runtime
 * (Graph's OpenAPI spec is 50MB+, impractical for CLI).
 */
interface ParamDef {
  name: string;
  type: string;
  location: 'path' | 'query' | 'body' | 'header';
  required: boolean;
  description: string;
}

interface EndpointSchema {
  method: string;
  path: string;
  description: string;
  scopes: string[];
  parameters: ParamDef[];
  responseHint?: string;
}

const ENDPOINTS: Record<string, EndpointSchema> = {
  'mail.list': {
    method: 'GET', path: '/me/messages',
    description: 'List messages across all folders of the signed-in user\'s mailbox (`--folder` uses /me/mailFolders/{folder}/messages).',
    scopes: ['Mail.Read'],
    parameters: [
      { name: '$top', type: 'integer', location: 'query', required: false, description: 'Number of items to return (max 1000)' },
      { name: '$filter', type: 'string', location: 'query', required: false, description: 'OData filter expression' },
      { name: '$select', type: 'string', location: 'query', required: false, description: 'Comma-separated properties to include' },
      { name: '$orderby', type: 'string', location: 'query', required: false, description: 'Sort order (e.g. "receivedDateTime desc")' },
      { name: '$search', type: 'string', location: 'query', required: false, description: 'Full-text search; the CLI sends it as one double-quoted phrase with \\ and " backslash-escaped' },
      { name: '$skip', type: 'integer', location: 'query', required: false, description: 'Number of items to skip' },
      { name: '$count', type: 'boolean', location: 'query', required: false, description: 'Include count of results' },
    ],
    responseHint: '{ value: Message[], @odata.nextLink?: string }',
  },
  'mail.get': {
    method: 'GET', path: '/me/messages/{message-id}',
    description: 'Get a specific message by ID.',
    scopes: ['Mail.Read'],
    parameters: [
      { name: 'message-id', type: 'string', location: 'path', required: true, description: 'Message ID' },
      { name: '$select', type: 'string', location: 'query', required: false, description: 'Properties to include' },
    ],
    responseHint: 'Message object with subject, from, body, hasAttachments, categories, flag, etc.',
  },
  'mail.send': {
    method: 'POST', path: '/me/sendMail',
    description: 'Send a new email message.',
    scopes: ['Mail.Send'],
    parameters: [
      { name: 'message', type: 'object', location: 'body', required: true, description: '{ subject, body: { contentType, content }, toRecipients: [{ emailAddress: { address } }] }' },
      { name: 'saveToSentItems', type: 'boolean', location: 'body', required: false, description: 'Save to Sent Items folder (default true)' },
    ],
  },
  'calendar.events': {
    method: 'GET', path: '/me/events',
    description: 'List calendar events for the signed-in user.',
    scopes: ['Calendars.Read'],
    parameters: [
      { name: '$top', type: 'integer', location: 'query', required: false, description: 'Number of events to return' },
      { name: '$filter', type: 'string', location: 'query', required: false, description: 'OData filter' },
      { name: '$select', type: 'string', location: 'query', required: false, description: 'Properties to include' },
      { name: '$orderby', type: 'string', location: 'query', required: false, description: 'Sort order' },
      { name: 'Prefer', type: 'string', location: 'header', required: false, description: 'outlook.timezone="<IANA zone>": the CLI asks for times in the local zone' },
    ],
    responseHint: '{ value: Event[] }',
  },
  'calendar.view': {
    method: 'GET', path: '/me/calendarView',
    description: 'Get calendar events within a date range (expands recurring events).',
    scopes: ['Calendars.Read'],
    parameters: [
      { name: 'startDateTime', type: 'string', location: 'query', required: true, description: 'ISO 8601 start datetime' },
      { name: 'endDateTime', type: 'string', location: 'query', required: true, description: 'ISO 8601 end datetime' },
      { name: '$select', type: 'string', location: 'query', required: false, description: 'Properties to include' },
      { name: '$top', type: 'integer', location: 'query', required: false, description: 'Events per page' },
      { name: 'Prefer', type: 'string', location: 'header', required: false, description: 'outlook.timezone="<IANA zone>": the CLI asks for times in the local zone' },
    ],
    responseHint: '{ value: Event[] }',
  },
  'calendar.create': {
    method: 'POST', path: '/me/events',
    description: 'Create a new calendar event.',
    scopes: ['Calendars.ReadWrite'],
    parameters: [
      { name: 'subject', type: 'string', location: 'body', required: true, description: 'Event subject/title' },
      { name: 'start', type: 'object', location: 'body', required: true, description: '{ dateTime: string, timeZone: string }' },
      { name: 'end', type: 'object', location: 'body', required: true, description: '{ dateTime: string, timeZone: string }' },
      { name: 'body', type: 'object', location: 'body', required: false, description: '{ contentType: "Text"|"HTML", content: string }' },
      { name: 'location', type: 'object', location: 'body', required: false, description: '{ displayName: string }' },
      { name: 'attendees', type: 'array', location: 'body', required: false, description: '[{ emailAddress: { address }, type: "required"|"optional" }]' },
    ],
  },
  'drive.list': {
    method: 'GET', path: '/me/drives',
    description: 'List available drives for the signed-in user.',
    scopes: ['Files.Read'],
    parameters: [],
    responseHint: '{ value: Drive[] }',
  },
  'drive.files': {
    method: 'GET', path: '/drives/{drive-id}/items/{item-id}/children',
    description: 'List children of a drive item (folder).',
    scopes: ['Files.Read'],
    parameters: [
      { name: 'drive-id', type: 'string', location: 'path', required: true, description: 'Drive ID' },
      { name: 'item-id', type: 'string', location: 'path', required: true, description: 'Folder item ID' },
      { name: '$top', type: 'integer', location: 'query', required: false, description: 'Number of items' },
      { name: '$filter', type: 'string', location: 'query', required: false, description: 'OData filter' },
      { name: '$select', type: 'string', location: 'query', required: false, description: 'Properties to include' },
    ],
    responseHint: '{ value: DriveItem[] }',
  },
  'drive.upload': {
    method: 'PUT', path: '/drives/{drive-id}/items/{parent-id}:/{filename}:/content',
    description: 'Upload a file into a folder. Files > 4 MiB use an upload session (POST …:/createUploadSession with item.@microsoft.graph.conflictBehavior).',
    scopes: ['Files.ReadWrite'],
    parameters: [
      { name: 'drive-id', type: 'string', location: 'path', required: true, description: 'Drive ID' },
      { name: 'parent-id', type: 'string', location: 'path', required: true, description: 'Parent folder item ID' },
      { name: 'filename', type: 'string', location: 'path', required: true, description: 'File name (percent-encoded)' },
      { name: '@microsoft.graph.conflictBehavior', type: 'string', location: 'query', required: false, description: '"rename" | "replace" | "fail" (CLI --conflict, default rename)' },
      { name: 'Content-Type', type: 'string', location: 'header', required: true, description: 'MIME type (auto-detected from extension)' },
    ],
  },
  'todo.lists': {
    method: 'GET', path: '/me/todo/lists',
    description: 'List To Do task lists.',
    scopes: ['Tasks.Read'],
    parameters: [
      { name: '$top', type: 'integer', location: 'query', required: false, description: 'Lists per page' },
    ],
    responseHint: '{ value: TodoTaskList[] }',
  },
  'todo.tasks': {
    method: 'GET', path: '/me/todo/lists/{list-id}/tasks',
    description: 'List tasks in a To Do list.',
    scopes: ['Tasks.Read'],
    parameters: [
      { name: 'list-id', type: 'string', location: 'path', required: true, description: 'Task list ID' },
      { name: '$top', type: 'integer', location: 'query', required: false, description: 'Number of tasks' },
      { name: '$filter', type: 'string', location: 'query', required: false, description: 'OData filter' },
    ],
    responseHint: '{ value: TodoTask[] }',
  },
  'todo.create': {
    method: 'POST', path: '/me/todo/lists/{list-id}/tasks',
    description: 'Create a new task in a To Do list.',
    scopes: ['Tasks.ReadWrite'],
    parameters: [
      { name: 'list-id', type: 'string', location: 'path', required: true, description: 'Task list ID' },
      { name: 'title', type: 'string', location: 'body', required: true, description: 'Task title' },
      { name: 'body', type: 'object', location: 'body', required: false, description: '{ contentType: "text", content: string }' },
      { name: 'dueDateTime', type: 'object', location: 'body', required: false, description: '{ dateTime: string, timeZone: string }' },
      { name: 'importance', type: 'string', location: 'body', required: false, description: '"low" | "normal" | "high" (validated by the CLI)' },
    ],
  },
  'contacts.list': {
    method: 'GET', path: '/me/contacts',
    description: 'List Outlook contacts.',
    scopes: ['Contacts.Read'],
    parameters: [
      { name: '$top', type: 'integer', location: 'query', required: false, description: 'Number of contacts' },
      { name: '$filter', type: 'string', location: 'query', required: false, description: 'OData filter' },
      { name: '$select', type: 'string', location: 'query', required: false, description: 'Properties to include' },
      { name: '$search', type: 'string', location: 'query', required: false, description: 'Search query; the CLI sends it as one double-quoted phrase with \\ and " backslash-escaped' },
    ],
    responseHint: '{ value: Contact[] }',
  },
  'teams.list': {
    method: 'GET', path: '/me/joinedTeams',
    description: 'List teams the user is a member of.',
    scopes: ['Team.ReadBasic.All'],
    parameters: [],
    responseHint: '{ value: Team[] }',
  },
  'chat.send': {
    method: 'POST', path: '/chats/{chat-id}/messages',
    description: 'Send a message to a chat.',
    scopes: ['ChatMessage.Send'],
    parameters: [
      { name: 'chat-id', type: 'string', location: 'path', required: true, description: 'Chat ID' },
      { name: 'body', type: 'object', location: 'body', required: true, description: '{ contentType: "text"|"html", content: string }' },
    ],
  },
  'planner.tasks': {
    method: 'GET', path: '/planner/plans/{plan-id}/tasks',
    description: 'List tasks in a Planner plan.',
    scopes: ['Tasks.Read'],
    parameters: [
      { name: 'plan-id', type: 'string', location: 'path', required: true, description: 'Plan ID' },
    ],
    responseHint: '{ value: PlannerTask[] }',
  },
  'search.query': {
    method: 'POST', path: '/search/query',
    description: 'Search across Microsoft 365 content.',
    scopes: ['Mail.Read', 'Files.Read.All'],
    parameters: [
      { name: 'requests', type: 'array', location: 'body', required: true, description: '[{ entityTypes, query: { queryString: string }, from: number, size: number (1-1000) }]. entityTypes must come from one group: driveItem,drive,list,listItem,site,externalItem | message | event | chatMessage | person | acronym,bookmark,qna' },
    ],
  },
  'mail.triage': {
    method: 'GET', path: '/me/mailFolders/{folder}/messages',
    description: 'Inbox (or --folder) summary, newest first. The filter always leads with receivedDateTime so $orderby is accepted.',
    scopes: ['Mail.Read'],
    parameters: [
      { name: 'folder', type: 'string', location: 'path', required: true, description: 'Folder ID or well-known name (default inbox)' },
      { name: '$filter', type: 'string', location: 'query', required: true, description: 'receivedDateTime ge 1900-01-01T00:00:00Z and (<--filter, default isRead eq false>)' },
      { name: '$orderby', type: 'string', location: 'query', required: false, description: 'receivedDateTime desc' },
      { name: '$select', type: 'string', location: 'query', required: false, description: 'Summary properties' },
      { name: '$top', type: 'integer', location: 'query', required: false, description: 'Messages per page (default 20)' },
    ],
    responseHint: '{ value: Message[] }',
  },
  'mail.download-attachment': {
    method: 'GET', path: '/me/messages/{message-id}/attachments/{attachment-id}/$value',
    description: 'Download raw attachment bytes. The CLI first GETs the attachment with $select=id,name,contentType,size; reference attachments are refused.',
    scopes: ['Mail.Read'],
    parameters: [
      { name: 'message-id', type: 'string', location: 'path', required: true, description: 'Message ID' },
      { name: 'attachment-id', type: 'string', location: 'path', required: true, description: 'Attachment ID' },
    ],
    responseHint: 'Raw bytes',
  },
  'mail.create-rule': {
    method: 'POST', path: '/me/mailFolders/inbox/messageRules',
    description: 'Create an inbox rule. Without --sequence the CLI first GETs the rules and appends; a rule with no conditions needs interactive confirmation.',
    scopes: ['MailboxSettings.ReadWrite'],
    parameters: [
      { name: 'displayName', type: 'string', location: 'body', required: true, description: 'Rule name' },
      { name: 'sequence', type: 'integer', location: 'body', required: true, description: 'Order among rules' },
      { name: 'actions', type: 'object', location: 'body', required: true, description: 'At least one of moveToFolder, forwardTo, delete, markAsRead, markImportance, stopProcessingRules' },
      { name: 'conditions', type: 'object', location: 'body', required: false, description: '{ fromAddresses?: Recipient[], subjectContains?: string[] }' },
      { name: 'isEnabled', type: 'boolean', location: 'body', required: false, description: 'Default true (--no-enabled sends false)' },
    ],
  },
  'calendar.rooms': {
    method: 'GET', path: '/me/findRooms',
    description: 'List rooms (Graph beta; work/school accounts).',
    scopes: ['User.Read'],
    parameters: [],
    responseHint: '{ value: EmailAddress[] } (beta)',
  },
  'calendar.room-lists': {
    method: 'GET', path: '/me/findRoomLists',
    description: 'List room lists (Graph beta; work/school accounts).',
    scopes: ['User.Read'],
    parameters: [],
    responseHint: '{ value: EmailAddress[] } (beta)',
  },
  'drive.cat': {
    method: 'GET', path: '/me/drive/root:/{path}:/content',
    description: 'Download a file by path; raw bytes to stdout or --out.',
    scopes: ['Files.Read'],
    parameters: [
      { name: 'path', type: 'string', location: 'path', required: true, description: 'File path, each segment percent-encoded' },
    ],
    responseHint: 'Raw bytes',
  },
  'drive.cp': {
    method: 'PUT', path: '/me/drive/root:/{path}:/content',
    description: 'Upload a local file by path (a path ending in / is a folder). Files > 4 MiB use an upload session.',
    scopes: ['Files.ReadWrite'],
    parameters: [
      { name: 'path', type: 'string', location: 'path', required: true, description: 'Target path, each segment percent-encoded' },
      { name: '@microsoft.graph.conflictBehavior', type: 'string', location: 'query', required: false, description: '"rename" | "replace" | "fail" (CLI --conflict, default replace)' },
      { name: 'Content-Type', type: 'string', location: 'header', required: true, description: 'MIME type (auto-detected from extension)' },
    ],
  },
  'excel.format-range': {
    method: 'PATCH', path: '/drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/range(address=\'{address}\')/format/font',
    description: 'Format a range: PATCH …/format/font (bold, color) and/or PATCH …/format/fill (color).',
    scopes: ['Files.ReadWrite'],
    parameters: [
      { name: 'address', type: 'string', location: 'path', required: true, description: 'Range address, e.g. A1:D10' },
      { name: 'bold', type: 'boolean', location: 'body', required: false, description: 'Font bold (font request)' },
      { name: 'color', type: 'string', location: 'body', required: false, description: 'Font color (font request) or fill color (fill request)' },
    ],
  },
  'excel.sort-range': {
    method: 'POST', path: '/drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/range(address=\'{address}\')/sort/apply',
    description: 'Sort a range.',
    scopes: ['Files.ReadWrite'],
    parameters: [
      { name: 'address', type: 'string', location: 'path', required: true, description: 'Range address, e.g. A1:D10' },
      { name: 'fields', type: 'array', location: 'body', required: true, description: '[{ key: number, ascending: boolean }]' },
      { name: 'matchCase', type: 'boolean', location: 'body', required: false, description: 'Case-sensitive sort (default false)' },
    ],
  },
  'sharepoint.delta': {
    method: 'GET', path: '/sites/delta()',
    description: 'Changed sites. Prints { value, nextLink | deltaLink }; --delta-link continues from a graph.microsoft.com /sites/delta() link.',
    scopes: ['Sites.Read.All'],
    parameters: [
      { name: '$top', type: 'integer', location: 'query', required: false, description: 'Sites per page' },
    ],
    responseHint: '{ value: Site[], nextLink?: string, deltaLink?: string }',
  },
  'planner.create-plan': {
    method: 'POST', path: '/planner/plans',
    description: 'Create a plan in a group.',
    scopes: ['Tasks.ReadWrite'],
    parameters: [
      { name: 'title', type: 'string', location: 'body', required: true, description: 'Plan title' },
      { name: 'container', type: 'object', location: 'body', required: true, description: '{ containerId: <group-id>, type: "group" }' },
    ],
  },
  'onenote.page': {
    method: 'GET', path: '/me/onenote/pages/{page-id}/content',
    description: 'Page content as raw HTML (stdout or --out).',
    scopes: ['Notes.Read'],
    parameters: [
      { name: 'page-id', type: 'string', location: 'path', required: true, description: 'Page ID' },
    ],
    responseHint: 'text/html',
  },
  'chat.reply': {
    method: 'POST', path: '/chats/{chat-id}/messages/replyWithQuote',
    description: 'Post a new chat message quoting an earlier one (chats have no threaded replies).',
    scopes: ['ChatMessage.Send'],
    parameters: [
      { name: 'chat-id', type: 'string', location: 'path', required: true, description: 'Chat ID' },
      { name: 'messageIds', type: 'array', location: 'body', required: true, description: '[<message-id>]' },
      { name: 'replyMessage', type: 'object', location: 'body', required: true, description: '{ body: { contentType: "text"|"html", content: string } }' },
    ],
  },
  'user.set-ooo': {
    method: 'PATCH', path: '/me/mailboxSettings',
    description: 'Set automatic replies. externalAudience is sent only when --audience is given; scheduled times are sent in UTC.',
    scopes: ['MailboxSettings.ReadWrite'],
    parameters: [
      { name: 'automaticRepliesSetting', type: 'object', location: 'body', required: true, description: '{ status: "alwaysEnabled"|"scheduled"|"disabled", externalAudience?: "all"|"contactsOnly"|"none", internalReplyMessage?, externalReplyMessage?, scheduledStartDateTime?/scheduledEndDateTime?: { dateTime, timeZone: "UTC" } }' },
    ],
  },
  'user.set-presence': {
    method: 'POST', path: '/me/presence/setUserPreferredPresence',
    description: 'Set your preferred presence (shown while a Teams/app presence session exists).',
    scopes: ['Presence.ReadWrite'],
    parameters: [
      { name: 'availability', type: 'string', location: 'body', required: true, description: 'Available | Busy | DoNotDisturb | BeRightBack | Away | Offline' },
      { name: 'activity', type: 'string', location: 'body', required: true, description: 'Must match availability (Offline -> OffWork, otherwise the same name); the CLI derives it' },
      { name: 'expirationDuration', type: 'string', location: 'body', required: false, description: 'ISO 8601 duration (CLI default PT60M)' },
    ],
  },
  'user.clear-presence': {
    method: 'POST', path: '/me/presence/clearUserPreferredPresence',
    description: 'Clear your preferred presence (revert to automatic).',
    scopes: ['Presence.ReadWrite'],
    parameters: [],
  },
  'user.download-photo': {
    method: 'GET', path: '/me/photo/$value',
    description: 'Profile photo bytes; --out defaults to profile-photo.<ext from content type>, never overwritten without --force.',
    scopes: ['User.Read'],
    parameters: [],
    responseHint: 'Raw image bytes',
  },
  'groups.distribution-lists': {
    method: 'GET', path: '/groups',
    description: 'Mail-enabled, non-security groups; Microsoft 365 (Unified) groups are dropped client-side.',
    scopes: ['Group.Read.All'],
    parameters: [
      { name: '$filter', type: 'string', location: 'query', required: true, description: 'mailEnabled eq true and securityEnabled eq false' },
      { name: '$select', type: 'string', location: 'query', required: false, description: 'id,displayName,mail,description,membershipRule,groupTypes' },
      { name: '$top', type: 'integer', location: 'query', required: false, description: 'Groups per page (default 25)' },
    ],
    responseHint: '{ value: Group[] }',
  },
};

/** Endpoint keys with a schema (`mail.list`, …). */
export function endpointKeys(): string[] {
  return Object.keys(ENDPOINTS);
}

function editDistance(a: string, b: string): number {
  let prev = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    const row = [i];
    for (let j = 1; j <= b.length; j++) {
      row[j] = Math.min(prev[j] + 1, row[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = row;
  }
  return prev[b.length];
}

/** Keys of the same service, else keys within a small edit distance. */
function suggestEndpoints(endpoint: string): string[] {
  const keys = endpointKeys();
  const service = endpoint.split('.')[0];
  const sameService = keys.filter(k => k.split('.')[0] === service);
  if (sameService.length > 0) return sameService;
  return keys.filter(k => editDistance(k, endpoint) <= 3);
}

export function registerSchemaCommands(program: Command): void {
  program
    .command('schema [endpoint]')
    .description('Introspect API schema for an endpoint (e.g. mail.list, calendar.create)')
    .option('--list', 'List all available endpoint schemas')
    .addOption(outputOption('json'))
    .action(handleErrors(async (endpoint: string | undefined, opts: { list?: boolean; output: OutputFormat }) => {
      if (opts.list || endpoint === 'list') {
        const list = Object.entries(ENDPOINTS).map(([key, val]) => ({
          endpoint: key,
          method: val.method,
          path: val.path,
          description: val.description,
        }));
        console.log(formatOutput(list, opts.output));
        return;
      }

      if (!endpoint) throw new Error('Specify an endpoint (e.g. ms365 schema mail.list) or run: ms365 schema --list');

      const schema = ENDPOINTS[endpoint];
      if (!schema) {
        const suggestions = suggestEndpoints(endpoint);
        const hint = suggestions.length > 0 ? `\nDid you mean: ${suggestions.join(', ')}` : '';
        throw new Error(`Unknown endpoint: ${endpoint}${hint}\nRun: ms365 schema --list`);
      }

      const output: Record<string, unknown> = {
        httpMethod: schema.method,
        path: schema.path,
        description: schema.description,
        scopes: schema.scopes,
        parameters: Object.fromEntries(schema.parameters.map(p => [p.name, {
          type: p.type,
          required: p.required,
          location: p.location,
          description: p.description,
        }])),
      };
      if (schema.responseHint) output.response = schema.responseHint;

      console.log(formatOutput(output, opts.output));
    }));
}
