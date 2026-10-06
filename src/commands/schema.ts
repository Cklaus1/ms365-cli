import { Command } from 'commander';
import { handleErrors } from '../errors.js';

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
    description: 'List messages in the signed-in user\'s mailbox.',
    scopes: ['Mail.Read'],
    parameters: [
      { name: '$top', type: 'integer', location: 'query', required: false, description: 'Number of items to return (max 1000)' },
      { name: '$filter', type: 'string', location: 'query', required: false, description: 'OData filter expression' },
      { name: '$select', type: 'string', location: 'query', required: false, description: 'Comma-separated properties to include' },
      { name: '$orderby', type: 'string', location: 'query', required: false, description: 'Sort order (e.g. "receivedDateTime desc")' },
      { name: '$search', type: 'string', location: 'query', required: false, description: 'Full-text search query' },
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
    description: 'Upload or replace a file. Files > 4MB use upload sessions automatically.',
    scopes: ['Files.ReadWrite'],
    parameters: [
      { name: 'drive-id', type: 'string', location: 'path', required: true, description: 'Drive ID' },
      { name: 'parent-id', type: 'string', location: 'path', required: true, description: 'Parent folder item ID' },
      { name: 'filename', type: 'string', location: 'path', required: true, description: 'File name' },
      { name: 'Content-Type', type: 'string', location: 'header', required: true, description: 'MIME type (auto-detected from extension)' },
    ],
  },
  'todo.lists': {
    method: 'GET', path: '/me/todo/lists',
    description: 'List To Do task lists.',
    scopes: ['Tasks.Read'],
    parameters: [],
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
      { name: 'importance', type: 'string', location: 'body', required: false, description: '"low" | "normal" | "high"' },
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
      { name: '$search', type: 'string', location: 'query', required: false, description: 'Search query' },
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
      { name: 'requests', type: 'array', location: 'body', required: true, description: '[{ entityTypes: ["message"|"driveItem"|"site"|"event"|"chatMessage"], query: { queryString: string }, from: number, size: number }]' },
    ],
  },
};

export function registerSchemaCommands(program: Command): void {
  program
    .command('schema <endpoint>')
    .description('Introspect API schema for an endpoint (e.g. mail.list, calendar.create)')
    .option('--list', 'List all available endpoint schemas')
    .action(handleErrors(async (endpoint: string, opts) => {
      if (opts.list || endpoint === 'list') {
        console.log('Available schemas:\n');
        const grouped: Record<string, string[]> = {};
        for (const [key, val] of Object.entries(ENDPOINTS)) {
          const [service] = key.split('.');
          if (!grouped[service]) grouped[service] = [];
          grouped[service].push(`  ${key.padEnd(25)} ${val.method.padEnd(6)} ${val.path}`);
        }
        for (const [service, entries] of Object.entries(grouped)) {
          console.log(`${service}:`);
          entries.forEach(e => console.log(e));
          console.log('');
        }
        console.log(`Usage: ms365 schema <endpoint>  (e.g. ms365 schema mail.send)`);
        return;
      }

      const schema = ENDPOINTS[endpoint];
      if (!schema) {
        const suggestions = Object.keys(ENDPOINTS)
          .filter(k => k.startsWith(endpoint.split('.')[0]))
          .join(', ');
        console.error(`Unknown endpoint: ${endpoint}`);
        if (suggestions) {
          console.error(`Did you mean: ${suggestions}`);
        }
        console.error(`\nRun: ms365 schema --list`);
        process.exit(1);
      }

      const output: Record<string, unknown> = {
        httpMethod: schema.method,
        path: schema.path,
        description: schema.description,
        scopes: schema.scopes,
        parameters: {} as Record<string, unknown>,
      };

      for (const param of schema.parameters) {
        (output.parameters as Record<string, unknown>)[param.name] = {
          type: param.type,
          required: param.required,
          location: param.location,
          description: param.description,
        };
      }

      if (schema.responseHint) {
        output.response = schema.responseHint;
      }

      console.log(JSON.stringify(output, null, 2));
    }));
}
