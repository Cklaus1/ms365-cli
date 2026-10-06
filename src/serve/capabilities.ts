/**
 * Command → capability table for `ms365 serve` (REMOTE-1c). Keys are canonical dotted command paths
 * (`buildCommandTree`). A command runs remotely only when it is mapped here and the caller holds every listed
 * capability; `remote: false` marks commands that never run remotely. A test fails when a registered command is
 * missing from the table, so every new command needs a conscious decision.
 *
 * The same table records local file arguments, which the server and `ms365-remote` translate:
 * - `input`: index of the positional naming a local file to read (uploaded as a multipart part).
 * - `output`: the command writes a local file via `--out` (and `--force`). `stdout` commands print the bytes
 *   when `--out` is omitted; `named` commands then write a file named by the CLI into the current directory.
 */

export const CAPABILITIES = [
  'mail.read', 'mail.draft', 'mail.send', 'mail.manage',
  'calendar.read', 'calendar.write',
  'drive.read', 'drive.write',
  'sharepoint.read',
  'teams.read', 'teams.send',
  'planner.read', 'planner.write',
  'todo.read', 'todo.write',
  'contacts.read', 'contacts.write',
  'excel.read', 'excel.write',
  'onenote.read', 'onenote.write',
  'user.read', 'user.write',
  'groups.read',
  'search.read',
  'meta',
] as const;

export type Capability = (typeof CAPABILITIES)[number];

export type OutputMode = 'stdout' | 'named';

export interface RemoteCommand {
  remote: true;
  /** Capabilities the caller must hold (all of them). */
  caps: readonly Capability[];
  /** Index of the positional that names a local input file. */
  input?: number;
  /** The command writes a local output file through `--out`. */
  output?: OutputMode;
}

export interface LocalOnlyCommand {
  remote: false;
}

export type CommandPolicy = RemoteCommand | LocalOnlyCommand;

const LOCAL: LocalOnlyCommand = { remote: false };
const cap = (...caps: Capability[]): RemoteCommand => ({ remote: true, caps });

/** Map every leaf of `group` listed in `names` to `caps`. */
function group(prefix: string, caps: Capability[], names: string[]): Record<string, CommandPolicy> {
  return Object.fromEntries(names.map(n => [`${prefix}.${n}`, cap(...caps)]));
}

export const COMMAND_POLICY: Readonly<Record<string, CommandPolicy>> = {
  ...Object.fromEntries(['login', 'logout', 'status', 'accounts', 'select', 'remove', 'setup'].map(n => [`auth.${n}`, LOCAL])),
  'generate-skills': LOCAL,
  serve: LOCAL,
  schema: cap('meta'),

  ...group('mail', ['mail.read'], [
    'list', 'get', 'read', 'triage', 'folders', 'folder', 'drafts', 'attachments', 'attachment',
    'shared-list', 'shared-get', 'rules', 'rule',
  ]),
  'mail.download-attachment': { ...cap('mail.read'), output: 'named' },
  ...group('mail', ['mail.draft'], ['draft', 'update-draft', 'add-attachment', 'delete-attachment']),
  ...group('mail', ['mail.send'], ['send', 'send-draft', 'reply', 'forward', 'shared-send']),
  ...group('mail', ['mail.manage'], [
    'delete', 'move', 'create-folder', 'mark-read', 'mark-unread', 'flag', 'unflag', 'complete', 'categorize',
    'create-rule', 'update-rule', 'batch-read', 'batch-move',
  ]),

  ...group('calendar', ['calendar.read'], ['list', 'events', 'get', 'view', 'free-busy', 'rooms', 'room-lists']),
  ...group('calendar', ['calendar.write'], [
    'create', 'update', 'delete', 'accept', 'decline', 'tentative', 'create-calendar', 'update-calendar',
  ]),

  ...group('drive', ['drive.read'], ['list', 'root', 'files', 'search', 'permissions', 'quota', 'info', 'ls']),
  'drive.download': { ...cap('drive.read'), output: 'stdout' },
  'drive.download-shared': { ...cap('drive.read'), output: 'stdout' },
  'drive.cat': { ...cap('drive.read'), output: 'stdout' },
  ...group('drive', ['drive.write'], ['delete', 'mkdir', 'share']),
  'drive.upload': { ...cap('drive.write'), input: 2 },
  'drive.cp': { ...cap('drive.write'), input: 0 },

  ...group('todo', ['todo.read'], ['lists', 'tasks', 'get']),
  ...group('todo', ['todo.write'], ['create', 'update', 'delete', 'create-list', 'update-list']),

  ...group('contacts', ['contacts.read'], ['list', 'get', 'folders', 'folder-contacts']),
  ...group('contacts', ['contacts.write'], ['create', 'update', 'delete', 'create-folder', 'update-folder']),

  ...group('teams', ['teams.read'], ['list', 'get', 'members', 'channels', 'channel', 'channel-messages']),
  ...group('teams', ['teams.send'], ['send-channel-message', 'reply-channel-message']),
  ...group('chat', ['teams.read'], ['list', 'get', 'messages']),
  ...group('chat', ['teams.send'], ['send', 'reply']),

  ...group('sharepoint', ['sharepoint.read'], [
    'search', 'site', 'site-by-path', 'drives', 'drive', 'lists', 'list', 'list-items', 'list-item', 'item', 'delta',
  ]),

  ...group('planner', ['planner.read'], [
    'plans', 'plan', 'plan-details', 'my-tasks', 'tasks', 'task', 'task-details', 'buckets', 'bucket', 'bucket-tasks',
  ]),
  ...group('planner', ['planner.write'], [
    'create-task', 'update-task', 'update-plan-details', 'update-task-details', 'delete-task', 'delete-plan',
    'delete-bucket', 'update-plan', 'update-bucket', 'create-plan', 'create-bucket',
  ]),

  ...group('onenote', ['onenote.read'], ['notebooks', 'sections', 'pages']),
  'onenote.page': { ...cap('onenote.read'), output: 'stdout' },
  ...group('onenote', ['onenote.write'], ['create-page']),

  ...group('excel', ['excel.read'], ['worksheets', 'range']),
  ...group('excel', ['excel.write'], ['create-chart', 'format-range', 'sort-range']),

  ...group('search', ['search.read'], ['query', 'people']),

  ...group('user', ['user.read'], ['me', 'list', 'ooo', 'presence', 'photo', 'mailbox-settings', 'working-hours']),
  'user.download-photo': { ...cap('user.read'), output: 'named' },
  ...group('user', ['user.write'], ['set-ooo', 'set-presence', 'clear-presence', 'set-timezone', 'set-language']),
  'user.upload-photo': { ...cap('user.write'), input: 0 },

  // Workflows combine several services: the caller needs every capability they touch.
  'workflow.standup': cap('calendar.read', 'mail.read'),
  'workflow.meeting-prep': cap('calendar.read'),
  'workflow.email-to-task': cap('mail.read', 'todo.write'),
  'workflow.digest': cap('calendar.read', 'mail.read', 'todo.read'),
  'workflow.file-announce': cap('drive.read', 'teams.send'),
  'workflow.focus-time': cap('calendar.read'),

  ...group('groups', ['groups.read'], ['list', 'my', 'get', 'members', 'owners', 'distribution-lists', 'teams']),
};

/** Policy for a canonical command path (undefined: unmapped, so not remote). */
export function commandPolicy(path: string): CommandPolicy | undefined {
  return Object.hasOwn(COMMAND_POLICY, path) ? COMMAND_POLICY[path] : undefined;
}

/** The policy when `path` may run remotely at all. */
export function remotePolicy(path: string): RemoteCommand | undefined {
  const policy = commandPolicy(path);
  return policy?.remote ? policy : undefined;
}
