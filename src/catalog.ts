/**
 * Static metadata for every leaf command, served by `ms365 describe` (spec §6 "Discover").
 * test/describe.test.ts fails when a registered command has no entry here, or an entry has no command.
 *
 * Scopes are the Graph delegated permissions a command needs. A command with no scopes makes
 * no Graph call, so it is offline.
 */
export interface CommandMeta {
  scopes: string[];
  /** Deletes data (asks for confirmation where the source calls confirm()). */
  destructive: boolean;
  /** Repeating the command with the same input leaves the same end state. */
  idempotent: boolean;
  /** Never calls Graph. */
  offline: boolean;
  /** Only reads (catalog kind read). */
  readOnly: boolean;
}

/**
 * read   — only GETs. Group read scopes.
 * update — sets state in place (PATCH/PUT, or a POST that converges). Group write scopes.
 * action — creates, sends or moves something; repeating it does it again. Group write scopes.
 * delete — removes something. Group write scopes.
 */
type Kind = 'read' | 'update' | 'action' | 'delete';
/** A kind using the group's default scopes, or a kind with its own scopes. */
type Entry = Kind | [Kind, string[]];

const GROUP_SCOPES: Record<string, { read: string[]; write: string[] }> = {
  auth: { read: [], write: [] },
  mcp: { read: [], write: [] },
  resolve: { read: ['Tasks.Read', 'Team.ReadBasic.All', 'Channel.ReadBasic.All'], write: [] },
  mail: { read: ['Mail.Read'], write: ['Mail.ReadWrite'] },
  calendar: { read: ['Calendars.Read'], write: ['Calendars.ReadWrite'] },
  drive: { read: ['Files.Read'], write: ['Files.ReadWrite'] },
  todo: { read: ['Tasks.Read'], write: ['Tasks.ReadWrite'] },
  contacts: { read: ['Contacts.Read'], write: ['Contacts.ReadWrite'] },
  teams: { read: ['Team.ReadBasic.All'], write: ['ChannelMessage.Send'] },
  chat: { read: ['Chat.Read'], write: ['ChatMessage.Send'] },
  sharepoint: { read: ['Sites.Read.All'], write: [] },
  planner: { read: ['Tasks.Read'], write: ['Tasks.ReadWrite'] },
  onenote: { read: ['Notes.Read'], write: ['Notes.Create'] },
  excel: { read: ['Files.Read'], write: ['Files.ReadWrite'] },
  search: { read: [], write: [] },
  user: { read: ['User.Read'], write: ['MailboxSettings.ReadWrite'] },
  workflow: { read: [], write: [] },
  groups: { read: ['Group.Read.All'], write: [] },
  schema: { read: [], write: [] },
  'generate-skills': { read: [], write: [] },
  describe: { read: [], write: [] },
  meetings: { read: ['OnlineMeetings.Read'], write: [] },
  watch: { read: [], write: [] },
  audit: { read: [], write: [] },
};

const MAILBOX_READ = ['MailboxSettings.Read'];
const MAILBOX_WRITE = ['MailboxSettings.ReadWrite'];

const COMMANDS: Record<string, Entry> = {
  // auth: local token-cache management; login/status/setup verify the login against /me.
  'auth login': ['action', ['User.Read']],
  'auth logout': 'delete',
  'auth status': ['read', ['User.Read']],
  'auth scopes': 'read',
  'auth accounts': 'read',
  'auth select': 'update',
  'auth remove': 'delete',
  'auth setup': ['action', ['User.Read']],
  'auth profile list': 'read',
  'auth profile show': 'read',
  'auth profile set': 'update',
  'auth profile delete': 'delete',

  'mail list': 'read',
  'mail get': 'read',
  'mail read': 'read',
  'mail triage': 'read',
  'mail send': ['action', ['Mail.Send']],
  'mail delete': 'delete',
  'mail move': 'action',
  'mail folders': 'read',
  'mail folder': 'read',
  'mail create-folder': 'action',
  'mail draft': 'action',
  'mail drafts': 'read',
  'mail update-draft': 'update',
  'mail send-draft': ['action', ['Mail.Send']],
  'mail attachments': 'read',
  'mail attachment': 'read',
  'mail add-attachment': 'action',
  'mail delete-attachment': 'delete',
  'mail download-attachment': 'read',
  'mail reply': ['action', ['Mail.Send']],
  'mail forward': ['action', ['Mail.Send']],
  'mail mark-read': 'update',
  'mail mark-unread': 'update',
  'mail flag': 'update',
  'mail unflag': 'update',
  'mail complete': 'update',
  'mail categorize': 'update',
  'mail shared-list': ['read', ['Mail.Read.Shared']],
  'mail shared-get': ['read', ['Mail.Read.Shared']],
  'mail shared-send': ['action', ['Mail.Send.Shared']],
  'mail rules': ['read', MAILBOX_READ],
  'mail rule': ['read', MAILBOX_READ],
  'mail create-rule': ['action', MAILBOX_WRITE],
  'mail update-rule': ['update', MAILBOX_WRITE],
  'mail batch-read': 'update',
  'mail batch-move': 'action',

  'calendar list': 'read',
  'calendar events': 'read',
  'calendar get': 'read',
  'calendar create': 'action',
  'calendar update': 'update',
  'calendar delete': 'delete',
  'calendar view': 'read',
  'calendar accept': 'action',
  'calendar decline': 'action',
  'calendar tentative': 'action',
  'calendar free-busy': 'read',
  'calendar find-time': ['read', ['Calendars.Read.Shared']],
  'calendar rooms': ['read', ['Place.Read.All']],
  'calendar room-lists': ['read', ['Place.Read.All']],
  'calendar create-calendar': 'action',
  'calendar update-calendar': 'update',

  'drive list': 'read',
  'drive root': 'read',
  'drive files': 'read',
  'drive download': 'read',
  'drive upload': 'update',
  'drive delete': 'delete',
  'drive download-shared': ['read', ['Files.Read.All']],
  'drive search': 'read',
  'drive mkdir': 'action',
  // createLink returns the existing link when one of the same type exists.
  'drive share': 'update',
  'drive permissions': 'read',
  'drive quota': 'read',
  'drive info': 'read',
  'drive ls': 'read',
  'drive cat': 'read',
  'drive cp': 'update',

  'todo lists': 'read',
  'todo tasks': 'read',
  'todo get': 'read',
  'todo create': 'action',
  'todo update': 'update',
  'todo delete': 'delete',
  'todo create-list': 'action',
  'todo update-list': 'update',

  'contacts list': 'read',
  'contacts get': 'read',
  'contacts create': 'action',
  'contacts update': 'update',
  'contacts delete': 'delete',
  'contacts folders': 'read',
  'contacts create-folder': 'action',
  'contacts update-folder': 'update',
  'contacts folder-contacts': 'read',

  'teams list': 'read',
  'teams get': 'read',
  'teams members': ['read', ['TeamMember.Read.All']],
  'teams channels': ['read', ['Channel.ReadBasic.All']],
  'teams channel': ['read', ['Channel.ReadBasic.All']],
  'teams channel-messages': ['read', ['ChannelMessage.Read.All']],
  'teams send-channel-message': 'action',
  'teams reply-channel-message': 'action',

  'chat list': 'read',
  'chat get': 'read',
  'chat messages': ['read', ['ChatMessage.Read']],
  'chat send': 'action',
  'chat reply': ['action', ['ChatMessage.Read', 'ChatMessage.Send']],

  'sharepoint search': 'read',
  'sharepoint site': 'read',
  'sharepoint site-by-path': 'read',
  'sharepoint drives': 'read',
  'sharepoint drive': 'read',
  'sharepoint lists': 'read',
  'sharepoint list': 'read',
  'sharepoint list-items': 'read',
  'sharepoint list-item': 'read',
  'sharepoint items': 'read',
  'sharepoint item': 'read',
  'sharepoint delta': 'read',

  'planner plans': 'read',
  'planner plan': 'read',
  'planner plan-details': 'read',
  'planner my-tasks': 'read',
  'planner tasks': 'read',
  'planner task': 'read',
  'planner task-details': 'read',
  'planner buckets': 'read',
  'planner bucket': 'read',
  'planner bucket-tasks': 'read',
  'planner create-task': 'action',
  'planner update-task': 'update',
  'planner update-plan-details': 'update',
  'planner update-task-details': 'update',
  'planner delete-task': 'delete',
  'planner delete-plan': 'delete',
  'planner delete-bucket': 'delete',
  'planner update-plan': 'update',
  'planner update-bucket': 'update',
  'planner create-plan': 'action',
  'planner create-bucket': 'action',

  'onenote notebooks': 'read',
  'onenote sections': 'read',
  'onenote pages': 'read',
  'onenote page': 'read',
  'onenote create-page': 'action',

  'excel worksheets': 'read',
  'excel range': 'read',
  'excel create-chart': 'action',
  'excel format-range': 'update',
  'excel sort-range': 'update',

  // Which scope applies depends on --entity (driveItem, listItem/site, message, event, chatMessage).
  'search query': ['read', ['Files.Read.All', 'Sites.Read.All', 'Mail.Read', 'Calendars.Read', 'Chat.Read']],
  'search people': ['read', ['People.Read']],

  'user me': 'read',
  'user list': ['read', ['User.Read.All']],
  'user ooo': ['read', MAILBOX_READ],
  'user set-ooo': 'update',
  'user presence': ['read', ['Presence.Read']],
  'user set-presence': ['update', ['Presence.ReadWrite']],
  'user clear-presence': ['update', ['Presence.ReadWrite']],
  'user photo': 'read',
  'user download-photo': 'read',
  'user upload-photo': ['update', ['User.ReadWrite']],
  'user mailbox-settings': ['read', MAILBOX_READ],
  'user set-timezone': 'update',
  'user set-language': 'update',
  'user working-hours': ['read', MAILBOX_READ],

  'workflow standup': ['read', ['Calendars.Read', 'Mail.Read']],
  'workflow meeting-prep': ['read', ['Calendars.Read']],
  'workflow email-to-task': ['action', ['Mail.Read', 'Tasks.ReadWrite']],
  'workflow digest': ['read', ['Calendars.Read', 'Mail.Read', 'Tasks.Read']],
  'workflow file-announce': ['action', ['Files.Read', 'ChannelMessage.Send']],
  'workflow focus-time': ['read', ['Calendars.Read']],

  'groups list': 'read',
  'groups my': 'read',
  'groups get': 'read',
  'groups members': 'read',
  'groups owners': 'read',
  'groups distribution-lists': 'read',
  'groups teams': 'read',

  'meetings list': ['read', ['Calendars.Read']],
  'meetings get': 'read',
  'meetings transcripts': ['read', ['OnlineMeetings.Read', 'OnlineMeetingTranscript.Read.All']],
  'meetings transcript': ['read', ['OnlineMeetings.Read', 'OnlineMeetingTranscript.Read.All']],
  'meetings recordings': ['read', ['OnlineMeetings.Read', 'OnlineMeetingRecording.Read.All']],
  'meetings recording': ['read', ['OnlineMeetings.Read', 'OnlineMeetingRecording.Read.All']],
  'meetings all-transcripts': ['read', ['User.Read', 'OnlineMeetingTranscript.Read.All']],

  // Delta feeds: mail, events, files or contacts; the scope follows the <resource> argument.
  watch: ['read', ['Mail.Read', 'Calendars.Read', 'Files.Read', 'Contacts.Read']],

  'audit tail': 'read',
  'mcp serve': 'read',
  'resolve': 'read',
  'audit path': 'read',

  schema: 'read',
  'generate-skills': 'read',
  describe: 'read',
};

/** Metadata for a command path such as "mail list", or undefined when the catalog has no entry. */
export function commandMeta(path: string): CommandMeta | undefined {
  const entry = COMMANDS[path];
  if (!entry) return undefined;
  const [kind, own] = typeof entry === 'string' ? [entry, undefined] : entry;
  const group = GROUP_SCOPES[path.split(' ')[0]];
  const scopes = own ?? (kind === 'read' ? group.read : group.write);
  return {
    scopes: [...scopes],
    destructive: kind === 'delete',
    idempotent: kind !== 'action',
    offline: scopes.length === 0,
    readOnly: kind === 'read',
  };
}

/** Every command path the catalog describes. */
export function catalogPaths(): string[] {
  return Object.keys(COMMANDS);
}
