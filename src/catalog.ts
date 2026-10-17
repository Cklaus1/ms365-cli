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
  /** Scopes that depend on an option's value(s) or a positional argument; `scopes` is then their union. */
  scopesBy: ScopeSelector | null;
}

/**
 * Scopes chosen by what the command is asked to touch (Q-16 to Q-19). Each selected value is one any-of clause:
 * `option` reads the option's value (comma-separated when `split`); `arg` reads a positional argument by index.
 * `withOption` adds a second clause for the selected argument value when that option is given.
 */
export type ScopeSelector =
  | { option: string; split?: ','; map: Record<string, string[]> }
  | { arg: number; map: Record<string, string[]>; withOption?: { option: string; map: Record<string, string[]> } };

/** One any-of clause of a command's requirement; `for` names the option or argument value that selected it. */
export interface ScopeClause {
  scopes: string[];
  for?: string;
}

/**
 * read   — only GETs. Group read scopes.
 * update — sets state in place (PATCH/PUT, or a POST that converges). Group write scopes.
 * action — creates, sends or moves something; repeating it does it again. Group write scopes.
 * delete — removes something. Group write scopes.
 */
type Kind = 'read' | 'update' | 'action' | 'delete';
/** A kind using the group's default scopes, a kind with its own scopes, or scopes chosen by a selector. */
type Entry = Kind | [Kind, string[]] | [Kind, ScopeSelector];

const GROUP_SCOPES: Record<string, { read: string[]; write: string[] }> = {
  auth: { read: [], write: [] },
  mcp: { read: [], write: [] },
  completion: { read: [], write: [] },
  notes: { read: ['ShortNotes.Read'], write: ['ShortNotes.ReadWrite'] },
  resolve: { read: [], write: [] },
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
  reports: { read: ['Reports.Read.All'], write: [] },
  'audit-trail': { read: ['AuditLog.Read.All'], write: [] },
  watch: { read: [], write: [] },
  audit: { read: [], write: [] },
  idempotency: { read: [], write: [] },
};

const MAILBOX_READ = ['MailboxSettings.Read'];
const MAILBOX_WRITE = ['MailboxSettings.ReadWrite'];

const COMMANDS: Record<string, Entry> = {
  // auth: local token-cache management; login/status/setup verify the login against /me.
  'auth login': ['action', ['User.Read']],
  'auth logout': 'delete',
  'auth status': ['read', ['User.Read']],
  'auth scopes': 'read',
  'auth refresh': 'update',
  'auth grant-scopes': 'update',
  'auth accounts': 'read',
  'auth select': 'update',
  'auth remove': 'delete',
  'auth setup': ['action', ['User.Read']],
  'auth storage show': 'read',
  'auth storage use': 'update',
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
  'drive copy': 'action',

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
  'teams send-channel-message': ['action', ['ChannelMessage.Send', 'User.ReadBasic.All', 'Team.ReadBasic.All', 'Channel.ReadBasic.All']],
  'teams reply-channel-message': 'action',

  'chat list': 'read',
  'chat get': 'read',
  'chat messages': ['read', ['ChatMessage.Read']],
  'chat send': ['action', ['ChatMessage.Send', 'User.ReadBasic.All']],
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

  'search query': ['read', {
    option: '--entity', split: ',',
    map: {
      driveItem: ['Files.Read.All', 'Sites.Read.All'], listItem: ['Sites.Read.All'], site: ['Sites.Read.All'],
      message: ['Mail.Read'], event: ['Calendars.Read'], chatMessage: ['Chat.Read'],
    },
  }],
  'search people': ['read', ['People.Read']],

  'user me': 'read',
  // GET /users needs only User.ReadBasic.All for the default (basic) properties.
  'user list': ['read', ['User.ReadBasic.All', 'User.Read.All']],
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
  'reports email': 'read',
  'reports teams': 'read',
  'reports active-users': 'read',
  'reports by-department': ['read', ['Reports.Read.All', 'User.Read.All']],

  // audit-trail (AO-25): `mailbox` creates a saved audit query (a POST; no data is changed), so it is not `read`.
  'audit-trail mailbox': ['action', ['AuditLogsQuery-Exchange.Read.All']],
  'audit-trail changes': 'read',
  // apps: every GET; Directory.Read.All is the least-privileged scope for listing grants (AO-28, OQ-9/E4) and also reads service principals.
  'audit-trail apps': ['read', ['Directory.Read.All']],
  // access: same saved-query POST as mailbox (AO-25); the target annotation's User.Read.All is already in ORG_SCOPES (AO-30).
  'audit-trail access': ['action', ['AuditLogsQuery-Exchange.Read.All']],
  'meetings transcript': ['read', ['OnlineMeetings.Read', 'OnlineMeetingTranscript.Read.All']],
  'meetings recordings': ['read', ['OnlineMeetings.Read', 'OnlineMeetingRecording.Read.All']],
  'meetings recording': ['read', ['OnlineMeetings.Read', 'OnlineMeetingRecording.Read.All']],
  'meetings all-transcripts': ['read', ['User.Read', 'OnlineMeetingTranscript.Read.All']],

  watch: ['read', { arg: 0, map: { mail: ['Mail.Read'], events: ['Calendars.Read'], files: ['Files.Read'], contacts: ['Contacts.Read'] } }],

  'audit tail': 'read',
  'excel tables': 'read',
  'excel table-rows': 'read',
  'excel add-rows': 'action',
  'todo checklist': 'read',
  'todo check-add': 'action',
  'todo check-done': 'update',
  'todo check-delete': 'delete',
  'chat create': ['action', ['Chat.Create', 'User.Read']],
  'notes list': 'read',
  'notes get': 'read',
  'notes create': 'action',
  'notes update': 'update',
  'notes delete': 'delete',
  'mcp serve': 'read',
  'completion': 'read',
  resolve: ['read', {
    arg: 0,
    map: {
      'todo-list': ['Tasks.Read'], team: ['Team.ReadBasic.All'], channel: ['Channel.ReadBasic.All'],
      plan: ['Tasks.Read'], bucket: ['Tasks.Read'], 'mail-folder': ['Mail.Read'],
    },
    withOption: { option: '--user', map: { 'mail-folder': ['Mail.Read.Shared'] } },
  }],
  'audit path': 'read',
  // Local ledger only; release forgets a key so the next send with it runs again (R-6).
  'idempotency list': 'read',
  'idempotency release': 'action',

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
  const scopesBy = own && !Array.isArray(own) ? own : null;
  const scopes = scopesBy ? selectorUnion(scopesBy) : Array.isArray(own) ? own : kind === 'read' ? group.read : group.write;
  return {
    scopes: [...scopes],
    destructive: kind === 'delete',
    idempotent: kind !== 'action',
    offline: scopes.length === 0,
    readOnly: kind === 'read',
    scopesBy,
  };
}

function selectorUnion(sel: ScopeSelector): string[] {
  const maps = [sel.map, ...('withOption' in sel && sel.withOption ? [sel.withOption.map] : [])];
  return [...new Set(maps.flatMap(m => Object.values(m).flat()))];
}

/** `--entity` → the commander opts key `entity`. */
const optKey = (flag: string) => flag.replace(/^--/, '').replace(/-(\w)/g, (_, c: string) => c.toUpperCase());

/**
 * The scopes a command run needs, in conjunctive normal form (F14): every clause must be satisfied, any scope
 * within a clause satisfies it. A selector yields one clause per selected value; values it does not know are
 * skipped (the command's own validation rejects them). Without a selector, or when nothing was selected, the
 * catalog scopes form one clause. Offline commands need nothing.
 */
export function requiredScopes(path: string, opts: Record<string, unknown> = {}, args: unknown[] = []): ScopeClause[] {
  const meta = commandMeta(path);
  if (!meta || meta.scopes.length === 0) return [];
  const sel = meta.scopesBy;
  if (!sel) return [{ scopes: meta.scopes }];
  const clauses: ScopeClause[] = [];
  if ('option' in sel) {
    const raw = opts[optKey(sel.option)];
    const values = typeof raw === 'string' ? (sel.split ? raw.split(sel.split).map(v => v.trim()) : [raw]) : [];
    for (const v of values) if (sel.map[v]) clauses.push({ scopes: sel.map[v], for: `${sel.option} ${v}` });
  } else {
    const v = String(args[sel.arg] ?? '');
    if (sel.map[v]) clauses.push({ scopes: sel.map[v], for: v });
    const extra = sel.withOption;
    if (extra && opts[optKey(extra.option)] !== undefined && extra.map[v]) {
      clauses.push({ scopes: extra.map[v], for: `${v} with ${extra.option}` });
    }
  }
  return clauses.length ? clauses : [{ scopes: meta.scopes }];
}

/** Every command path the catalog describes. */
export function catalogPaths(): string[] {
  return Object.keys(COMMANDS);
}
