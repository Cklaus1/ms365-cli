/**
 * Default `-o table` columns per Graph resource shape (spec PM-10). Wide Graph objects make a
 * table unreadable, so table output shows these columns unless --fields is set or
 * MS365_CLI_TABLE_ALL_COLUMNS=1. Columns are dot paths into the item. No shape has an id column:
 * table cells are cut to 50 characters, which breaks most Graph IDs (use --fields id,… or -o json; F9).
 */
export interface Shape {
  name: string;
  /** @odata.type values (without the "#microsoft.graph." prefix). */
  types: string[];
  /** Key heuristic for items without @odata.type (most list responses). */
  match: (has: (key: string) => boolean) => boolean;
  columns: string[];
}

// Order matters: the first shape whose heuristic matches wins.
export const SHAPES: Shape[] = [
  {
    name: 'chatMessage', types: ['chatMessage'],
    match: has => has('messageType') && has('body'),
    columns: ['createdDateTime', 'from.user.displayName', 'body.content'],
  },
  {
    name: 'message', types: ['message', 'eventMessage', 'eventMessageRequest', 'eventMessageResponse'],
    match: has => has('receivedDateTime'),
    columns: ['receivedDateTime', 'from.emailAddress.address', 'subject', 'isRead', 'hasAttachments'],
  },
  {
    name: 'event', types: ['event'],
    match: has => has('start') && has('end') && has('subject'),
    columns: ['subject', 'start.dateTime', 'end.dateTime', 'location.displayName', 'organizer.emailAddress.name'],
  },
  {
    name: 'plannerTask', types: ['plannerTask'],
    match: has => has('planId') && has('bucketId'),
    columns: ['title', 'percentComplete', 'priority', 'dueDateTime'],
  },
  {
    name: 'todoTask', types: ['todoTask'],
    match: has => has('title') && has('status') && has('importance'),
    columns: ['title', 'status', 'dueDateTime.dateTime', 'importance'],
  },
  {
    name: 'todoTaskList', types: ['todoTaskList'],
    match: has => has('wellknownListName'),
    columns: ['displayName', 'wellknownListName', 'isOwner'],
  },
  {
    name: 'plannerPlan', types: ['plannerPlan'],
    match: has => has('title') && (has('owner') || has('container')),
    columns: ['title', 'container.type', 'createdDateTime'],
  },
  {
    name: 'notebook', types: ['notebook'],
    match: has => has('sectionsUrl') && has('isDefault'),
    columns: ['displayName', 'isDefault', 'isShared', 'lastModifiedDateTime'],
  },
  {
    name: 'onenoteSection', types: ['onenoteSection'],
    match: has => has('pagesUrl'),
    columns: ['displayName', 'isDefault', 'lastModifiedDateTime'],
  },
  {
    name: 'onenotePage', types: ['onenotePage'],
    match: has => has('contentUrl') && has('title'),
    columns: ['title', 'createdDateTime', 'lastModifiedDateTime'],
  },
  {
    name: 'driveItem', types: ['driveItem'],
    match: has => has('name') && (has('file') || has('folder') || has('parentReference')),
    columns: ['name', 'size', 'lastModifiedDateTime', 'webUrl'],
  },
  {
    name: 'group', types: ['group'],
    match: has => has('groupTypes'),
    columns: ['displayName', 'mail', 'groupTypes'],
  },
  {
    name: 'user', types: ['user'],
    match: has => has('userPrincipalName'),
    columns: ['displayName', 'mail', 'jobTitle'],
  },
  {
    name: 'chat', types: ['chat'],
    match: has => has('chatType'),
    // members comes from chat list's table-only $expand=members (X-13): topic is null for 1:1 chats.
    columns: ['topic', 'members', 'chatType', 'lastUpdatedDateTime'],
  },
  {
    name: 'channel', types: ['channel'],
    match: has => has('membershipType'),
    columns: ['displayName', 'membershipType', 'description'],
  },
  {
    name: 'team', types: ['team'],
    match: has => has('isArchived') && has('displayName'),
    columns: ['displayName', 'description', 'isArchived'],
  },
  {
    name: 'contact', types: ['contact'],
    match: has => has('emailAddresses') && has('displayName'),
    columns: ['displayName', 'emailAddresses', 'mobilePhone', 'companyName'],
  },
];

/** The shape of one Graph item, by @odata.type first, then by its keys. */
export function detectShape(item: unknown): Shape | undefined {
  if (item === null || typeof item !== 'object' || Array.isArray(item)) return undefined;
  const obj = item as Record<string, unknown>;
  const type = typeof obj['@odata.type'] === 'string' ? (obj['@odata.type'] as string).replace(/^#microsoft\.graph\./, '') : '';
  if (type) {
    const byType = SHAPES.find(s => s.types.includes(type));
    if (byType) return byType;
  }
  const has = (key: string) => key in obj;
  return SHAPES.find(s => s.match(has));
}

/** Default columns for a list of items (detected from the first object), or null to keep them all. */
export function defaultColumns(items: unknown[]): string[] | null {
  const first = items.find(i => i !== null && typeof i === 'object');
  return detectShape(first)?.columns ?? null;
}
