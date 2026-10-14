import { GraphClient } from './graph-client.js';
import { CliError, ExitCode } from './errors.js';

/** Resource types that accept a name wherever an ID is expected (spec §6 "Call: names, not IDs"). */
export const RESOLVE_TYPES = ['todo-list', 'team', 'channel', 'plan', 'bucket', 'mail-folder'] as const;
export type ResolveType = typeof RESOLVE_TYPES[number];

export interface Resolved {
  type: ResolveType;
  id: string;
  displayName: string | null;
}

interface TypeSpec {
  label: string;
  /** Whether the value already has the shape of this type's Graph ID. */
  isId: (value: string) => boolean;
  /** Collection to list candidates from; `parent` is the team/plan ID for nested types. */
  endpoint: (parent?: string) => string;
  /** Query parameters for the first listing request (nextLinks carry their own). */
  query?: Record<string, string>;
  nameField: string;
  /** Parent option name, for nested types. */
  parent?: string;
  /** The parent narrows the lookup but isn't required (mail-folder: another user's mailbox). */
  parentOptional?: boolean;
  /** 'exact' allows only exact, case-insensitive name matches (no prefix/substring). */
  match?: 'exact';
}

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
/** Planner plan/bucket IDs are 28-char base64url strings. */
const PLANNER_ID = /^[A-Za-z0-9_-]{28}$/;

/** Folder names Graph accepts in place of a mail folder ID. */
export const WELL_KNOWN_FOLDERS = [
  'inbox', 'drafts', 'sentitems', 'deleteditems', 'archive',
  'junkemail', 'outbox', 'clutter', 'conflicts',
  'conversationhistory', 'localfailures', 'msgfolderroot',
  'recoverableitemsdeletions', 'scheduled', 'searchfolders', 'serverfailures',
];

const FOLDER_QUERY = { '$select': 'id,displayName,childFolderCount', '$top': '100' };
/** How many folders the breadth-first fallback lists before giving up. */
const FOLDER_SEARCH_CAP = 500;

/** `/me`, or `/users/{id}` for another mailbox. */
function mailbox(user?: string): string {
  return user ? `/users/${encodeURIComponent(user)}` : '/me';
}

const SPECS: Record<ResolveType, TypeSpec> = {
  'todo-list': {
    label: 'To Do list',
    // Exchange-backed IDs: 'AAMk…' base64, or any long space-free base64 string
    isId: v => /^AAMk[A-Za-z0-9+/=_-]+$/.test(v) || /^[A-Za-z0-9+/=_-]{40,}$/.test(v),
    endpoint: () => '/me/todo/lists',
    nameField: 'displayName',
  },
  team: {
    label: 'team',
    isId: v => GUID.test(v),
    endpoint: () => '/me/joinedTeams',
    nameField: 'displayName',
  },
  channel: {
    label: 'channel',
    isId: v => /^19:\S+@thread\.[a-z0-9]+$/i.test(v),
    endpoint: team => `/teams/${encodeURIComponent(team!)}/channels`,
    nameField: 'displayName',
    parent: 'team',
  },
  plan: {
    label: 'plan',
    isId: v => PLANNER_ID.test(v),
    endpoint: () => '/me/planner/plans',
    nameField: 'title',
  },
  bucket: {
    label: 'bucket',
    isId: v => PLANNER_ID.test(v),
    endpoint: plan => `/planner/plans/${encodeURIComponent(plan!)}/buckets`,
    nameField: 'name',
    parent: 'plan',
  },
  'mail-folder': {
    label: 'mail folder',
    // Exchange-backed IDs: 'AAMk…'/'AQMk…' base64, or a long base64 string with no '/' (a '/' means a path)
    isId: v => /^A[AQ]Mk[A-Za-z0-9+/=_-]+$/.test(v) || /^[A-Za-z0-9+=_-]{40,}$/.test(v),
    endpoint: user => `${mailbox(user)}/mailFolders`,
    query: FOLDER_QUERY,
    nameField: 'displayName',
    parent: 'user',
    parentOptional: true,
    match: 'exact',
  },
};

/** Whether `value` has the shape of a `type` ID (and so resolves without a request). */
export function looksLikeId(type: ResolveType, value: string): boolean {
  return SPECS[type].isId(value);
}

/** Per-process cache, keyed by client so independent clients never share answers. */
const cache = new WeakMap<GraphClient, Map<string, Resolved>>();

type Item = Record<string, unknown>;

/** All candidates of a collection, or null when dry-run returned no listing. */
async function listCandidates(client: GraphClient, endpoint: string, query?: Record<string, string>): Promise<Item[] | null> {
  const items: Item[] = [];
  let next: string | undefined = endpoint;
  while (next) {
    const response = await client.request(next, next === endpoint ? { queryParams: query } : {});
    const data = response.data as { value?: unknown } | undefined;
    if (!Array.isArray(data?.value)) return items.length ? items : null;
    items.push(...(data.value as Item[]));
    next = response.nextLink;
  }
  return items;
}

function pick(items: Item[], name: string, field: string, match?: 'exact'): Item[] {
  const needle = name.toLowerCase();
  const named = items.map(item => ({ item, name: String(item[field] ?? '').toLowerCase() }));
  const tests = [(n: string) => n === needle, (n: string) => n.startsWith(needle), (n: string) => n.includes(needle)];
  for (const test of match === 'exact' ? tests.slice(0, 1) : tests) {
    const hits = named.filter(c => test(c.name)).map(c => c.item);
    if (hits.length) return hits;
  }
  return [];
}

function notFound(label: string, what: string, names: string[], extra = '', fix?: string): CliError {
  const shown = names.slice(0, 10).join(', ') || '(none)';
  const more = names.length > 10 ? ` (and ${names.length - 10} more)` : '';
  return new CliError(`No ${label} named ${what}${extra}. Available: ${shown}${more}`,
    { code: 'not_found', exitCode: ExitCode.NOT_FOUND, fix });
}

function ambiguous(value: string, label: string, list: string[]): CliError {
  return new CliError(`"${value}" is ambiguous: it matches ${list.length} ${label}s: ${list.join(', ')}`,
    { code: 'ambiguous_name', exitCode: ExitCode.USAGE, fix: 'Pass the exact name or the ID.' });
}

/** Match `value` against one listing; null when dry-run returned no listing. */
async function findByName(client: GraphClient, type: ResolveType, endpoint: string, value: string): Promise<Resolved | null> {
  const spec = SPECS[type];
  const items = await listCandidates(client, endpoint, spec.query);
  if (items === null) return null;
  const matches = pick(items, value, spec.nameField, spec.match);
  if (matches.length === 0) throw notFound(spec.label, `"${value}"`, items.map(i => String(i[spec.nameField] ?? '')));
  if (matches.length > 1) {
    throw ambiguous(value, spec.label, matches.map(i => `${String(i[spec.nameField] ?? '')} (${String(i.id)})`));
  }
  return { type, id: String(matches[0].id), displayName: String(matches[0][spec.nameField] ?? '') };
}

/** The well-known folder name `value` stands for ("Sent Items" → sentitems), if any. */
function wellKnownFolder(value: string): string | undefined {
  const normalized = value.toLowerCase().replace(/\s+/g, '');
  return WELL_KNOWN_FOLDERS.find(f => f === normalized);
}

const PATH_FIX = 'Pass a path such as Inbox/Projects, or the folder ID.';

/**
 * Resolve a mail folder name (F11: exact, case-insensitive) or path (F12). `Inbox/Projects` walks
 * childFolders one segment at a time; the first segment may be a well-known name. A plain name
 * that misses at the top level is searched breadth-first, listing at most 500 folders; two
 * matches are ambiguous. Null when dry-run returned no listing.
 */
async function findMailFolder(client: GraphClient, value: string, user?: string): Promise<Resolved | null> {
  const root = SPECS['mail-folder'].endpoint(user);
  const children = (id: string) => `${root}/${encodeURIComponent(id)}/childFolders`;
  const segments = value.split('/').map(s => s.trim());
  if (segments.some(s => !s)) {
    throw new CliError(`Mail folder path "${value}" has an empty segment.`, { code: 'usage', exitCode: ExitCode.USAGE, fix: PATH_FIX });
  }

  if (segments.length > 1) {
    const first = wellKnownFolder(segments[0]);
    let current = first
      ? { type: 'mail-folder' as const, id: first, displayName: segments[0] }
      : await findByName(client, 'mail-folder', root, segments[0]);
    for (const segment of segments.slice(1)) {
      if (current === null) return null;
      const items = await listCandidates(client, children(current.id), FOLDER_QUERY);
      if (items === null) return null;
      const matches = pick(items, segment, 'displayName', 'exact');
      if (matches.length !== 1) {
        const names = items.map(i => String(i.displayName ?? ''));
        if (matches.length === 0) throw notFound('mail folder', `"${segment}" under "${current.displayName}"`, names, '', PATH_FIX);
        throw ambiguous(segment, 'mail folder', matches.map(i => `${String(i.displayName)} (${String(i.id)})`));
      }
      current = { type: 'mail-folder', id: String(matches[0].id), displayName: String(matches[0].displayName ?? '') };
    }
    return current;
  }

  const top = await listCandidates(client, root, FOLDER_QUERY);
  if (top === null) return null;
  const direct = pick(top, value, 'displayName', 'exact');
  if (direct.length === 1) return { type: 'mail-folder', id: String(direct[0].id), displayName: String(direct[0].displayName ?? '') };
  if (direct.length > 1) throw ambiguous(value, 'mail folder', direct.map(i => `${String(i.displayName)} (${String(i.id)})`));

  // Breadth-first over folders that have children; `path` is the display path for messages.
  type Node = { item: Item; path: string };
  const asNodes = (items: Item[], prefix: string): Node[] =>
    items.map(item => ({ item, path: prefix + String(item.displayName ?? '') }));
  const seen = asNodes(top.slice(0, FOLDER_SEARCH_CAP), '');
  let capped = top.length > FOLDER_SEARCH_CAP;
  for (let i = 0; i < seen.length && !capped; i++) {
    if (Number(seen[i].item.childFolderCount ?? 0) === 0) continue;
    const items = await listCandidates(client, children(String(seen[i].item.id)), FOLDER_QUERY) ?? [];
    const room = FOLDER_SEARCH_CAP - seen.length;
    if (items.length > room) capped = true;
    seen.push(...asNodes(items.slice(0, room), `${seen[i].path}/`));
  }
  const needle = value.toLowerCase();
  const hits = seen.filter(n => String(n.item.displayName ?? '').toLowerCase() === needle);
  if (hits.length === 1) {
    return { type: 'mail-folder', id: String(hits[0].item.id), displayName: String(hits[0].item.displayName ?? '') };
  }
  if (hits.length > 1) throw ambiguous(value, 'mail folder', hits.map(n => `${n.path} (${String(n.item.id)})`));
  const extra = capped ? ` (searched the first ${FOLDER_SEARCH_CAP} folders)` : '';
  throw notFound('mail folder', `"${value}"`, top.map(i => String(i.displayName ?? '')), extra, PATH_FIX);
}

/**
 * Turn an ID or a name into an ID. Values shaped like the type's ID pass through with no
 * request; names are matched against the listed candidates (exact, then prefix, then substring,
 * case-insensitive; exact only for mail-folder). Nested types (channel, bucket) need the parent's ID.
 * For mail-folder, `parent` is an optional user ID or UPN whose mailbox is searched (default /me);
 * see findMailFolder for well-known names, paths and the nested-folder search.
 */
export async function resolveId(client: GraphClient, type: ResolveType, value: string, parent?: string): Promise<Resolved> {
  const spec = SPECS[type];
  if (looksLikeId(type, value)) return { type, id: value, displayName: null };
  if (type === 'mail-folder') {
    const wellKnown = wellKnownFolder(value);
    if (wellKnown) return { type, id: wellKnown, displayName: null };
  }
  if (spec.parent && !spec.parentOptional && !parent) {
    throw new CliError(`Resolving a ${spec.label} by name needs its ${spec.parent}; pass --${spec.parent} <ID or name>.`,
      { code: 'usage', exitCode: ExitCode.USAGE });
  }

  const key = `${type}\0${parent ?? ''}\0${value}`;
  let perClient = cache.get(client);
  const hit = perClient?.get(key);
  if (hit) return hit;

  const resolved = type === 'mail-folder'
    ? await findMailFolder(client, value, parent)
    : await findByName(client, type, spec.endpoint(parent), value);
  if (resolved === null) return { type, id: value, displayName: null };

  if (!perClient) cache.set(client, perClient = new Map());
  perClient.set(key, resolved);
  return resolved;
}

/** resolveId for callers that only need the ID. */
export async function resolve(client: GraphClient, type: ResolveType, value: string, parent?: string): Promise<string> {
  return (await resolveId(client, type, value, parent)).id;
}
