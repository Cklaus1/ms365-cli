import { GraphClient } from './graph-client.js';
import { CliError, ExitCode } from './errors.js';

/** Resource types that accept a name wherever an ID is expected (spec §6 "Call: names, not IDs"). */
export const RESOLVE_TYPES = ['todo-list', 'team', 'channel', 'plan', 'bucket'] as const;
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
  nameField: string;
  /** Parent option name, for nested types. */
  parent?: string;
}

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
/** Planner plan/bucket IDs are 28-char base64url strings. */
const PLANNER_ID = /^[A-Za-z0-9_-]{28}$/;

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
};

/** Whether `value` has the shape of a `type` ID (and so resolves without a request). */
export function looksLikeId(type: ResolveType, value: string): boolean {
  return SPECS[type].isId(value);
}

/** Per-process cache, keyed by client so independent clients never share answers. */
const cache = new WeakMap<GraphClient, Map<string, Resolved>>();

type Item = Record<string, unknown>;

/** All candidates of a collection, or null when dry-run returned no listing. */
async function listCandidates(client: GraphClient, endpoint: string): Promise<Item[] | null> {
  const items: Item[] = [];
  let next: string | undefined = endpoint;
  while (next) {
    const response = await client.request(next);
    const data = response.data as { value?: unknown } | undefined;
    if (!Array.isArray(data?.value)) return items.length ? items : null;
    items.push(...(data.value as Item[]));
    next = response.nextLink;
  }
  return items;
}

function pick(items: Item[], name: string, field: string): Item[] {
  const needle = name.toLowerCase();
  const named = items.map(item => ({ item, name: String(item[field] ?? '').toLowerCase() }));
  for (const test of [(n: string) => n === needle, (n: string) => n.startsWith(needle), (n: string) => n.includes(needle)]) {
    const hits = named.filter(c => test(c.name)).map(c => c.item);
    if (hits.length) return hits;
  }
  return [];
}

/**
 * Turn an ID or a name into an ID. Values shaped like the type's ID pass through with no
 * request; names are matched against the listed candidates (exact, then prefix, then substring,
 * case-insensitive). Nested types (channel, bucket) need the parent's ID.
 */
export async function resolveId(client: GraphClient, type: ResolveType, value: string, parent?: string): Promise<Resolved> {
  const spec = SPECS[type];
  if (looksLikeId(type, value)) return { type, id: value, displayName: null };
  if (spec.parent && !parent) {
    throw new CliError(`Resolving a ${spec.label} by name needs its ${spec.parent}; pass --${spec.parent} <ID or name>.`,
      { code: 'usage', exitCode: ExitCode.USAGE });
  }

  const key = `${type}\0${parent ?? ''}\0${value}`;
  let perClient = cache.get(client);
  const hit = perClient?.get(key);
  if (hit) return hit;

  const items = await listCandidates(client, spec.endpoint(parent));
  if (items === null) return { type, id: value, displayName: null };

  const matches = pick(items, value, spec.nameField);
  if (matches.length === 0) {
    const names = items.map(i => String(i[spec.nameField] ?? ''));
    const shown = names.slice(0, 10).join(', ') || '(none)';
    const more = names.length > 10 ? ` (and ${names.length - 10} more)` : '';
    throw new CliError(`No ${spec.label} named "${value}". Available: ${shown}${more}`,
      { code: 'not_found', exitCode: ExitCode.NOT_FOUND });
  }
  if (matches.length > 1) {
    const list = matches.map(i => `${String(i[spec.nameField] ?? '')} (${String(i.id)})`).join(', ');
    throw new CliError(`"${value}" is ambiguous: it matches ${matches.length} ${spec.label}s: ${list}`,
      { code: 'ambiguous_name', exitCode: ExitCode.USAGE, fix: 'Pass the exact name or the ID.' });
  }

  const resolved: Resolved = { type, id: String(matches[0].id), displayName: String(matches[0][spec.nameField] ?? '') };
  if (!perClient) cache.set(client, perClient = new Map());
  perClient.set(key, resolved);
  return resolved;
}

/** resolveId for callers that only need the ID. */
export async function resolve(client: GraphClient, type: ResolveType, value: string, parent?: string): Promise<string> {
  return (await resolveId(client, type, value, parent)).id;
}
