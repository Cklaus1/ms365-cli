/**
 * AH-3: Outlook ids that survive folder moves, and source references a host can store.
 *
 * Every mail and calendar request asks for immutable ids (`Prefer: IdType="ImmutableId"`, OQ-6). With `--refs`
 * (or MS365_CLI_REFS=1) every message and event a read returns carries `_ref`: who it was read as, which item and
 * which version, so a host can map it onto its `source-reference` record (README "Agent hosts").
 */

/** Who the token belongs to. `tenantId` is null when the token is opaque (the account then comes from GET /me). */
export interface RefIdentity {
  tenantId: string | null;
  accountId: string | null;
  /** The signed-in user's UPN (the mailbox read, unless the path names another user). */
  upn: string | null;
}

export type RefResource = 'message' | 'event';

export interface SourceRef {
  provider: 'microsoft-graph';
  tenantId: string | null;
  accountId: string | null;
  mailbox: string | null;
  resource: RefResource;
  id: string | null;
  version: string | null;
  lastModified: string | null;
  webLink: string | null;
  observedAt: string;
}

/** First path segments (after the owner) that are Outlook mail or calendar resources. */
const OUTLOOK_SEGMENTS: Record<string, true> = {
  messages: true, mailFolders: true, events: true, calendar: true, calendars: true, calendarView: true, calendarGroups: true, sendMail: true,
};

/** Fields `_ref` reads; added to a caller's `$select` so a narrowed read still has them. */
const REF_FIELDS = ['id', 'changeKey', 'lastModifiedDateTime', 'webLink'];

/** The path segments under the API version, owner first (`me` or `users/<id>` or `groups/<id>`). */
function outlookPath(url: string): { owner: string; rest: string[] } | null {
  let segments: string[];
  try {
    segments = new URL(url, 'https://graph.microsoft.com').pathname.split('/').filter(Boolean);
  } catch {
    return null;
  }
  if (segments[0] === 'v1.0' || segments[0] === 'beta') segments = segments.slice(1);
  let owner: string;
  let rest: string[];
  if (segments[0] === 'me') { owner = 'me'; rest = segments.slice(1); }
  else if ((segments[0] === 'users' || segments[0] === 'groups') && segments.length > 2) { owner = decodeURIComponent(segments[1]); rest = segments.slice(2); }
  else return null;
  // Graph's deltaLinks use the key form `mailFolders('AAMk…')/messages/delta` (E10): match on the segment name.
  return OUTLOOK_SEGMENTS[(rest[0] ?? '').replace(/\(.*\)$/, '')] === true ? { owner, rest } : null;
}

/** True for a mail or calendar request (absolute URL or path under the API base). */
export function isOutlookRequest(url: string): boolean {
  return outlookPath(url) !== null;
}

/** `headers` plus `Prefer: IdType="ImmutableId"` (merged into an existing Prefer) for mail and calendar requests. */
export function withImmutableIds(url: string, headers: Record<string, string> | undefined): Record<string, string> | undefined {
  if (!isOutlookRequest(url)) return headers;
  const key = Object.keys(headers ?? {}).find(k => k.toLowerCase() === 'prefer');
  const prefer = key ? headers![key] : '';
  if (/IdType\s*=/i.test(prefer)) return headers;
  const out = { ...headers };
  if (key) delete out[key];
  out.Prefer = prefer ? `${prefer}, IdType="ImmutableId"` : 'IdType="ImmutableId"';
  return out;
}

/**
 * Which kind of item a mail/calendar read returns, or null: a message or event collection, one item, a delta page
 * or an event's instances. Attachments, folders and calendars are not referenced.
 */
export function refResource(url: string): RefResource | null {
  const p = outlookPath(url);
  if (!p) return null;
  const rest = p.rest[p.rest.length - 1] === 'delta' ? p.rest.slice(0, -1) : p.rest;
  const last = rest[rest.length - 1];
  const parent = rest[rest.length - 2];
  if (last === 'messages' || parent === 'messages') return 'message';
  if (last === 'events' || last === 'calendarView' || last === 'instances' || parent === 'events' || parent === 'calendarView' || parent === 'instances') return 'event';
  return null;
}

/** The mailbox a path reads: the named user, else the signed-in user's UPN. */
export function refMailbox(url: string, identity: RefIdentity): string | null {
  const p = outlookPath(url);
  return p && p.owner !== 'me' ? p.owner : identity.upn;
}

/** `$select` widened with the fields `_ref` needs; unchanged when the caller did not narrow it. */
export function selectForRefs(select: string | undefined): string | undefined {
  if (!select) return select;
  const fields = select.split(',').map(f => f.trim()).filter(Boolean);
  for (const f of REF_FIELDS) if (!fields.includes(f)) fields.push(f);
  return fields.join(',');
}

const str = (v: unknown): string | null => (typeof v === 'string' && v ? v : null);

/** The `_ref` of one Graph item. */
export function sourceRef(item: Record<string, unknown>, resource: RefResource, identity: RefIdentity, mailbox: string | null, observedAt: string): SourceRef {
  return {
    provider: 'microsoft-graph',
    tenantId: identity.tenantId,
    accountId: identity.accountId,
    mailbox,
    resource,
    id: str(item.id),
    version: str(item.changeKey) ?? str(item['@odata.etag']),
    lastModified: str(item.lastModifiedDateTime),
    webLink: str(item.webLink),
    observedAt,
  };
}

/** Add `_ref` to every item of a Graph response body (`{value: [...]}` or one object), in place. */
export function attachRefs(data: unknown, resource: RefResource, identity: RefIdentity, mailbox: string | null, observedAt: string): void {
  if (!data || typeof data !== 'object') return;
  const value: unknown = Reflect.get(data, 'value');
  const items: unknown[] = Array.isArray(value) ? value : [data];
  for (const item of items) {
    if (item && typeof item === 'object' && !Array.isArray(item)) {
      Reflect.set(item, '_ref', sourceRef({ ...item }, resource, identity, mailbox, observedAt));
    }
  }
}
