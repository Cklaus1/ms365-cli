/**
 * Small, pure builders for Graph request pieces shared by command modules:
 * OData literals, `$search` phrases, drive colon paths, and the outlook time-zone Prefer header.
 */

/** An OData string literal: `O'Brien` -> `'O''Brien'` (quotes included). Use inside `$filter`. */
export function odataQuote(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

/** A `$search` phrase: wrapped in `"`, with `\` and `"` backslash-escaped (`say "hi"` -> `"say \"hi\""`). */
export function searchQuote(value: string): string {
  return `"${value.replace(/[\\"]/g, (c) => `\\${c}`)}"`;
}

/**
 * A user drive path as encoded colon-path segments: `encodeURIComponent` per segment so `#`, `?`, `%`
 * reach Graph literally; empty segments (leading/trailing/double `/`) are dropped; root -> `''`.
 */
export function encodeDrivePath(userPath: string): string {
  return userPath.split('/').filter((s) => s !== '').map(encodeURIComponent).join('/');
}

/**
 * Address a drive item by path under `root` (e.g. `/me/drive/root`, `/drives/{id}/root`,
 * `/drives/{id}/items/{parentId}`), optionally followed by a sub-resource `suffix` (`/content`, `/children`):
 * `drivePath('/me/drive/root', '/Docs/a#1.md', '/content')` -> `/me/drive/root:/Docs/a%231.md:/content`.
 * The root itself (`/`, `''`) is addressed without a colon path: `/me/drive/root/children`.
 * Without `suffix` the result has no trailing colon (`root:/Docs/a.md`), the form `uploadDriveItem` takes.
 */
export function drivePath(root: string, userPath: string, suffix = ''): string {
  const encoded = encodeDrivePath(userPath);
  if (encoded === '') return `${root}${suffix}`;
  return suffix ? `${root}:/${encoded}:${suffix}` : `${root}:/${encoded}`;
}

/**
 * Whether `value` starts with a `YYYY-MM-DD` that names a real calendar day. `Date` silently rolls
 * impossible days over (`2026-02-30` -> March 2), so the parts are round-tripped through `Date.UTC`.
 */
export function isCalendarDate(value: string): boolean {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(value);
  if (!m) return false;
  const [year, month, day] = [Number(m[1]), Number(m[2]), Number(m[3])];
  const d = new Date(Date.UTC(year, month - 1, day));
  return d.getUTCFullYear() === year && d.getUTCMonth() === month - 1 && d.getUTCDate() === day;
}

/** Validate and return a `YYYY-MM-DD` date that exists on the calendar (`2026-02-30` is rejected). */
export function calendarDate(input: string): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(input) || !isCalendarDate(input)) {
    throw new Error(`Invalid date: "${input}". Use YYYY-MM-DD format.`);
  }
  return input;
}

/**
 * The local IANA time zone (e.g. `Europe/Berlin`), or `UTC` when the runtime cannot name one: a POSIX
 * `TZ` such as `JST-9` resolves to `undefined`, and `TZ=` to `Etc/Unknown`, neither of which Graph or
 * `Intl` accepts.
 */
export function localTimeZone(): string {
  const zone = Intl.DateTimeFormat().resolvedOptions().timeZone;
  if (typeof zone !== 'string' || zone === '' || zone === 'Etc/Unknown') return 'UTC';
  try {
    new Intl.DateTimeFormat('en-US', { timeZone: zone });
    return zone;
  } catch {
    return 'UTC';
  }
}

/**
 * The IANA names Graph documents for `dateTimeTimeZone.timeZone` besides Windows names ("Additional
 * time zones", https://learn.microsoft.com/en-us/graph/api/resources/datetimetimezone). Entries
 * written `Doc (Alias)` in the doc use a name the runtime may not know; the alias is the IANA city.
 */
const GRAPH_TIME_ZONES = [
  'Etc/GMT+12', 'Etc/GMT+11', 'Pacific/Honolulu', 'America/Anchorage',
  'America/Santa_Isabel', 'America/Los_Angeles', 'America/Phoenix', 'America/Chihuahua',
  'America/Denver', 'America/Guatemala', 'America/Chicago', 'America/Mexico_City',
  'America/Regina', 'America/Bogota', 'America/New_York', 'America/Indiana/Indianapolis',
  'America/Caracas', 'America/Asuncion', 'America/Halifax', 'America/Cuiaba',
  'America/La_Paz', 'America/Santiago', 'America/St_Johns', 'America/Sao_Paulo',
  'America/Argentina/Buenos_Aires', 'America/Cayenne', 'America/Godthab', 'America/Montevideo',
  'America/Bahia', 'Etc/GMT+2', 'Atlantic/Azores', 'Atlantic/Cape_Verde',
  'Africa/Casablanca', 'Etc/GMT', 'Europe/London', 'Atlantic/Reykjavik',
  'Europe/Berlin', 'Europe/Budapest', 'Europe/Paris', 'Europe/Warsaw',
  'Africa/Lagos', 'Africa/Windhoek', 'Europe/Bucharest', 'Asia/Beirut',
  'Africa/Cairo', 'Asia/Damascus', 'Africa/Johannesburg', 'Europe/Kyiv',
  'Europe/Istanbul', 'Asia/Jerusalem', 'Asia/Amman', 'Asia/Baghdad',
  'Europe/Kaliningrad', 'Asia/Riyadh', 'Africa/Nairobi', 'Asia/Tehran',
  'Asia/Dubai', 'Asia/Baku', 'Europe/Moscow', 'Indian/Mauritius',
  'Asia/Tbilisi', 'Asia/Yerevan', 'Asia/Kabul', 'Asia/Karachi',
  'Asia/Toshkent (Tashkent)', 'Asia/Kolkata', 'Asia/Colombo', 'Asia/Kathmandu',
  'Asia/Astana (Almaty)', 'Asia/Dhaka', 'Asia/Yekaterinburg', 'Asia/Yangon (Rangoon)',
  'Asia/Bangkok', 'Asia/Novosibirsk', 'Asia/Shanghai', 'Asia/Krasnoyarsk',
  'Asia/Singapore', 'Australia/Perth', 'Asia/Taipei', 'Asia/Ulaanbaatar',
  'Asia/Irkutsk', 'Asia/Tokyo', 'Asia/Seoul', 'Australia/Adelaide',
  'Australia/Darwin', 'Australia/Brisbane', 'Australia/Sydney', 'Pacific/Port_Moresby',
  'Australia/Hobart', 'Asia/Yakutsk', 'Pacific/Guadalcanal', 'Asia/Vladivostok',
  'Pacific/Auckland', 'Etc/GMT-12', 'Pacific/Fiji', 'Asia/Magadan',
  'Pacific/Tongatapu', 'Pacific/Apia', 'Pacific/Kiritimati',
];

/** The name `Intl` reports for `zone` (ICU keeps legacy names: `Asia/Kolkata` -> `Asia/Calcutta`), or undefined. */
function canonicalZone(zone: string): string | undefined {
  try {
    return new Intl.DateTimeFormat('en-US', { timeZone: zone }).resolvedOptions().timeZone;
  } catch {
    return undefined;
  }
}

let graphZoneByRuntimeName: Map<string, string> | undefined;

/** Runtime zone name -> the documented Graph name, built once on first use. */
function graphZones(): Map<string, string> {
  if (graphZoneByRuntimeName) return graphZoneByRuntimeName;
  const map = new Map<string, string>();
  for (const entry of GRAPH_TIME_ZONES) {
    const m = /^(\S+)(?: \((\w+)\))?$/.exec(entry)!;
    const docName = m[1];
    const region = docName.split('/')[0];
    const candidates = m[2] ? [docName, `${region}/${m[2]}`] : [docName];
    for (const candidate of candidates) {
      const runtime = canonicalZone(candidate);
      if (runtime) map.set(runtime, docName);
    }
  }
  graphZoneByRuntimeName = map;
  return map;
}

/**
 * The local zone as a name Graph documents for writes: `UTC`, or the documented IANA name the
 * runtime zone maps to. Undefined when the local zone is not on Graph's list (e.g. `Europe/Amsterdam`)
 * or the runtime cannot name it (`TZ=JST-9` is +09:00 but has no name, so it must not become `UTC`).
 */
export function graphTimeZone(): string | undefined {
  const raw = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const zone = raw ? canonicalZone(raw) : undefined;
  if (!zone || zone === 'Etc/Unknown') return undefined;
  if (zone === 'UTC') return 'UTC';
  return graphZones().get(zone);
}

/** `YYYY-MM-DD` of `d` in the local zone. */
export function localDate(d: Date = new Date()): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

export interface DateTimeTimeZone {
  dateTime: string;
  timeZone: string;
}

/** `YYYY-MM-DD`, optionally `THH:MM[:SS[.fff]]` (or a space instead of `T`), then optionally `Z`/`±HH[:]MM`. */
const EVENT_TIME = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,9}))?)?(Z|[+-]\d{2}:?\d{2})?)?$/i;

/**
 * A Graph `dateTimeTimeZone` for a user-supplied event time. Only ISO 8601 forms are accepted (a
 * date-only value means local midnight); anything else, or a day/time that does not exist, throws
 * before any request. An explicit `timeZone` is sent as given. Otherwise a wall-clock time (no
 * `Z`/offset) is local: sent with the local zone when Graph documents it, else converted to the UTC
 * instant. A time with `Z` or an offset is always sent as UTC.
 */
export function eventTime(dateTime: string, timeZone?: string): DateTimeTimeZone {
  const m = EVENT_TIME.exec(dateTime);
  const [year, month, day, hour = 0, minute = 0, second = 0] = (m ?? []).slice(1, 7).map(v => (v === undefined ? undefined : Number(v)));
  if (!m || !isCalendarDate(`${m[1]}-${m[2]}-${m[3]}`) || hour! > 23 || minute! > 59 || second! > 59) {
    throw new Error(`Invalid date/time: "${dateTime}". Use ISO 8601, e.g. 2026-10-07T09:00.`);
  }
  const date = `${m[1]}-${m[2]}-${m[3]}`;
  const millis = m[7] ? Number(m[7].slice(0, 3).padEnd(3, '0')) : 0;
  const offset = m[8];
  const wallClock = m[4] === undefined ? `${date}T00:00:00` : dateTime.replace(' ', 'T');
  if (timeZone) return { dateTime: wallClock, timeZone };
  if (offset) {
    const iso = `${date}T${m[4]}:${m[5]}:${m[6] ?? '00'}.${String(millis).padStart(3, '0')}${offset.length === 5 ? `${offset.slice(0, 3)}:${offset.slice(3)}` : offset}`;
    return { dateTime: new Date(iso).toISOString().slice(0, 19), timeZone: 'UTC' };
  }
  const zone = graphTimeZone();
  if (zone) return { dateTime: wallClock, timeZone: zone };
  const local = new Date(year!, month! - 1, day!, hour, minute, second, millis);
  return { dateTime: local.toISOString().slice(0, 19), timeZone: 'UTC' };
}

/**
 * A date-only value (To Do due date, mail flag dates) as local midnight in the local zone when Graph
 * documents it, else midnight UTC: converting would move the date itself.
 */
export function dateOnly(date: string): DateTimeTimeZone {
  return { dateTime: `${calendarDate(date)}T00:00:00`, timeZone: graphTimeZone() ?? 'UTC' };
}

/**
 * `headers` plus `Prefer: outlook.timezone="<tz>"` (appended to an existing Prefer value), so Graph
 * returns event times in `tz` (default: the local zone) instead of offset-less UTC. Throws on a zone
 * name the runtime does not know. Returns a new object.
 */
export function preferTimezone(headers: Record<string, string> = {}, timeZone: string = localTimeZone()): Record<string, string> {
  try {
    new Intl.DateTimeFormat('en-US', { timeZone });
  } catch {
    throw new Error(`Unknown time zone: ${timeZone} (expected an IANA name such as Europe/Berlin).`);
  }
  const pref = `outlook.timezone="${timeZone}"`;
  const out = { ...headers };
  const key = Object.keys(out).find((k) => k.toLowerCase() === 'prefer');
  if (key && out[key]) out[key] = `${out[key]}, ${pref}`;
  else out[key ?? 'Prefer'] = pref;
  return out;
}

/** `data[key]` when `data` is an object that has `key`, else undefined (Graph payloads are untyped JSON). */
export function field(data: unknown, key: string): unknown {
  return typeof data === 'object' && data !== null && key in data ? Reflect.get(data, key) : undefined;
}

/** `data[key]` when it is a string, else undefined. */
export function stringField(data: unknown, key: string): string | undefined {
  const value = field(data, key);
  return typeof value === 'string' ? value : undefined;
}

/** The `value` array of a Graph collection payload, or undefined when `data` is not one. */
export function collectionItems(data: unknown): unknown[] | undefined {
  const value = field(data, 'value');
  return Array.isArray(value) ? value : undefined;
}
