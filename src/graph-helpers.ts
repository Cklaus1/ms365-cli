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
