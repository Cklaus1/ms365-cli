/**
 * Strict parser for the `Tailscale-App-Capabilities` header set by `tailscale serve --accept-app-caps <name>`
 * (REMOTE-1b). The header is a JSON object: capability name → array of grant values. Our values are
 * `{ agent?: string, allow: string[] }`; Serve concatenates the values of every matching grant.
 * Anything unexpected is a denial: there is no partial decoding and no default grant.
 */

export const APP_CAPS_HEADER = 'tailscale-app-capabilities';

export interface Caller {
  /** Agent label(s) from the grant values, for audit and discovery (`unknown` when no grant names one). */
  agent: string;
  /** Union of every `allow` list. */
  allow: Set<string>;
}

export class CapsHeaderError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'CapsHeaderError';
  }
}

/**
 * Parse the header value(s) for capability key `appCap`; throws CapsHeaderError on anything but a usable grant.
 * Pass every received copy (`req.headersDistinct`): Node's `req.headers` joins repeated copies with ", ".
 */
export function parseAppCapabilities(header: string | string[] | undefined, appCap: string): Caller {
  if (Array.isArray(header)) {
    if (header.length > 1) throw new CapsHeaderError(`Repeated ${APP_CAPS_HEADER} header`);
    header = header[0];
  }
  if (header === undefined || header === '') throw new CapsHeaderError(`Missing ${APP_CAPS_HEADER} header`);
  if (header.trimStart().startsWith('=?')) throw new CapsHeaderError(`RFC 2047 encoded ${APP_CAPS_HEADER} header is not accepted`);

  let parsed: unknown;
  try {
    parsed = JSON.parse(header);
  } catch {
    throw new CapsHeaderError(`Malformed ${APP_CAPS_HEADER} header`);
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new CapsHeaderError(`Malformed ${APP_CAPS_HEADER} header`);
  }
  if (!Object.hasOwn(parsed, appCap)) throw new CapsHeaderError(`No ${appCap} grant`);

  const values: unknown = Reflect.get(parsed, appCap);
  if (!Array.isArray(values) || values.length === 0) throw new CapsHeaderError(`Malformed ${appCap} grant`);
  const agents = new Set<string>();
  const allow = new Set<string>();
  for (const value of values) {
    if (typeof value !== 'object' || value === null || Array.isArray(value)) throw new CapsHeaderError(`Malformed ${appCap} grant`);
    const agent = 'agent' in value ? value.agent : undefined;
    const list = 'allow' in value ? value.allow : undefined;
    if (agent !== undefined && typeof agent !== 'string') throw new CapsHeaderError(`Malformed ${appCap} grant: agent`);
    if (!Array.isArray(list) || !list.every(c => typeof c === 'string')) {
      throw new CapsHeaderError(`Malformed ${appCap} grant: allow`);
    }
    if (agent) agents.add(agent);
    for (const c of list) allow.add(c);
  }
  if (allow.size === 0) throw new CapsHeaderError(`Empty ${appCap} grant`);
  return { agent: agents.size > 0 ? [...agents].join(',') : 'unknown', allow };
}
