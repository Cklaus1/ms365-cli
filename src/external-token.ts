import { CliError, ExitCode } from './errors.js';
import { decodeAccessToken } from './auth.js';
import { profileType, type Profile } from './profiles.js';

/**
 * AH-2: the `external` token source. A host (Chief's broker, an MCP host) supplies the access token in
 * MS365_CLI_TOKEN; ms365 never caches it, never loads MSAL, and checks it before the first request.
 * Only the real process environment counts: src/load-env.ts drops values that came from a `.env` file.
 */

/** External mode: a token in MS365_CLI_TOKEN, or a `type: host` profile (which never falls back to the cache). */
export function externalTokenMode(profile: Profile): boolean {
  return !!process.env.MS365_CLI_TOKEN || profileType(profile) === 'host';
}

/** Microsoft Graph's audience, as a URL (v1 tokens, with or without the slash) or the app ID (v2 tokens). */
const GRAPH_AUDIENCES: Record<string, true> = {
  'https://graph.microsoft.com': true,
  'https://graph.microsoft.com/': true,
  '00000003-0000-0000-c000-000000000000': true,
};
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
/** Refuse a token this close to expiry: a run that starts with it may not finish (same skew as MSAL refresh). */
export const EXPIRY_MARGIN_MS = 5 * 60_000;
const FRESH_TOKEN = 'get a fresh token from the host (the agent host sets MS365_CLI_TOKEN)';

export interface ExternalTokenExpectation {
  /** A `type: host` profile: no token or an unverifiable one is refused. */
  host: boolean;
  /** The profile's tenant: a GUID is compared with the token's `tid` (a domain cannot be). */
  tenantId?: string;
  /** --account / MS365_CLI_ACCOUNT / the profile's account: a UPN, an object ID or a homeAccountId (`<oid>.<tid>`). */
  account?: string;
  profileName?: string | null;
}

const authError = (message: string, code: string, fix = FRESH_TOKEN) => new CliError(message, { code, exitCode: ExitCode.AUTH, fix });

/** `MS365_CLI_TOKEN`, or exit 3 `token_missing` (only reachable in external mode without one: a host profile). */
export function requireExternalToken(profileName: string | null | undefined): string {
  const token = process.env.MS365_CLI_TOKEN;
  if (token) return token;
  throw authError(`Host profile "${profileName ?? ''}" takes its token from the agent host, but MS365_CLI_TOKEN is not set.`, 'token_missing',
    'the host sets MS365_CLI_TOKEN in this process\'s environment (ms365 never signs in under a host profile)');
}

/**
 * Refuse a supplied token before any request (AH-2): expired or within 5 min of it → `token_expired`; not for
 * Microsoft Graph → `token_audience`; another tenant or user than the profile pins → `token_identity_mismatch`.
 * All exit 3. An opaque (non-JWT) token cannot be checked: refused under a host profile, allowed otherwise
 * (personal-account tokens are opaque).
 */
export function checkExternalToken(token: string, expect: ExternalTokenExpectation, now = Date.now()): void {
  let claims: Record<string, unknown>;
  try {
    claims = decodeAccessToken(token);
  } catch {
    if (!expect.host) return;
    throw authError(`Host profile "${expect.profileName ?? ''}" cannot verify the supplied token: it is not a JWT, so its audience, tenant and expiry are unknown.`,
      'token_audience', 'the host supplies a Microsoft Graph access token (a JWT) for this tenant');
  }
  const str = (v: unknown) => (typeof v === 'string' ? v : undefined);
  if (typeof claims.exp === 'number' && claims.exp * 1000 - EXPIRY_MARGIN_MS <= now) {
    const when = new Date(claims.exp * 1000).toISOString();
    throw authError(`The supplied token expires at ${when}, less than 5 minutes from now (or has expired).`, 'token_expired');
  }
  const aud = str(claims.aud);
  if (!aud || GRAPH_AUDIENCES[aud] !== true) {
    throw authError(`The supplied token is for ${aud ? `"${aud}"` : 'no audience'}, not Microsoft Graph.`, 'token_audience',
      'the host requests the token for https://graph.microsoft.com');
  }
  const mismatch = (what: string) => authError(`The supplied token is for ${what}.`, 'token_identity_mismatch',
    `get a token for the ${expect.profileName ? `tenant and account of profile "${expect.profileName}"` : 'pinned tenant and account'} from the host`);
  const tid = str(claims.tid);
  if (expect.tenantId && GUID.test(expect.tenantId) && tid?.toLowerCase() !== expect.tenantId.toLowerCase()) {
    throw mismatch(`another tenant than the pinned one (token tid ${tid ?? 'missing'})`);
  }
  if (expect.account) {
    const wanted = expect.account.toLowerCase();
    const user = (str(claims.upn) ?? str(claims.preferred_username) ?? str(claims.unique_name))?.toLowerCase();
    const oid = str(claims.oid)?.toLowerCase();
    const ok = wanted.includes('@') ? user === wanted : oid !== undefined && (wanted === oid || wanted.startsWith(`${oid}.`));
    if (!ok) throw mismatch(`another account than the pinned one (${wanted.includes('@') ? `token user ${user ?? 'missing'}` : `token oid ${oid ?? 'missing'}`})`);
  }
}

/** A 401 for a supplied token: ms365 cannot refresh it (exit 3, never "ms365 auth login"). */
export const tokenRejectedError = () =>
  authError('Microsoft Graph rejected the supplied token (401): it was revoked, expired early or lacks consent.', 'token_rejected');

/** `auth login/logout/…` while a token is supplied: there is no ms365 sign-in or cache to touch (exit 2). */
export function externalTokenModeError(subcommand: string): CliError {
  return new CliError(`auth ${subcommand} manages ms365's own sign-in, but a token is supplied by the host (MS365_CLI_TOKEN or a host profile), so nothing is cached.`, {
    code: 'external_token_mode', exitCode: ExitCode.USAGE,
    fix: 'unset MS365_CLI_TOKEN and drop the host profile to manage ms365\'s own sign-in',
  });
}
