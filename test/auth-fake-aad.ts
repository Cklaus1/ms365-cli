// Canned Microsoft identity platform (AAD v2) responses, so real MSAL runs device-code login and
// refresh-token flows offline. Used both as e2e routes (fake-fetch.mjs) and by in-process fetch stubs.
import type { FakeRoute } from './e2e/run-cli.js';

export const TENANT_ID = '11111111-1111-1111-1111-111111111111';
export const USER_OID = '22222222-2222-2222-2222-222222222222';
export const USERNAME = 'ada@contoso.example';
/** MSAL's homeAccountId for the fake user: `<uid>.<utid>` from client_info. */
export const HOME_ACCOUNT_ID = `${USER_OID}.${TENANT_ID}`;

const b64url = (value: unknown) => Buffer.from(JSON.stringify(value)).toString('base64url');

export interface TokenOptions {
  clientId: string;
  accessToken: string;
  refreshToken: string;
  /** Seconds; anything below MSAL's 300 s renewal offset makes the next silent call refresh. */
  expiresIn?: number;
  /** Granted scopes; MSAL serves a cached access token only when it covers the requested scopes. */
  scope?: string;
}

/** Body of a successful token endpoint response (device-code or refresh-token grant). */
export function tokenResponse(o: TokenOptions): Record<string, unknown> {
  const now = Math.floor(Date.now() / 1000);
  const idToken = [
    b64url({ alg: 'none', typ: 'JWT' }),
    b64url({
      aud: o.clientId,
      iss: `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,
      tid: TENANT_ID,
      oid: USER_OID,
      sub: 'subject',
      preferred_username: USERNAME,
      name: 'Ada Lovelace',
      iat: now,
      nbf: now,
      exp: now + 3600,
      ver: '2.0',
    }),
    'signature',
  ].join('.');
  const expiresIn = o.expiresIn ?? 3600;
  return {
    token_type: 'Bearer',
    scope: o.scope ?? 'https://graph.microsoft.com/.default',
    expires_in: expiresIn,
    ext_expires_in: expiresIn,
    access_token: o.accessToken,
    refresh_token: o.refreshToken,
    id_token: idToken,
    client_info: b64url({ uid: USER_OID, utid: TENANT_ID }),
  };
}

export const DEVICE_CODE_MESSAGE = 'To sign in, open https://microsoft.com/devicelogin and enter the code FAKE-CODE.';

export const DEVICE_CODE_RESPONSE = {
  user_code: 'FAKE-CODE',
  device_code: 'device-code-1',
  verification_uri: 'https://microsoft.com/devicelogin',
  expires_in: 900,
  interval: 1,
  message: DEVICE_CODE_MESSAGE,
};

export const ME = { displayName: 'Ada Lovelace', userPrincipalName: USERNAME };

/** Fake-fetch routes for login + refresh: every token request succeeds with the given tokens. */
export function aadRoutes(o: TokenOptions & { me?: FakeRoute | false }): FakeRoute[] {
  const routes: FakeRoute[] = [
    { match: '/oauth2/v2\\.0/devicecode', method: 'POST', body: DEVICE_CODE_RESPONSE },
    { match: '/oauth2/v2\\.0/token', method: 'POST', body: tokenResponse(o) },
  ];
  if (o.me !== false) routes.push(o.me ?? { match: 'graph\\.microsoft\\.com/v1\\.0/me$', body: ME });
  return routes;
}

/** AAD error body, as returned by the token/devicecode endpoints with HTTP 400. */
export function aadError(code: number, error = 'invalid_grant', suberror?: string): Record<string, unknown> {
  return {
    error,
    error_description: `AADSTS${code}: Fake failure for code ${code}.\r\nTrace ID: t\r\nCorrelation ID: c\r\nTimestamp: 2026-10-06 00:00:00Z`,
    error_codes: [code],
    ...(suberror ? { suberror } : {}),
  };
}
