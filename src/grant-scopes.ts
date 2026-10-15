import { spawnSync } from 'child_process';
import { GraphClient } from './graph-client.js';
import { CliError, ExitCode } from './errors.js';
import { odataQuote } from './utils.js';

/**
 * `ms365 auth grant-scopes`: make an app registration's Microsoft Graph permissions match the
 * scopes this CLI requests, and grant tenant-wide admin consent for them.
 *
 * Admin rights come from the Azure CLI's sign-in (one `az account get-access-token` call); the
 * Graph work itself goes through GraphClient, so --dry-run, --read-only, the audit log, retries
 * and exit codes behave as for every other command. The CLI's own app never needs directory
 * permissions. Without a usable `az`, the fallback is a one-click admin-consent link.
 */

export const GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000';
export const CONSENT_REDIRECT_URI = 'https://login.microsoftonline.com/common/oauth2/nativeclient';

export interface AzToken { accessToken: string; tenant?: string }

/** Injected in tests. Runs `az` and returns its result. */
export const azDeps = {
  run(args: string[]): { status: number | null; stdout: string; stderr: string; missing: boolean } {
    const r = spawnSync('az', args, { encoding: 'utf8', timeout: 60_000, killSignal: 'SIGKILL' });
    const missing = (r.error as NodeJS.ErrnoException | undefined)?.code === 'ENOENT';
    return { status: r.status, stdout: r.stdout ?? '', stderr: r.stderr ?? '', missing };
  },
};

export type AzResult = { ok: true; token: AzToken } | { ok: false; reason: 'missing' | 'not_logged_in' | 'failed'; detail: string };

/** An admin's Graph token from the Azure CLI, or why there isn't one. */
export function azGraphToken(tenant?: string): AzResult {
  const args = ['account', 'get-access-token', '--resource', 'https://graph.microsoft.com', '-o', 'json'];
  if (tenant && tenant !== 'common' && tenant !== 'organizations') args.push('--tenant', tenant);
  const r = azDeps.run(args);
  if (r.missing) return { ok: false, reason: 'missing', detail: 'the Azure CLI (az) is not installed' };
  if (r.status !== 0) {
    const notLoggedIn = /az login|not logged in|No subscription|AADSTS50076|interaction/i.test(r.stderr);
    return { ok: false, reason: notLoggedIn ? 'not_logged_in' : 'failed', detail: r.stderr.trim().split('\n').pop() ?? '' };
  }
  try {
    const j = JSON.parse(r.stdout) as { accessToken?: string; tenant?: string };
    if (!j.accessToken) throw new Error('no accessToken');
    return { ok: true, token: { accessToken: j.accessToken, tenant: j.tenant } };
  } catch {
    return { ok: false, reason: 'failed', detail: 'could not read the token from az output' };
  }
}

/** One-click v2 admin-consent link for these scopes (needs CONSENT_REDIRECT_URI registered on the app). */
export function adminConsentUrl(clientId: string, tenant: string, scopes: string[]): string {
  const u = new URL(`https://login.microsoftonline.com/${tenant}/v2.0/adminconsent`);
  u.searchParams.set('client_id', clientId);
  u.searchParams.set('scope', scopes.map(s => `https://graph.microsoft.com/${s}`).join(' '));
  u.searchParams.set('redirect_uri', CONSENT_REDIRECT_URI);
  u.searchParams.set('state', 'ms365-cli');
  return u.toString();
}

interface Scope { id: string; value: string }
interface RequiredResourceAccess { resourceAppId: string; resourceAccess: Array<{ id: string; type: string }> }

export interface GrantResult {
  appId: string;
  tenant?: string;
  /** Delegated permissions newly added to the app registration's list. */
  permissionsAdded: string[];
  /** Scopes newly added to the tenant-wide (AllPrincipals) consent grant. */
  consentAdded: string[];
  /** Requested scopes that were already listed and consented. */
  alreadyGranted: number;
  createdServicePrincipal: boolean;
}

/**
 * Add the missing delegated Graph permissions to the app registration and the tenant-wide
 * consent grant. Never removes anything: other tools may share the app registration.
 */
export async function grantScopes(client: GraphClient, appId: string, scopes: string[]): Promise<GrantResult> {
  const graph = await client.request(`/servicePrincipals(appId='${odataQuote(GRAPH_APP_ID)}')`, { queryParams: { '$select': 'id,oauth2PermissionScopes' } });
  const graphSp = graph.data as { id?: string; oauth2PermissionScopes?: Scope[] };
  const preview = client.isPreview(graph);
  const scopeIds = new Map((graphSp.oauth2PermissionScopes ?? []).map(s => [s.value, s.id]));
  if (!preview) {
    const unknown = scopes.filter(s => !scopeIds.has(s));
    if (unknown.length) {
      throw new CliError(`Microsoft Graph has no delegated permission named: ${unknown.join(', ')}`, { code: 'unknown_scope', exitCode: ExitCode.USAGE });
    }
  }
  const graphSpId = graphSp.id ?? '<graph-service-principal>';

  // 1. The app registration's permission list (requiredResourceAccess).
  const appRes = await client.request(`/applications(appId='${odataQuote(appId)}')`, { queryParams: { '$select': 'id,requiredResourceAccess' } });
  const app = appRes.data as { id?: string; requiredResourceAccess?: RequiredResourceAccess[] };
  const rra = app.requiredResourceAccess ?? [];
  const graphEntry = rra.find(r => r.resourceAppId === GRAPH_APP_ID) ?? { resourceAppId: GRAPH_APP_ID, resourceAccess: [] };
  const listedIds = new Set(graphEntry.resourceAccess.filter(a => a.type === 'Scope').map(a => a.id));
  const permissionsAdded = preview ? [...scopes] : scopes.filter(s => !listedIds.has(scopeIds.get(s)!));
  if (permissionsAdded.length) {
    const resourceAccess = [
      ...graphEntry.resourceAccess,
      ...permissionsAdded.map(s => ({ id: scopeIds.get(s) ?? `<id of ${s}>`, type: 'Scope' })),
    ];
    const updated = [...rra.filter(r => r.resourceAppId !== GRAPH_APP_ID), { resourceAppId: GRAPH_APP_ID, resourceAccess }];
    await client.request(`/applications/${app.id ?? '<application-object-id>'}`, { method: 'PATCH', body: JSON.stringify({ requiredResourceAccess: updated }) });
  }

  // 2. The app's service principal in this tenant (created on first consent if missing).
  const spRes = await client.request('/servicePrincipals', { queryParams: { '$filter': `appId eq '${odataQuote(appId)}'`, '$select': 'id' } });
  let spId = ((spRes.data as { value?: Array<{ id: string }> }).value ?? [])[0]?.id;
  let createdServicePrincipal = false;
  if (!spId && !preview) {
    const created = await client.request('/servicePrincipals', { method: 'POST', body: JSON.stringify({ appId }) });
    spId = (created.data as { id: string }).id;
    createdServicePrincipal = true;
  }
  spId ??= '<client-service-principal>';

  // 3. The tenant-wide delegated consent grant (consentType AllPrincipals): union, never replace.
  const grantsRes = await client.request('/oauth2PermissionGrants', {
    queryParams: { '$filter': `clientId eq '${odataQuote(spId)}' and consentType eq 'AllPrincipals' and resourceId eq '${odataQuote(graphSpId)}'` },
  });
  const grant = ((grantsRes.data as { value?: Array<{ id: string; scope?: string }> }).value ?? [])[0];
  const consented = new Set((grant?.scope ?? '').split(' ').filter(Boolean));
  const consentAdded = preview ? [...scopes] : scopes.filter(s => !consented.has(s));
  if (consentAdded.length) {
    const scope = [...consented, ...consentAdded].join(' ');
    if (grant) {
      await client.request(`/oauth2PermissionGrants/${grant.id}`, { method: 'PATCH', body: JSON.stringify({ scope }) });
    } else {
      await client.request('/oauth2PermissionGrants', {
        method: 'POST',
        body: JSON.stringify({ clientId: spId, consentType: 'AllPrincipals', resourceId: graphSpId, scope }),
      });
    }
  }

  const alreadyGranted = scopes.filter(s => !permissionsAdded.includes(s) && !consentAdded.includes(s)).length;
  return { appId, permissionsAdded, consentAdded, alreadyGranted, createdServicePrincipal };
}
