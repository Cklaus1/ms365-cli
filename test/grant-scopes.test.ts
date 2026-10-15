import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { runCommand, jsonResponse } from './helpers/cli.js';
import { registerAuthCommands } from '../src/commands/auth.js';
import { azDeps, adminConsentUrl, GRAPH_APP_ID } from '../src/grant-scopes.js';

const APP = '286411c0-c092-4520-b4f3-7dce2a309999';
const TENANT = 'c3087746-51aa-4f4f-8fa3-755cd4df2293';
const SCOPES = ['User.Read', 'Mail.Send', 'Group.Read.All'];
const IDS: Record<string, string> = { 'User.Read': 'id-user', 'Mail.Send': 'id-send', 'Group.Read.All': 'id-group', 'Chat.ReadWrite.All': 'id-chatrw' };

const fakeAuth = (scp = 'User.Read') => {
  const jwt = (s: string) => ['{"alg":"none"}', JSON.stringify({ scp: s, upn: 'u@x.com', exp: Math.floor(Date.now() / 1000) + 3600 })]
    .map(p => Buffer.from(p).toString('base64url')).join('.') + '.sig';
  return {
    getClientId: () => APP,
    getTenantId: () => TENANT,
    requestedScopes: () => [...SCOPES],
    initialize: async () => {},
    getToken: async (force = false) => jwt(force ? 'User.Read Mail.Send Group.Read.All' : scp),
  };
};

function graphRoutes(opts: { listed?: string[]; consented?: string; grantExists?: boolean; spExists?: boolean } = {}) {
  const listed = opts.listed ?? ['User.Read'];
  return [
    ['GET', `/servicePrincipals(appId='${GRAPH_APP_ID}')`, () => ({ id: 'graph-sp', oauth2PermissionScopes: Object.entries(IDS).map(([value, id]) => ({ value, id })) })],
    ['GET', `/applications(appId='${APP}')`, () => ({ id: 'app-obj', requiredResourceAccess: [
      { resourceAppId: 'other-api', resourceAccess: [{ id: 'x', type: 'Scope' }] },
      { resourceAppId: GRAPH_APP_ID, resourceAccess: [...listed.map(s => ({ id: IDS[s], type: 'Scope' })), { id: 'role-1', type: 'Role' }] },
    ] })],
    ['GET', '/servicePrincipals?', () => ({ value: opts.spExists === false ? [] : [{ id: 'client-sp' }] })],
    ['POST', '/servicePrincipals', () => jsonResponse({ id: 'new-sp' }, 201)],
    ['GET', '/oauth2PermissionGrants', () => ({ value: opts.grantExists === false ? [] : [{ id: 'grant-1', scope: opts.consented ?? 'User.Read Chat.ReadWrite.All' }] })],
    ['PATCH', '/', () => new Response(null, { status: 204 })],
    ['POST', '/oauth2PermissionGrants', () => jsonResponse({ id: 'grant-new' }, 201)],
  ] as never;
}

describe('ms365 auth grant-scopes', () => {
  let azCalls: string[][];
  beforeEach(() => { azCalls = []; });
  afterEach(() => vi.restoreAllMocks());
  const az = (result: Partial<ReturnType<typeof azDeps.run>>) => vi.spyOn(azDeps, 'run').mockImplementation((args) => {
    azCalls.push(args);
    return { status: 0, stdout: JSON.stringify({ accessToken: 'az-admin-token', tenant: TENANT }), stderr: '', missing: false, ...result };
  });
  const run = (argv: string[], routes = graphRoutes(), extra = {}) =>
    runCommand((p) => registerAuthCommands(p, () => fakeAuth() as never), ['auth', 'grant-scopes', ...argv], { routes, ...extra });

  it('adds missing permissions and unions the consent grant, never removing anything', async () => {
    az({});
    const r = await run(['-o', 'json']);
    expect(r.exitCode, r.stderr).toBe(0);
    // uses the az admin token for Graph
    expect(r.calls.every(c => c.headers.Authorization === 'Bearer az-admin-token')).toBe(true);
    expect(azCalls[0]).toEqual(['account', 'get-access-token', '--resource', 'https://graph.microsoft.com', '-o', 'json', '--tenant', TENANT]);
    const appPatch = r.calls.find(c => c.method === 'PATCH' && c.path === '/applications/app-obj')!;
    const rra = (appPatch.body as { requiredResourceAccess: Array<{ resourceAppId: string; resourceAccess: Array<{ id: string; type: string }> }> }).requiredResourceAccess;
    expect(rra.find(x => x.resourceAppId === 'other-api')).toBeDefined(); // other APIs kept
    const graphIds = rra.find(x => x.resourceAppId === GRAPH_APP_ID)!.resourceAccess.map(a => a.id);
    expect(graphIds).toEqual(['id-user', 'role-1', 'id-send', 'id-group']); // existing kept (incl. app role), new appended
    const grantPatch = r.calls.find(c => c.method === 'PATCH' && c.path === '/oauth2PermissionGrants/grant-1')!;
    expect((grantPatch.body as { scope: string }).scope).toBe('User.Read Chat.ReadWrite.All Mail.Send Group.Read.All'); // union
    const out = JSON.parse(r.stdout);
    expect(out).toMatchObject({ appId: APP, permissionsAdded: ['Mail.Send', 'Group.Read.All'], consentAdded: ['Mail.Send', 'Group.Read.All'], alreadyGranted: 1, tenant: TENANT });
    expect(out.tokenScopesAdded).toEqual(['Mail.Send', 'Group.Read.All']);
  });

  it('creates the service principal and the grant when missing', async () => {
    az({});
    const r = await run(['-o', 'json'], graphRoutes({ spExists: false, grantExists: false }));
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.calls.some(c => c.method === 'POST' && c.path === '/servicePrincipals')).toBe(true);
    const post = r.calls.find(c => c.method === 'POST' && c.path === '/oauth2PermissionGrants')!;
    expect(post.body).toEqual({ clientId: 'new-sp', consentType: 'AllPrincipals', resourceId: 'graph-sp', scope: 'User.Read Mail.Send Group.Read.All' });
  });

  it('makes no writes when everything is already granted', async () => {
    az({});
    const r = await run(['-o', 'json'], graphRoutes({ listed: SCOPES, consented: SCOPES.join(' ') }));
    expect(r.exitCode).toBe(0);
    expect(r.calls.filter(c => c.method !== 'GET')).toHaveLength(0);
    expect(JSON.parse(r.stdout).alreadyGranted).toBe(3);
  });

  it('falls back to the admin-consent link when az is missing', async () => {
    az({ missing: true, status: null });
    const r = await run([]);
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(0);
    expect(r.stdout).toContain(`/${TENANT}/v2.0/adminconsent?client_id=${APP}`);
    expect(r.stderr).toMatch(/not installed/);
    expect(r.stderr).toContain('nativeclient');
  });

  it('falls back with an az login hint when az is not signed in', async () => {
    az({ status: 1, stdout: '', stderr: "ERROR: Please run 'az login' to setup account." });
    const r = await run([]);
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain(`az login --tenant ${TENANT}`);
    expect(r.stdout).toContain('adminconsent');
  });

  it('--link prints only the link and never runs az', async () => {
    const spy = az({});
    const r = await run(['--link']);
    expect(spy).not.toHaveBeenCalled();
    expect(r.stdout.trim()).toBe(adminConsentUrl(APP, TENANT, SCOPES));
  });

  it('refuses when az is signed in to a different tenant (exit 2)', async () => {
    az({ stdout: JSON.stringify({ accessToken: 't', tenant: 'other-tenant' }) });
    const r = await run([]);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('a 403 from Graph exits 4 with an admin-role hint', async () => {
    az({});
    const routes = [['GET', '/', () => new Response('{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges"}}', { status: 403 })]] as never;
    const r = await run([], routes);
    expect(r.exitCode).toBe(4);
    expect(r.stderr).toMatch(/Application Administrator/);
  });

  it('dry-run never calls az and previews the writes', async () => {
    const spy = az({});
    const r = await run([], graphRoutes(), { dryRun: true });
    expect(spy).not.toHaveBeenCalled();
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toMatch(/\[dry-run\] PATCH .*applications/);
    expect(r.stderr).toMatch(/\[dry-run\] PATCH .*oauth2PermissionGrants|\[dry-run\] POST .*oauth2PermissionGrants/);
  });

  it('read-only mode blocks the writes (exit 4)', async () => {
    az({});
    const r = await runCommand((p) => { p.option('--read-only'); registerAuthCommands(p, () => fakeAuth() as never); }, ['--read-only', 'auth', 'grant-scopes'], { routes: graphRoutes() });
    expect(r.exitCode).toBe(4);
  });
});
