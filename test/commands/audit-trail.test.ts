import { describe, it, expect, vi, afterEach } from 'vitest';
import { runCommand, jsonResponse, type Route } from '../helpers/cli.js';
import { setFields } from '../../src/formatter.js';
import {
  registerAuditTrailCommands, projectRecord, accessFacts, timing, tenant, MAILBOX_FIELDS, DEFAULT_OPERATIONS, ACCESS_FIELDS, APP_FIELDS,
  GRAPH_APP_ID, EXCHANGE_APP_ID,
} from '../../src/commands/audit-trail.js';
import { commandMeta } from '../../src/catalog.js';

/**
 * audit-trail (AO-9, AO-11..AO-16, AO-25). The query engine is asynchronous: POST a query, poll it, then page its
 * records. Polling sleeps go through `timing`, which these tests replace with an instant fake clock.
 */

const QUERIES = '/security/auditLog/queries';
const OWNER = 'Alex@Contoso.com';

// Content that must never reach any output stream (AO-9, AC 7).
const SUBJECT_ITEM = 'Q3 layoffs SECRET-SUBJECT-ITEM';
const SUBJECT_AFFECTED = 'Merger term sheet SECRET-SUBJECT-AFFECTED';
const IMID = '<SECRET-IMID-0001@mail.contoso.com>';
const RECIPIENT = 'secret-recipient@external.example';
const mark = (field: string) => `ZZLEAK_${field}_ZZ`;
const FORBIDDEN = [SUBJECT_ITEM, SUBJECT_AFFECTED, IMID, RECIPIENT, 'SECRET-', 'ZZLEAK_'];

/** An auditData object with every non-allow-listed field carrying a unique marker. */
function leakyAuditData(owner = OWNER, extra: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    // allow-listed (printed or used for filtering)
    MailboxOwnerUPN: owner,
    LogonType: 2,
    ClientInfoString: 'Client=REST;Client=RESTSystem;;',
    AppId: '11111111-2222-3333-4444-555555555555',
    OperationCount: 3,
    ExternalAccess: false,
    // never printed
    Item: { Subject: SUBJECT_ITEM, Id: mark('ItemId'), ParentFolder: { Path: mark('ParentPath'), Name: mark('ParentName') } },
    AffectedItems: [{ Subject: SUBJECT_AFFECTED, Id: mark('AffectedId'), InternetMessageId: mark('AffectedImid') }],
    Folders: [{ Id: mark('FolderId'), Path: mark('FolderPath'), FolderItems: [{ InternetMessageId: IMID, SizeInBytes: 4242, Id: mark('FolderItemId') }] }],
    Recipients: [RECIPIENT], To: [RECIPIENT], Cc: RECIPIENT, SendAsUserSmtp: RECIPIENT,
    MailboxGuid: mark('MailboxGuid'), MailboxOwnerSid: mark('OwnerSid'), MailboxOwnerMasterAccountSid: mark('MasterSid'),
    LogonUserSid: mark('LogonSid'), InternalLogonType: mark('InternalLogonType'), OrganizationName: mark('OrgName'),
    OriginatingServer: mark('Server'), SessionId: mark('SessionId'), ClientProcessName: mark('ProcessName'),
    ClientVersion: mark('ClientVersion'), ClientRequestId: mark('ClientRequestId'), ResultStatus: mark('ResultStatus'),
    OperationProperties: [{ Name: mark('OpPropName'), Value: mark('OpPropValue') }],
    AppAccessContext: { ClientAppName: mark('ClientAppName'), APIId: mark('ApiId'), UniqueTokenId: mark('TokenId') },
    ModifiedProperties: [mark('Modified')], Attachments: mark('Attachments'), Body: mark('Body'),
    ...extra,
  };
}

let seq = 0;
/** A Graph auditLogRecord whose envelope also carries markers in every non-allow-listed property. */
function record(owner = OWNER, extra: Record<string, unknown> = {}) {
  seq++;
  return {
    id: mark(`RecordId${seq}`),
    createdDateTime: '2026-10-10T09:00:00Z',
    auditLogRecordType: mark('RecordType'),
    operation: 'MailItemsAccessed',
    organizationId: mark('OrgId'),
    userType: 'regular',
    userId: mark('UserId'),
    service: mark('Service'),
    objectId: mark('ObjectId'),
    userPrincipalName: 'delegate@contoso.com',
    clientIp: '203.0.113.7',
    administrativeUnits: [mark('AdminUnit')],
    auditData: leakyAuditData(owner, extra),
  };
}

interface Lifecycle {
  statuses?: string[];
  pages?: unknown[][];
  query?: Record<string, unknown>;
}

/** Routes for one query: POST → notStarted, GET query → statuses in turn, records paged by nextLink. */
function lifecycle({ statuses = ['running', 'succeeded'], pages = [[record()]], query = {} }: Lifecycle = {}): Route[] {
  let polled = 0;
  return [
    ['POST', QUERIES, () => jsonResponse({ id: 'q1', status: 'notStarted' }, 201)],
    ['GET', /^\/security\/auditLog\/queries\/q1\/records/, (call) => {
      const m = /\$skiptoken=p(\d+)/.exec(call.path);
      const i = m ? Number(m[1]) : 0;
      const body: Record<string, unknown> = { value: pages[i] ?? [] };
      if (i + 1 < pages.length) body['@odata.nextLink'] = `https://graph.microsoft.com/v1.0${QUERIES}/q1/records?$skiptoken=p${i + 1}`;
      return body;
    }],
    ['GET', /^\/security\/auditLog\/queries\/q1$/, () => {
      const status = statuses[Math.min(polled, statuses.length - 1)];
      polled++;
      return { id: 'q1', status, ...query };
    }],
  ];
}

/** Replace the poll clock: sleep advances a fake now() instantly. Returns the sleeps requested. */
function fakeClock() {
  let now = 0;
  const sleeps: number[] = [];
  vi.spyOn(timing, 'now').mockImplementation(() => now);
  vi.spyOn(timing, 'sleep').mockImplementation(async (ms: number) => { sleeps.push(ms); now += ms; });
  return sleeps;
}

afterEach(() => { vi.restoreAllMocks(); });

const run = (argv: string[], opts: Parameters<typeof runCommand>[2] = {}) => runCommand(registerAuditTrailCommands, argv, opts);

describe('audit-trail mailbox: query lifecycle (AO-9, AO-13)', () => {
  it('POSTs the query, polls notStarted → running → succeeded with backoff, and prints every record page', async () => {
    const sleeps = fakeClock();
    const pages = [[record(), record()], [record()]];
    const r = await run(['audit-trail', 'mailbox', OWNER, '--since', '2026-10-01T00:00Z', '--until', '2026-10-15T00:00Z', '-o', 'json'],
      { routes: lifecycle({ pages }) });
    expect(r.exitCode, r.stderr).toBe(0);

    const post = r.calls.find(c => c.method === 'POST')!;
    expect(post.path).toBe(QUERIES);
    expect(post.body).toMatchObject({
      filterStartDateTime: '2026-10-01T00:00:00.000Z',
      filterEndDateTime: '2026-10-15T00:00:00.000Z',
      recordTypeFilters: ['exchangeItem', 'exchangeItemGroup', 'exchangeItemAggregated'],
      operationFilters: [...DEFAULT_OPERATIONS],
      keywordFilter: OWNER,
    });
    expect(post.body).not.toHaveProperty('userPrincipalNameFilters');
    expect(post.body).not.toHaveProperty('serviceFilter');
    expect(post.body).not.toHaveProperty('serviceFilters');

    expect(r.calls.filter(c => /queries\/q1$/.test(c.path))).toHaveLength(2);
    expect(sleeps).toEqual([2000, 4000]);
    expect(r.calls.filter(c => c.path.includes('/records'))).toHaveLength(2);

    const out = JSON.parse(r.stdout);
    expect(out.value).toHaveLength(3);
    expect(Object.keys(out.value[0]).sort()).toEqual([...MAILBOX_FIELDS].sort());
    expect(out.value[0]).toMatchObject({
      operation: 'MailItemsAccessed', actor: 'delegate@contoso.com', logonType: 'delegated', operationCount: 3, folderCount: 1,
      appId: '11111111-2222-3333-4444-555555555555', clientIp: '203.0.113.7', externalAccess: false,
    });
  });

  it('backoff doubles from 2 s and caps at 15 s', async () => {
    const sleeps = fakeClock();
    const r = await run(['audit-trail', 'mailbox', OWNER, '-o', 'json'],
      { routes: lifecycle({ statuses: ['notStarted', 'running', 'running', 'running', 'running', 'succeeded'] }) });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(sleeps).toEqual([2000, 4000, 8000, 15000, 15000, 15000]);
  });

  it('--operations and --by shape the query', async () => {
    fakeClock();
    const r = await run(['audit-trail', 'mailbox', OWNER, '--operations', 'MailItemsAccessed,SendAs', '--by', 'delegate@contoso.com', '-o', 'json'],
      { routes: lifecycle() });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.calls[0].body).toMatchObject({ operationFilters: ['MailItemsAccessed', 'SendAs'], userPrincipalNameFilters: ['delegate@contoso.com'] });
  });

  it('a failed query exits 1 naming the status', async () => {
    fakeClock();
    const r = await run(['audit-trail', 'mailbox', OWNER], { routes: lifecycle({ statuses: ['running', 'failed'] }) });
    expect(r.exitCode).toBe(1);
    expect(r.stderr).toMatch(/failed/);
    expect(r.calls.some(c => c.path.includes('/records'))).toBe(false);
  });

  it('a cancelled query exits 1 naming the status', async () => {
    fakeClock();
    const r = await run(['audit-trail', 'mailbox', OWNER], { routes: lifecycle({ statuses: ['cancelled'] }) });
    expect(r.exitCode).toBe(1);
    expect(r.stderr).toMatch(/cancelled/);
    expect(r.calls.some(c => c.path.includes('/records'))).toBe(false);
  });

  it('isRecordCountLimitExceeded reports truncation on stderr and still exits 0', async () => {
    fakeClock();
    const r = await run(['audit-trail', 'mailbox', OWNER, '-o', 'json'], {
      routes: lifecycle({ query: { isRecordCountLimitExceeded: true, recordCountLimit: 1000000, approximateReturnedRecordCount: 5 } }),
    });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.stderr).toMatch(/truncated at 1000000 records/);
    expect(r.stderr).toMatch(/--since\/--until or --operations/);
    expect(JSON.parse(r.stdout).value).toHaveLength(1);
  });

  it('does not warn when the limit flag is false, even if the approximate count equals the limit', async () => {
    fakeClock();
    const r = await run(['audit-trail', 'mailbox', OWNER, '-o', 'json'], {
      routes: lifecycle({ query: { isRecordCountLimitExceeded: false, recordCountLimit: 1, approximateReturnedRecordCount: 1 } }),
    });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).not.toMatch(/truncated/);
  });

  it('a query still running at --timeout exits 1 with a narrow-the-window hint', async () => {
    const sleeps = fakeClock();
    const r = await run(['audit-trail', 'mailbox', OWNER, '--timeout', '10'], { routes: lifecycle({ statuses: ['running'] }) });
    expect(r.exitCode).toBe(1);
    expect(r.stderr).toMatch(/did not finish within 10 s/);
    expect(r.stderr).toMatch(/narrow/);
    expect(sleeps.reduce((a, b) => a + b, 0)).toBe(10_000);
    expect(r.calls.some(c => c.path.includes('/records'))).toBe(false);
  });

  it('prints the coverage line on stderr (AO-12) and flags a window beyond Standard retention', async () => {
    fakeClock();
    const r = await run(['audit-trail', 'mailbox', OWNER, '--since', '-200d', '-o', 'json'], { routes: lifecycle({ pages: [[]] }) });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.stderr).toMatch(/Coverage: .*→/);
    expect(r.stderr).toMatch(/MailItemsAccessed/);
    expect(r.stderr).toMatch(/an empty result is not proof of no access/);
    expect(r.stderr).toMatch(/180 days/);
    expect(JSON.parse(r.stdout).value).toEqual([]);
  });

  it('rejects an invalid mailbox UPN with exit 2 and no request', async () => {
    const r = await run(['audit-trail', 'mailbox', 'not-an-upn']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('rejects --since after --until with exit 2 and no request', async () => {
    const r = await run(['audit-trail', 'mailbox', OWNER, '--since', '2026-10-15T00:00Z', '--until', '2026-10-01T00:00Z']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});

describe('audit-trail mailbox: client-side mailbox filter (AO-9, AC 10)', () => {
  it('keeps only records whose auditData.MailboxOwnerUPN matches the mailbox, case-insensitively', async () => {
    fakeClock();
    const recs = [
      record('alex@contoso.com', { OperationCount: 1 }),
      record('ALEX@CONTOSO.COM', { OperationCount: 2 }),
      record('someone.else@contoso.com', { OperationCount: 3 }),
      record('alex@contoso.com.evil', { OperationCount: 4 }),
      { ...record(), auditData: { OperationCount: 5 } }, // no owner at all
    ];
    const r = await run(['audit-trail', 'mailbox', 'Alex@Contoso.com', '-o', 'json'], { routes: lifecycle({ pages: [recs] }) });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(JSON.parse(r.stdout).value.map((v: { operationCount: number }) => v.operationCount)).toEqual([1, 2]);
  });
});

describe('audit-trail output allow-list (AO-9, AC 7) — property test', () => {
  const formats = ['json', 'ndjson', 'table', 'csv', 'text', 'yaml'];
  for (const format of formats) {
    it(`no subject, InternetMessageId, recipient or non-allow-listed field reaches stdout/stderr (-o ${format})`, async () => {
      fakeClock();
      const recs = [record(), record(OWNER, { LogonType: 0 }), record(OWNER, { LogonType: 6, ClientAppId: 'aaaaaaaa-0000-0000-0000-000000000000', AppId: undefined })];
      const r = await run(['audit-trail', 'mailbox', OWNER, '-o', format], { routes: lifecycle({ pages: [recs.slice(0, 2), recs.slice(2)] }) });
      expect(r.exitCode, r.stderr).toBe(0);
      expect(r.stdout).toContain('MailItemsAccessed');
      for (const s of FORBIDDEN) {
        expect(r.stdout, `stdout (${format}) leaks ${s}`).not.toContain(s);
        expect(r.stderr, `stderr (${format}) leaks ${s}`).not.toContain(s);
      }
    });
  }

  it('also holds under --fields naming forbidden paths', async () => {
    fakeClock();
    setFields(['operation', 'auditData', 'auditData.Item.Subject', 'Folders', 'id', 'objectId']);
    try {
      const r = await run(['audit-trail', 'mailbox', OWNER, '-o', 'json'], { routes: lifecycle() });
      expect(r.exitCode).toBe(0);
      expect(JSON.parse(r.stdout).value).toEqual([{ operation: 'MailItemsAccessed' }]);
      for (const s of FORBIDDEN) expect(r.stdout + r.stderr).not.toContain(s);
    } finally { setFields(null); }
  });

  it('projectRecord returns exactly the allow-listed fields, whatever auditData holds', () => {
    for (let i = 0; i < 25; i++) {
      // random extra keys, nested objects and arrays laced with markers
      const extra: Record<string, unknown> = {};
      for (let k = 0; k < 10; k++) extra[`Rand${i}_${k}`] = k % 3 === 0 ? { Subject: mark(`r${i}${k}`) } : k % 3 === 1 ? [mark(`a${i}${k}`)] : mark(`s${i}${k}`);
      const rec = record(OWNER, extra);
      const row = projectRecord(rec.auditData, rec);
      expect(Object.keys(row).sort()).toEqual([...MAILBOX_FIELDS].sort());
      const text = JSON.stringify(row);
      for (const s of FORBIDDEN) expect(text).not.toContain(s);
    }
  });

  it('projectRecord works from auditData alone (CSV import shape) and truncates the client string', () => {
    const row = projectRecord({
      CreationTime: '2026-10-10T09:00:00', Operation: 'FolderBind', UserId: 'd@contoso.com', UserType: 2, ClientIPAddress: '198.51.100.1',
      LogonType: 1, ClientInfoString: 'x'.repeat(300), ClientAppId: 'cid', ExternalAccess: true, Folders: [{}, {}],
      Item: { Subject: SUBJECT_ITEM },
    });
    expect(row).toMatchObject({
      createdDateTime: '2026-10-10T09:00:00', operation: 'FolderBind', actor: 'd@contoso.com', userType: 'admin',
      logonType: 'admin', clientIp: '198.51.100.1', appId: 'cid', externalAccess: true, folderCount: 2, operationCount: 1,
    });
    expect(String(row.client).length).toBe(120);
    expect(JSON.stringify(row)).not.toContain('SECRET');
  });

  it('accessFacts adds the mailbox owner and an item count for aggregation, never item content', () => {
    const rec = record();
    const facts = accessFacts(rec.auditData, rec);
    expect(facts.mailboxOwner).toBe(OWNER.toLowerCase());
    expect(facts.itemCount).toBe(1);
    expect(facts.logonTypeCode).toBe(2);
    for (const s of FORBIDDEN) expect(JSON.stringify(facts)).not.toContain(s);
  });
});

describe('audit-trail mailbox: --dry-run and --read-only (AO-13, AO-14)', () => {
  it('--dry-run traces the POST with its body and does not poll', async () => {
    const r = await run(['audit-trail', 'mailbox', OWNER], { dryRun: true });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toContain(`[dry-run] POST https://graph.microsoft.com/v1.0${QUERIES}`);
    expect(r.stderr).toMatch(/\[dry-run\] Body: .*"keywordFilter":"Alex@Contoso.com"/);
    expect(r.stderr).not.toMatch(/\[dry-run\] GET/);
    expect(r.stderr).toMatch(/\[dry-run\] 1 request\(s\) previewed/);
  });

  it('--read-only blocks the POST with exit 4 and a clear message', async () => {
    const r = await run(['audit-trail', 'mailbox', OWNER], { readOnly: true });
    expect(r.exitCode).toBe(4);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toMatch(/read-only/i);
    expect(r.stderr).toMatch(/audit search query/);
  });
});

describe('audit-trail 403 hints (AO-16)', () => {
  const forbidden = (message: string, code = 'Forbidden') => () => jsonResponse({ error: { code, message } }, 403);

  it('Purview role missing → Audit Reader hint (exit 4)', async () => {
    const r = await run(['audit-trail', 'mailbox', OWNER], { routes: [['POST', QUERIES, forbidden('User: 1234 dont have any permissions')]] });
    expect(r.exitCode).toBe(4);
    expect(r.stderr).toMatch(/Audit Reader/);
    expect(r.stderr).not.toMatch(/grant-scopes/);
  });

  it('Graph consent missing → grant-scopes hint naming the scope (exit 4)', async () => {
    const r = await run(['audit-trail', 'mailbox', OWNER], {
      routes: [['POST', QUERIES, forbidden("Missing scope permissions on the request. API requires one of 'AuditLogsQuery-Exchange.Read.All'.")]],
    });
    expect(r.exitCode).toBe(4);
    expect(r.stderr).toMatch(/AuditLogsQuery-Exchange\.Read\.All/);
    expect(r.stderr).toMatch(/grant-scopes/);
    expect(r.stderr).not.toMatch(/Audit Reader/);
  });

  it('changes: missing Entra role → Reports Reader / Security Reader hint (exit 4)', async () => {
    const r = await run(['audit-trail', 'changes'], {
      routes: [['GET', '/auditLogs/directoryAudits', forbidden('Insufficient privileges to complete the operation.', 'Authorization_RequestDenied')]],
    });
    expect(r.exitCode).toBe(4);
    expect(r.stderr).toMatch(/Security Reader/);
    expect(r.stderr).not.toMatch(/Audit Reader/);
  });

  it('changes: missing consent → grant-scopes hint naming AuditLog.Read.All', async () => {
    const r = await run(['audit-trail', 'changes'], {
      routes: [['GET', '/auditLogs/directoryAudits', forbidden("Missing scope permissions on the request. API requires one of 'AuditLog.Read.All'.")]],
    });
    expect(r.exitCode).toBe(4);
    expect(r.stderr).toMatch(/AuditLog\.Read\.All/);
    expect(r.stderr).toMatch(/grant-scopes/);
  });
});

describe('audit-trail changes (AO-11)', () => {
  const audits = [
    { id: 'a1', activityDateTime: '2026-10-12T10:00:00Z', activityDisplayName: 'Add member to role', category: 'RoleManagement', result: 'success',
      initiatedBy: { user: { userPrincipalName: 'admin@contoso.com', id: mark('initiatorId'), ipAddress: mark('ip') } },
      targetResources: [{ displayName: 'Global Administrator', type: 'Role', id: mark('roleId'), modifiedProperties: [{ newValue: mark('newValue') }] },
        { userPrincipalName: 'new.admin@contoso.com', type: 'User', id: mark('userId') }],
      additionalDetails: [{ key: mark('k'), value: mark('v') }], correlationId: mark('corr') },
    { id: 'a2', activityDateTime: '2026-10-11T10:00:00Z', activityDisplayName: 'Consent to application', category: 'ApplicationManagement', result: 'success',
      initiatedBy: { app: { displayName: 'Some Connector', appId: mark('appId') } },
      targetResources: [{ displayName: 'Some Connector', type: 'ServicePrincipal' }] },
    { id: 'a3', activityDateTime: '2026-10-10T10:00:00Z', activityDisplayName: 'Update application', category: 'ApplicationManagement', result: 'success',
      initiatedBy: {}, targetResources: [] },
  ];

  it('filters directoryAudits by time and category server-side, by activity client-side, and projects the allow-list', async () => {
    const r = await run(['audit-trail', 'changes', '--since', '2026-10-01T00:00Z', '-o', 'json'],
      { routes: [['GET', '/auditLogs/directoryAudits', () => ({ value: audits })]] });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.calls).toHaveLength(1);
    const q = r.calls[0].path;
    expect(q).toContain("activityDateTime ge 2026-10-01T00:00:00.000Z");
    expect(q).toContain("(category eq 'RoleManagement' or category eq 'ApplicationManagement')");
    expect(q).toContain('$orderby=activityDateTime desc');
    const out = JSON.parse(r.stdout);
    expect(out.value).toEqual([
      { activityDateTime: '2026-10-12T10:00:00Z', activityDisplayName: 'Add member to role', result: 'success', initiatedBy: 'admin@contoso.com',
        targets: 'Global Administrator (Role); new.admin@contoso.com (User)' },
      { activityDateTime: '2026-10-11T10:00:00Z', activityDisplayName: 'Consent to application', result: 'success', initiatedBy: 'Some Connector',
        targets: 'Some Connector (ServicePrincipal)' },
    ]);
    expect(r.stdout).not.toContain('ZZLEAK_');
  });

  it('--until adds an upper bound', async () => {
    const r = await run(['audit-trail', 'changes', '--since', '2026-10-01T00:00Z', '--until', '2026-10-05T00:00Z', '-o', 'json'],
      { routes: [['GET', '/auditLogs/directoryAudits', () => ({ value: [] })]] });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.calls[0].path).toContain('activityDateTime le 2026-10-05T00:00:00.000Z');
  });

  it('pages through printList (nextCursor surfaced)', async () => {
    const r = await run(['audit-trail', 'changes', '-o', 'json'], {
      routes: [['GET', '/auditLogs/directoryAudits', () => ({ value: audits.slice(0, 1), '@odata.nextLink': 'https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?$skiptoken=x' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(JSON.parse(r.stdout)).toHaveProperty('nextCursor');
  });

  it('is a GET: allowed under --read-only', async () => {
    const r = await run(['audit-trail', 'changes', '-o', 'json'], { readOnly: true, routes: [['GET', '/auditLogs/directoryAudits', () => ({ value: [] })]] });
    expect(r.exitCode, r.stderr).toBe(0);
  });
});


// --- AOT-5: apps (AO-10, AO-28, AO-33) ---

const HOME_TID = 'home-tenant-0000';
const json404 = () => jsonResponse({ error: { code: 'Request_ResourceNotFound', message: 'not found' } }, 404);

interface AppsFixture {
  grantsFilter400?: boolean;
  exo404?: boolean;
}

/** A tenant with: one app holding Mail.Read as an application permission, an external connector whose only broad
 * access is an AllPrincipals Mail.Read.Shared grant, an EWS grant on Exchange Online, and grants that must not count. */
function appsRoutes({ grantsFilter400 = false, exo404 = false }: AppsFixture = {}): Route[] {
  const sps: Record<string, Record<string, unknown>> = {
    'sp-backup': { id: 'sp-backup', appId: 'backup-app', displayName: 'Backup App', appOwnerOrganizationId: HOME_TID, verifiedPublisher: {},
      appRoleAssignmentRequired: true, disabledByMicrosoftStatus: null },
    'sp-conn': { id: 'sp-conn', appId: 'conn-app', displayName: 'Connector', appOwnerOrganizationId: 'vendor-tenant',
      verifiedPublisher: { displayName: 'Acme Inc', verifiedPublisherId: '123' }, appRoleAssignmentRequired: false, disabledByMicrosoftStatus: 'NotDisabled' },
    'sp-ews': { id: 'sp-ews', appId: 'ews-app', displayName: 'Old EWS Sync', appOwnerOrganizationId: 'vendor2', verifiedPublisher: {}, appRoleAssignmentRequired: false },
  };
  const grants = [
    { id: 'g1', clientId: 'sp-conn', consentType: 'AllPrincipals', principalId: null, resourceId: 'graph-sp', scope: 'openid User.Read Mail.Read.Shared offline_access' },
    { id: 'g2', clientId: 'sp-basic', consentType: 'AllPrincipals', principalId: null, resourceId: 'graph-sp', scope: 'Mail.ReadBasic User.Read' },
    { id: 'g3', clientId: 'sp-peruser', consentType: 'Principal', principalId: 'user-1', resourceId: 'graph-sp', scope: 'Mail.Read' },
    { id: 'g4', clientId: 'sp-substr', consentType: 'AllPrincipals', principalId: null, resourceId: 'graph-sp', scope: 'Mail.Read.SharedX Files.Read' },
    { id: 'g5', clientId: 'sp-ews', consentType: 'AllPrincipals', principalId: null, resourceId: 'exo-sp', scope: 'EWS.AccessAsUser.All' },
    { id: 'g6', clientId: 'sp-other-resource', consentType: 'AllPrincipals', principalId: null, resourceId: 'sharepoint-sp', scope: 'Mail.Read' },
  ];
  return [['GET', /.*/, (call) => {
    const p = call.path;
    if (p.startsWith(`/servicePrincipals(appId='${GRAPH_APP_ID}')`)) {
      return { id: 'graph-sp', appRoles: [
        { id: 'role-mail-read', value: 'Mail.Read' }, { id: 'role-user-read', value: 'User.Read.All' }, { id: 'role-chat', value: 'Chat.Read.All' },
      ] };
    }
    if (p.startsWith(`/servicePrincipals(appId='${EXCHANGE_APP_ID}')`)) return exo404 ? json404() : { id: 'exo-sp' };
    if (p.startsWith('/servicePrincipals/graph-sp/appRoleAssignedTo')) {
      return { value: [
        { principalId: 'sp-backup', principalType: 'ServicePrincipal', principalDisplayName: 'Backup App', appRoleId: 'role-mail-read', createdDateTime: '2025-03-01T00:00:00Z' },
        { principalId: 'sp-dir', principalType: 'ServicePrincipal', principalDisplayName: 'Directory Sync', appRoleId: 'role-user-read', createdDateTime: '2025-01-01T00:00:00Z' },
      ] };
    }
    if (p.startsWith('/oauth2PermissionGrants')) {
      if (p.includes('$filter')) {
        if (grantsFilter400) return jsonResponse({ error: { code: 'BadRequest', message: 'Unsupported filter' } }, 400);
        return { value: grants.filter(g => g.consentType === 'AllPrincipals') };
      }
      return { value: grants };
    }
    const assigned = /^\/servicePrincipals\/([^/?]+)\/appRoleAssignedTo/.exec(p);
    if (assigned) return { value: assigned[1] === 'sp-backup' ? [{ id: 'a' }, { id: 'b' }] : [] };
    const sp = /^\/servicePrincipals\/([^/?]+)\?/.exec(p);
    if (sp) return sps[sp[1]] ?? json404();
    return json404();
  }]];
}

describe('audit-trail apps (AO-10, AO-28, AO-33, AC 12, AC 18)', () => {
  const runApps = (extra: string[] = [], fx: AppsFixture = {}, o: Parameters<typeof runCommand>[2] = {}) => {
    vi.spyOn(tenant, 'id').mockResolvedValue(HOME_TID);
    return run(['audit-trail', 'apps', '-o', 'json', ...extra], { routes: appsRoutes(fx), ...o });
  };

  it('lists an app whose only broad access is an AllPrincipals Mail.Read.Shared grant: external, verified publisher (AC 12)', async () => {
    const r = await runApps();
    expect(r.exitCode, r.stderr).toBe(0);
    const rows = JSON.parse(r.stdout).value;
    const conn = rows.find((x: { appId: string }) => x.appId === 'conn-app');
    expect(conn).toEqual({
      displayName: 'Connector', appId: 'conn-app', servicePrincipalId: 'sp-conn', publisher: 'Acme Inc', origin: 'external',
      kind: 'delegated-all-users', scopes: 'Mail.Read.Shared', appRoleAssignmentRequired: false, disabledByMicrosoftStatus: 'NotDisabled',
      assignedPrincipals: 0, granted: null,
    });
    for (const row of rows) expect(Object.keys(row).sort()).toEqual([...APP_FIELDS].sort());
  });

  it('resolves application permissions by appRoles[].value, not hard-coded GUIDs; this-tenant + unverified shown', async () => {
    const r = await runApps();
    const rows = JSON.parse(r.stdout).value;
    expect(rows.find((x: { appId: string }) => x.appId === 'backup-app')).toMatchObject({
      kind: 'application', scopes: 'Mail.Read', origin: 'this tenant', publisher: 'unverified', appRoleAssignmentRequired: true,
      assignedPrincipals: 2, granted: '2025-03-01T00:00:00Z',
    });
    expect(rows.some((x: { servicePrincipalId: string }) => x.servicePrincipalId === 'sp-dir')).toBe(false); // User.Read.All is not broad
  });

  it('per-user consent, Mail.ReadBasic, substring look-alikes and other resources are not listed', async () => {
    const r = await runApps();
    const ids = JSON.parse(r.stdout).value.map((x: { servicePrincipalId: string }) => x.servicePrincipalId).sort();
    expect(ids).toEqual(['sp-backup', 'sp-conn', 'sp-ews']);
  });

  it('lists AllPrincipals EWS/IMAP/POP grants on Exchange Online (AO-33)', async () => {
    const r = await runApps();
    const ews = JSON.parse(r.stdout).value.find((x: { appId: string }) => x.appId === 'ews-app');
    expect(ews).toMatchObject({ kind: 'delegated-all-users (Exchange Online)', scopes: 'EWS.AccessAsUser.All', origin: 'external' });
  });

  it('skips Exchange Online grants when that service principal is absent (404)', async () => {
    const r = await runApps([], { exo404: true });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(JSON.parse(r.stdout).value.map((x: { appId: string }) => x.appId).sort()).toEqual(['backup-app', 'conn-app']);
  });

  it('falls back to the unfiltered grant list on a 400 for the consentType filter and still yields the AC 12 result (AC 18)', async () => {
    const r = await runApps([], { grantsFilter400: true });
    expect(r.exitCode, r.stderr).toBe(0);
    const grantCalls = r.calls.filter(c => c.path.startsWith('/oauth2PermissionGrants'));
    expect(grantCalls).toHaveLength(2);
    expect(grantCalls[1].path).not.toContain('$filter');
    const rows = JSON.parse(r.stdout).value;
    expect(rows.map((x: { servicePrincipalId: string }) => x.servicePrincipalId).sort()).toEqual(['sp-backup', 'sp-conn', 'sp-ews']);
    expect(rows.find((x: { appId: string }) => x.appId === 'conn-app')).toMatchObject({ origin: 'external', publisher: 'Acme Inc' });
  });

  it('origin is "unknown" when the token tenant cannot be read; no risk score field (E3)', async () => {
    const r = await run(['audit-trail', 'apps', '-o', 'json'], { routes: appsRoutes() }); // test token is not a JWT
    expect(r.exitCode, r.stderr).toBe(0);
    const rows = JSON.parse(r.stdout).value;
    expect(rows.every((x: { origin: string }) => x.origin === 'unknown')).toBe(true);
    expect(r.stdout).not.toMatch(/risk|score/i);
  });

  it('every request is a GET: allowed under --read-only (AC 12)', async () => {
    const r = await runApps([], {}, { readOnly: true });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.calls.every(c => c.method === 'GET')).toBe(true);
    expect(JSON.parse(r.stdout).value).toHaveLength(3);
  });

  it('--dry-run sends nothing and traces the first GET only (AC 12)', async () => {
    const r = await runApps([], {}, { dryRun: true });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr.match(/\[dry-run\] GET/g)).toHaveLength(1);
    expect(r.stderr).toContain(`servicePrincipals(appId='${GRAPH_APP_ID}')`);
  });

  it('a 403 maps to exit 4 with a directory-role / consent hint', async () => {
    const r = await run(['audit-trail', 'apps'], { routes: [['GET', /.*/, () => jsonResponse({ error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges to complete the operation.' } }, 403)]] });
    expect(r.exitCode).toBe(4);
    expect(r.stderr).toMatch(/Directory Readers|Global Reader/);
  });

  it('catalog: apps is read (Directory.Read.All), access is an action like mailbox', () => {
    expect(commandMeta('audit-trail apps')).toMatchObject({ readOnly: true, scopes: ['Directory.Read.All'] });
    expect(commandMeta('audit-trail access')).toMatchObject({ readOnly: false, scopes: ['AuditLogsQuery-Exchange.Read.All'] });
  });
});

// --- AOT-5: access (AO-29, AO-30) ---

const APP = '11111111-2222-3333-4444-555555555555'; // the AppId leakyAuditData carries
const ACTOR = 'user@contoso.com';

/** A record by `actor` on `owner`'s mailbox (envelope userPrincipalName = actor). */
function accessRecord(actor: string, owner: string, extra: Record<string, unknown> = {}, createdDateTime = '2026-10-10T09:00:00Z') {
  return { ...record(owner, extra), userPrincipalName: actor, createdDateTime };
}

/** The worked example: one actor, own mailbox + three delegated mailboxes (one a disabled admin account), one other app. */
function workedExample() {
  const items = (n: number) => ({ Folders: [{ Path: mark('p'), FolderItems: Array.from({ length: n }, (_, i) => ({ InternetMessageId: `${IMID}${i}` })) }] });
  return [
    accessRecord(ACTOR, ACTOR, { LogonType: 0 }),
    accessRecord(ACTOR, 'USER@contoso.com', { LogonType: 0 }),
    accessRecord(ACTOR, ACTOR, { LogonType: 2 }), // owner = actor, even if logged as delegated
    accessRecord(ACTOR, 'shared1@contoso.com', items(2), '2026-10-05T08:00:00Z'),
    accessRecord(ACTOR, 'Shared1@Contoso.com', items(3), '2026-10-07T08:00:00Z'),
    accessRecord(ACTOR, 'shared1@contoso.com', { ...items(1), ClientInfoString: 'Client=OWA' }, '2026-10-06T08:00:00Z'),
    accessRecord(ACTOR, 'team@contoso.com', items(4), '2026-10-08T08:00:00Z'),
    accessRecord(ACTOR, 'ck-admin@contoso.com', items(5), '2026-10-09T08:00:00Z'),
    accessRecord(ACTOR, 'ck-admin@contoso.com', { ...items(1), AppId: 'ffffffff-0000-0000-0000-000000000000' }, '2026-10-09T09:00:00Z'), // another app
  ];
}

type UserAnswer = Record<string, unknown> | (() => Response);
function accessRoutes(pages: unknown[][], users: Record<string, UserAnswer> = {}): Route[] {
  return [
    ['GET', /^\/users\//, (call) => {
      const upn = decodeURIComponent(/^\/users\/([^?]+)/.exec(call.path)![1]);
      const u = users[upn] ?? { accountEnabled: true, assignedLicenses: [{ skuId: 'e3' }] };
      return typeof u === 'function' ? u() : u;
    }],
    ...lifecycle({ pages }),
  ];
}
const DEFAULT_USERS: Record<string, UserAnswer> = { 'ck-admin@contoso.com': { accountEnabled: false, assignedLicenses: [] } };

describe('audit-trail access (AO-29, AO-30, AC 13, AC 17)', () => {
  it('without --app or --by exits 2 and sends no request (AC 13)', async () => {
    const r = await run(['audit-trail', 'access']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toMatch(/--app .* or --by/);
  });

  it.each([
    [['--by', 'a@contoso.com,b@contoso.com']], [['--by', '*@contoso.com']], [['--by', 'a@contoso.com', '--by', 'b@contoso.com']],
    [['--app', 'not-a-guid']], [['--app', APP, '--app', APP]],
  ])('%j exits 2 with no request', async (args) => {
    const r = await run(['audit-trail', 'access', ...args]);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('--app: keywordFilter is the appId, default window 30 days, same record types and operations', async () => {
    fakeClock();
    const r = await run(['audit-trail', 'access', '--app', APP.toUpperCase(), '-o', 'json'], { routes: accessRoutes([[]]) });
    expect(r.exitCode, r.stderr).toBe(0);
    const body = r.calls.find(c => c.method === 'POST')!.body as Record<string, string>;
    expect(body).toMatchObject({ keywordFilter: APP, recordTypeFilters: ['exchangeItem', 'exchangeItemGroup', 'exchangeItemAggregated'], operationFilters: [...DEFAULT_OPERATIONS] });
    expect(body).not.toHaveProperty('userPrincipalNameFilters');
    const days = (Date.parse(body.filterEndDateTime) - Date.parse(body.filterStartDateTime)) / 86_400_000;
    expect(Math.round(days)).toBe(30);
  });

  it('--by: userPrincipalNameFilters is the actor, no keywordFilter', async () => {
    fakeClock();
    const r = await run(['audit-trail', 'access', '--by', ACTOR, '-o', 'json'], { routes: accessRoutes([[]]) });
    expect(r.exitCode, r.stderr).toBe(0);
    const body = r.calls.find(c => c.method === 'POST')!.body;
    expect(body).toMatchObject({ userPrincipalNameFilters: [ACTOR] });
    expect(body).not.toHaveProperty('keywordFilter');
  });

  it('aggregates the worked example per (actor, mailbox, logonType); own mailbox omitted and counted; disabled target flagged', async () => {
    fakeClock();
    const recs = workedExample();
    const r = await run(['audit-trail', 'access', '--app', APP, '-o', 'json'],
      { routes: accessRoutes([recs.slice(0, 4), recs.slice(4)], DEFAULT_USERS) });
    expect(r.exitCode, r.stderr).toBe(0);
    const rows = JSON.parse(r.stdout).value;
    expect(rows).toEqual([
      { actor: ACTOR, mailbox: 'ck-admin@contoso.com', logonType: 'delegated', events: 1, items: 5, first: '2026-10-09T08:00:00Z', last: '2026-10-09T08:00:00Z',
        client: 'Client=REST;Client=RESTSystem;;', appId: APP, target: 'disabled-or-shared' },
      { actor: ACTOR, mailbox: 'shared1@contoso.com', logonType: 'delegated', events: 3, items: 6, first: '2026-10-05T08:00:00Z', last: '2026-10-07T08:00:00Z',
        client: 'Client=REST;Client=RESTSystem;;; Client=OWA', appId: APP, target: null },
      { actor: ACTOR, mailbox: 'team@contoso.com', logonType: 'delegated', events: 1, items: 4, first: '2026-10-08T08:00:00Z', last: '2026-10-08T08:00:00Z',
        client: 'Client=REST;Client=RESTSystem;;', appId: APP, target: null },
    ]);
    for (const row of rows) expect(Object.keys(row).sort()).toEqual([...ACCESS_FIELDS].sort());
    expect(rows.some((x: { mailbox: string }) => x.mailbox === ACTOR)).toBe(false);
    expect(r.stderr).toMatch(/Coverage: .*3 own-mailbox events omitted\./);
    expect(r.stderr).toMatch(/an empty result is not proof of no access/);
    expect(r.stderr).toMatch(/disabled-or-shared: disabled, unlicensed account: a shared mailbox or a retired account/);
    // one annotation GET per distinct mailbox, after aggregation (AO-30 cache)
    expect(r.calls.filter(c => c.path.startsWith('/users/')).map(c => c.path.split('?')[0]).sort())
      .toEqual(['/users/ck-admin@contoso.com', '/users/shared1@contoso.com', '/users/team@contoso.com']);
  });

  it('--by keeps every app of that actor and drops other actors client-side', async () => {
    fakeClock();
    const recs = [...workedExample(), accessRecord('other@contoso.com', 'team@contoso.com')];
    const r = await run(['audit-trail', 'access', '--by', ACTOR, '-o', 'json'], { routes: accessRoutes([recs], DEFAULT_USERS) });
    expect(r.exitCode, r.stderr).toBe(0);
    const rows = JSON.parse(r.stdout).value;
    expect(rows.every((x: { actor: string }) => x.actor === ACTOR)).toBe(true);
    expect(rows.find((x: { mailbox: string }) => x.mailbox === 'ck-admin@contoso.com')).toMatchObject({ events: 2, items: 6 });
  });

  it('annotation: 404 → not-found; 403 → unflagged with one stderr warning, never fatal (AC 17)', async () => {
    fakeClock();
    const recs = workedExample();
    const forbidden = () => jsonResponse({ error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges' } }, 403);
    let r = await run(['audit-trail', 'access', '--app', APP, '-o', 'json'], { routes: accessRoutes([recs], { 'team@contoso.com': json404 }) });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(JSON.parse(r.stdout).value.find((x: { mailbox: string }) => x.mailbox === 'team@contoso.com').target).toBe('not-found');

    fakeClock();
    r = await run(['audit-trail', 'access', '--app', APP, '-o', 'json'],
      { routes: accessRoutes([recs], { 'ck-admin@contoso.com': forbidden, 'shared1@contoso.com': forbidden, 'team@contoso.com': forbidden }) });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(JSON.parse(r.stdout).value.every((x: { target: unknown }) => x.target === null)).toBe(true);
    expect(r.stderr.match(/annotation skipped/g)).toHaveLength(1);
  });

  it('warns when the window is longer than 180 days', async () => {
    fakeClock();
    const r = await run(['audit-trail', 'access', '--by', ACTOR, '--since', '-200d', '-o', 'json'], { routes: accessRoutes([[]]) });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.stderr).toMatch(/spans 200 days/);
  });

  it('--dry-run traces the POST, does not poll or annotate, and prints the coverage line', async () => {
    const r = await run(['audit-trail', 'access', '--app', APP], { dryRun: true });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toContain(`[dry-run] POST https://graph.microsoft.com/v1.0${QUERIES}`);
    expect(r.stderr).not.toMatch(/\[dry-run\] GET/);
    expect(r.stderr).toMatch(/Coverage: /);
  });

  it('--read-only blocks the POST with exit 4', async () => {
    const r = await run(['audit-trail', 'access', '--by', ACTOR], { readOnly: true });
    expect(r.exitCode).toBe(4);
    expect(r.calls).toHaveLength(0);
    expect(r.stderr).toMatch(/audit search query/);
  });

  for (const format of ['json', 'ndjson', 'table', 'csv', 'text', 'yaml']) {
    it(`allow-list: no subject, InternetMessageId, recipient or non-allow-listed field in access output (-o ${format}, AC 7)`, async () => {
      fakeClock();
      const r = await run(['audit-trail', 'access', '--app', APP, '-o', format], { routes: accessRoutes([workedExample()], DEFAULT_USERS) });
      expect(r.exitCode, r.stderr).toBe(0);
      expect(r.stdout).toContain('shared1@contoso.com');
      for (const s of FORBIDDEN) {
        expect(r.stdout, `stdout (${format}) leaks ${s}`).not.toContain(s);
        expect(r.stderr, `stderr (${format}) leaks ${s}`).not.toContain(s);
      }
    });
  }
});

describe('audit-trail mailbox: target annotation (AO-30)', () => {
  it('flags a disabled, unlicensed mailbox on stderr', async () => {
    fakeClock();
    const r = await run(['audit-trail', 'mailbox', 'ck-admin@contoso.com', '-o', 'json'],
      { routes: [['GET', /^\/users\//, () => ({ accountEnabled: false, assignedLicenses: [] })], ...lifecycle({ pages: [[record('ck-admin@contoso.com')]] })] });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.stderr).toMatch(/disabled-or-shared \(disabled, unlicensed account: a shared mailbox or a retired account\)/);
    expect(JSON.parse(r.stdout).value).toHaveLength(1);
  });
});

describe('audit-trail: calling app vs resource app (real Purview record shape)', () => {
  // Real MailItemsAccessed records carry the RESOURCE in AppId (Microsoft Graph) and the calling
  // client in ClientAppId / AppAccessContext.ClientAppId. Found on a live export 2026-10-17.
  const GRAPH = '00000003-0000-0000-c000-000000000000';
  const CLIENT = '07c030f6-5743-41b7-ba00-0a6e85f37c17';
  const real = {
    Operation: 'MailItemsAccessed', UserId: 'pat@contoso.com', MailboxOwnerUPN: 'ops@contoso.com', LogonType: 2,
    AppId: GRAPH, ClientAppId: CLIENT, AppAccessContext: { ClientAppId: CLIENT, APIId: GRAPH },
  };
  it('projectRecord reports the calling app, not the resource', () => {
    expect(projectRecord(real).appId).toBe(CLIENT);
  });
  it('falls back to AppAccessContext.ClientAppId, then AppId', () => {
    expect(projectRecord({ ...real, ClientAppId: undefined }).appId).toBe(CLIENT);
    expect(projectRecord({ AppId: 'aaaaaaaa-0000-4000-8000-000000000001' }).appId).toBe('aaaaaaaa-0000-4000-8000-000000000001');
  });
  it('accessFacts keeps the calling app so --app filters match', () => {
    expect(accessFacts(real).appId).toBe(CLIENT);
  });
});

describe('audit-trail changes: 30-day directory audit window (live 2026-10-17)', () => {
  // Graph rejects activityDateTime earlier than 30 days ago with a 400, so --since is clamped with a note.
  it('clamps an older --since to 30 days ago and says so on stderr', async () => {
    // The clock is frozen: the command and this assertion each read "now", and a real clock moves between the two
    // reads (CI saw ageDays 30.0000000116 > 30). With one instant the clamp is exactly 30 days.
    const now = Date.parse('2026-10-17T12:00:00.000Z');
    vi.useFakeTimers({ toFake: ['Date'], now });
    try {
      const r = await runCommand(registerAuditTrailCommands, ['audit-trail', 'changes', '--since', '-60d', '-o', 'json'], {
        routes: [['GET', /directoryAudits/, () => ({ value: [] })]],
      });
      expect(r.exitCode, r.stderr).toBe(0);
      const ge = r.calls[0].path.match(/activityDateTime ge (\S+Z)/)![1];
      expect(ge).toBe(new Date(now - 30 * 86_400_000).toISOString());
      expect(r.stderr).toMatch(/30 days/);
    } finally {
      vi.useRealTimers();
    }
  });
  it('leaves a --since inside the window alone, with no note', async () => {
    const { registerAuditTrailCommands: reg } = await import('../../src/commands/audit-trail.js');
    const r = await runCommand(reg, ['audit-trail', 'changes', '--since', '-5d', '-o', 'json'], {
      routes: [['GET', /directoryAudits/, () => ({ value: [] })]],
    });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.stderr).not.toMatch(/30 days/);
  });
});
