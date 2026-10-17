import { describe, it, expect, vi, afterEach } from 'vitest';
import { runCommand, jsonResponse, type Route } from '../helpers/cli.js';
import { setFields } from '../../src/formatter.js';
import {
  registerAuditTrailCommands, projectRecord, accessFacts, timing, MAILBOX_FIELDS, DEFAULT_OPERATIONS,
} from '../../src/commands/audit-trail.js';

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
