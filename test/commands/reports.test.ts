import { describe, it, expect } from 'vitest';
import { runCommand } from '../helpers/cli.js';
import { registerReportsCommands, rollUp, MIN_GROUP_SIZE } from '../../src/commands/reports.js';

const people = (dept: string, n: number, prefix = dept.toLowerCase()) =>
  Array.from({ length: n }, (_, i) => ({ upn: `${prefix}${i}@x.com`, dept }));

function fixture(groups: Array<[string, number]>) {
  const users = groups.flatMap(([d, n]) => people(d, n));
  const deptOf = new Map(users.map(u => [u.upn, u.dept]));
  const details = users.map((u, i) => ({ 'User Principal Name': u.upn, 'Send Count': String(i + 1), 'Receive Count': '10' }));
  return { users, deptOf, details };
}

describe('rollUp', () => {
  it('sums per department and never returns per-person values', () => {
    const { deptOf, details } = fixture([['Sales', 6], ['Eng', 8]]);
    const { rows } = rollUp(details, deptOf, ['sendcount', 'receivecount']);
    expect(rows.map(r => [r.department, r.people])).toEqual([['Eng', 8], ['Sales', 6]]);
    expect(rows.find(r => r.department === 'Eng')!.receivecount).toBe(80);
    expect(JSON.stringify(rows)).not.toMatch(/@x\.com/);
  });

  it('pools groups under the minimum into Other, and drops a pool that is itself too small', () => {
    const a = fixture([['Sales', 6], ['Legal', 2], ['HR', 2], ['Board', 1]]);
    const pooled = rollUp(a.details, a.deptOf, ['sendcount']);
    expect(pooled.rows.map(r => [r.department, r.people])).toEqual([['Sales', 6], [`Other (groups under ${MIN_GROUP_SIZE} people)`, 5]]);
    const b = fixture([['Sales', 6], ['Legal', 2], ['HR', 2]]);
    const dropped = rollUp(b.details, b.deptOf, ['sendcount']);
    expect(dropped.rows.map(r => r.department)).toEqual(['Sales']); // pool of 4 is hidden
    for (const row of [...pooled.rows, ...dropped.rows]) expect(row.people).toBeGreaterThanOrEqual(MIN_GROUP_SIZE);
  });

  it('counts unmatched accounts and ignores concealed (hashed) names', () => {
    const { deptOf, details } = fixture([['Sales', 7]]);
    const extra = [...details, { 'User Principal Name': 'nodept@x.com' }, { 'User Principal Name': 'A1B2C3D4E5F6A7B8C9D0E1F2' }];
    const r = rollUp(extra, deptOf, ['sendcount'], 5);
    expect(r.unmatched).toBe(2);
    expect(r.rows[0].people).toBe(7);
    expect(r).not.toHaveProperty('excluded');
  });

  it('AO-22: Sales=5, Eng=7, Legal=2, HR=2 never exposes Legal+HR by differencing two minimums', () => {
    const { deptOf, details } = fixture([['Sales', 5], ['Eng', 7], ['Legal', 2], ['HR', 2]]);
    const at = (min: number) => rollUp(details, deptOf, ['sendcount'], min).rows.map(r => [r.department, r.people]);
    expect(at(5)).toEqual([['Eng', 7], ['Sales', 5]]);
    // At 6 the pool is Sales only (5): under the minimum, so omitted (AC 2). Legal+HR (4) never enter it, so
    // the old leak "Other(9)@6 − Sales@5 = Legal+HR" cannot happen.
    expect(at(6)).toEqual([['Eng', 7]]);
    expect(at(8)).toEqual([['Other (groups under 8 people)', 12]]);
  });

  it('AO-22 property: no difference between runs with different --min-group-size describes < 5 people', () => {
    // Each department gets one "marker" person whose metric is 2^index, so a row's metric sum is the exact
    // set of departments it covers. Any set an observer can derive (row, or difference of nested rows,
    // or union of disjoint rows) across all runs must cover at least the floor.
    let seed = 42;
    const rand = (n: number) => { seed = (seed * 1103515245 + 12345) % 2 ** 31; return seed % n; };
    const configs: Array<Array<[string, number]>> = [
      [['Sales', 5], ['Eng', 7], ['Legal', 2], ['HR', 2]],
      [['A', 3], ['B', 2], ['C', 6], ['D', 9], ['E', 1]],
      [['A', 4], ['B', 4], ['C', 5], ['D', 5], ['E', 11]],
      [['A', 1], ['B', 1], ['C', 1], ['D', 8]],
    ];
    for (let k = 0; k < 40; k++) {
      const n = 2 + rand(8);
      configs.push(Array.from({ length: n }, (_, i) => [`D${i}`, 1 + rand(14)] as [string, number]));
    }
    const mins = [5, 6, 7, 8, 9, 10, 12, 15, 20];
    for (const groups of configs) {
      const deptOf = new Map<string, string>();
      const details: Array<Record<string, unknown>> = [];
      const size = new Map<number, number>();
      groups.forEach(([d, people], idx) => {
        size.set(idx, people);
        for (let i = 0; i < people; i++) {
          const upn = `${d.toLowerCase()}${i}@x.com`;
          deptOf.set(upn, d);
          details.push({ userPrincipalName: upn, marker: i === 0 ? 2 ** idx : 0 });
        }
      });
      const peopleOf = (mask: number) => { let t = 0; size.forEach((p, idx) => { if (mask & (2 ** idx)) t += p; }); return t; };
      const sets = new Set<number>();
      for (const min of mins) {
        const run = rollUp(details, deptOf, ['marker'], min).rows;
        for (const r of run) {
          expect(r.people).toBeGreaterThanOrEqual(min);
          expect(peopleOf(r.marker as number)).toBe(r.people);
          sets.add(r.marker as number);
        }
      }
      // Close the observed sets under nested difference and disjoint union.
      let grew = true;
      while (grew) {
        grew = false;
        for (const a of [...sets]) for (const b of [...sets]) {
          const derived = [(a & b) === b && a !== b ? a & ~b : 0, (a & b) === 0 ? a | b : 0];
          for (const d of derived) if (d && !sets.has(d)) { sets.add(d); grew = true; }
        }
      }
      for (const m of sets) expect(peopleOf(m), `${JSON.stringify(groups)} derives a set of ${peopleOf(m)}`).toBeGreaterThanOrEqual(MIN_GROUP_SIZE);
    }
  });
});

describe('ms365 reports', () => {
  const run = (argv: string[], routes: any[] = [], extra = {}) => runCommand(registerReportsCommands, ['reports', ...argv], { routes, ...extra });

  it('tenant-level email counts: JSON format requested, rows normalized, period passed', async () => {
    const r = await run(['email', '--period', 'D7', '-o', 'json'], [
      ['GET', /getEmailActivityCounts\(period='D7'\)/, () => ({ value: [{ 'Report Date': '2026-10-14', 'Send': 12, Receive: 40 }] })],
    ]);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.calls[0].path).toContain('$format=application/json');
    const out = JSON.parse(r.stdout);
    expect(out).toMatchObject({ period: 'D7', report: 'email' });
    expect(out.value[0]).toMatchObject({ reportdate: '2026-10-14', send: 12 });
  });

  it('by-department aggregates, suppresses small groups and prints no user names', async () => {
    const { users, details } = fixture([['Sales', 6], ['Legal', 2]]);
    const r = await run(['by-department', '--workload', 'email', '-o', 'json'], [
      ['GET', /getEmailActivityUserDetail/, () => ({ value: details })],
      ['GET', '/users', () => ({ value: users.map(u => ({ userPrincipalName: u.upn, department: u.dept })) })],
    ]);
    expect(r.exitCode, r.stderr).toBe(0);
    const out = JSON.parse(r.stdout);
    expect(out.value.map((x: any) => x.department)).toEqual(['Sales']); // Legal (2) hidden
    expect(r.stdout).not.toMatch(/@x\.com/);
    expect(r.stderr).toBe('');
  });

  it('refuses a minimum group size below the floor (exit 2, no request)', async () => {
    const r = await run(['by-department', '--min-group-size', '3']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('concealed names → clear error with the admin setting to change', async () => {
    const hashed = Array.from({ length: 6 }, (_, i) => ({ 'User Principal Name': `A1B2C3D4E5F6A7B8C9D0E1F${i}` }));
    const r = await run(['by-department'], [
      ['GET', /getEmailActivityUserDetail/, () => ({ value: hashed })],
      ['GET', '/users', () => ({ value: [{ userPrincipalName: 'a@x.com', department: 'Sales' }] })],
    ]);
    expect(r.exitCode).toBe(1);
    expect(r.stderr).toContain('concealed');
  });

  it('403 maps to exit 4; dry-run previews both requests and prints nothing', async () => {
    const denied = await run(['teams'], [['GET', '/reports', () => new Response('{"error":{"message":"Forbidden"}}', { status: 403 })]]);
    expect(denied.exitCode).toBe(4);
    const dry = await run(['by-department'], [], { dryRun: true });
    expect(dry.exitCode).toBe(0);
    expect(dry.calls).toHaveLength(0);
    expect(dry.stdout).toBe('');
    expect(dry.stderr).toMatch(/getEmailActivityUserDetail/);
  });

  const REPORTS_FIX = /Reports Reader.*Global Reader.*reports email\|teams\|active-users.*auth grant-scopes/;
  const forbidden = () => new Response('{"error":{"code":"Authorization_RequestDenied","message":"Forbidden"}}', { status: 403, headers: { 'content-type': 'application/json' } });

  it('AO-27: a 403 on a count report names the reports roles (exit 4)', async () => {
    const r = await run(['email'], [['GET', '/reports', forbidden]]);
    expect(r.exitCode).toBe(4);
    expect(r.stderr).toMatch(REPORTS_FIX);
  });

  it('AO-27: a 403 on by-department names the reports roles (exit 4)', async () => {
    const r = await run(['by-department'], [['GET', '/reports', forbidden]]);
    expect(r.exitCode).toBe(4);
    expect(r.stderr).toMatch(REPORTS_FIX);
  });

  const csvRedirect = (csv: string, type = 'application/octet-stream') => [
    ['GET', /reports\.office\.com/, () => new Response(csv, { status: 200, headers: { 'content-type': type } })],
    ['GET', '/reports/', () => new Response(null, { status: 302, headers: { location: 'https://reports.office.com/data/download/abc.csv' } })],
  ];

  it('AO-4: a 302 → CSV count report is parsed (BOM, Title Case headers), token not sent off-Microsoft', async () => {
    const csv = '\uFEFFReport Refresh Date,Send,Receive,Read,Meeting Created,Meeting Interacted,Report Date,Report Period\r\n'
      + '2026-10-15,12,40,30,2,3,2026-10-14,7\r\n';
    const r = await run(['email', '--period', 'D7', '-o', 'json'], csvRedirect(csv));
    expect(r.exitCode, r.stderr).toBe(0);
    const download = r.calls.find(c => c.url.includes('reports.office.com'))!;
    expect(download.headers.Authorization).toBeUndefined();
    const out = JSON.parse(r.stdout);
    expect(out.value).toEqual([{ reportrefreshdate: '2026-10-15', send: 12, receive: 40, read: 30, meetingcreated: 2, meetinginteracted: 3, reportdate: '2026-10-14', reportperiod: 7 }]);
  });

  it('AO-4: a 302 → CSV detail report with a quoted "Assigned Products" rolls up correctly', async () => {
    const { users } = fixture([['Sales', 6], ['Eng', 5]]);
    const header = 'Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Last Activity Date,Send Count,Receive Count,Read Count,Meeting Created Count,Meeting Interacted Count,Assigned Products,Report Period';
    const lines = users.map(u => `2026-10-15,${u.upn},"Doe, Jane",False,,2026-10-14,2,10,5,0,1,"OFFICE 365 E3,MICROSOFT 365 E5",30`);
    for (const type of ['application/octet-stream', 'text/csv']) {
      const r = await run(['by-department', '-o', 'json'], [
        ...csvRedirect('\uFEFF' + [header, ...lines].join('\r\n') + '\r\n', type),
        ['GET', '/users', () => ({ value: users.map(u => ({ userPrincipalName: u.upn, department: u.dept, accountEnabled: true })) })],
      ]);
      expect(r.exitCode, r.stderr).toBe(0);
      const out = JSON.parse(r.stdout);
      expect(out.value.map((x: any) => [x.department, x.people, x.sendcount, x.receivecount])).toEqual([['Sales', 6, 12, 60], ['Eng', 5, 10, 50]]);
      expect(r.stdout).not.toMatch(/@x\.com|Doe/);
    }
  });

  it('AO-23: a two-page JSON detail report is summed', async () => {
    const { users, details } = fixture([['Sales', 6], ['Eng', 5]]);
    const r = await run(['by-department', '-o', 'json'], [
      ['GET', /getEmailActivityUserDetail.*\$skiptoken=p2/, () => ({ value: details.slice(4) })],
      ['GET', /getEmailActivityUserDetail/, (c: any) => ({ value: details.slice(0, 4), '@odata.nextLink': `${c.url}&$skiptoken=p2` })],
      ['GET', '/users', () => ({ value: users.map(u => ({ userPrincipalName: u.upn, department: u.dept, accountEnabled: true })) })],
    ]);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.calls.filter(c => c.path.includes('getEmailActivityUserDetail'))).toHaveLength(2);
    const out = JSON.parse(r.stdout);
    expect(out.value.map((x: any) => [x.department, x.people, x.receivecount])).toEqual([['Sales', 6, 60], ['Eng', 5, 50]]);
    expect(r.stderr).toBe('');
  });

  it('AO-23: warns on stderr when the report covers fewer people than the directory', async () => {
    const { users, details } = fixture([['Sales', 6], ['Eng', 5]]);
    const directory = [...users.map(u => ({ userPrincipalName: u.upn, department: u.dept, accountEnabled: true })),
      { userPrincipalName: 'new1@x.com', department: 'Eng', accountEnabled: true },
      { userPrincipalName: 'new2@x.com', department: 'Eng', accountEnabled: true },
      { userPrincipalName: 'gone@x.com', department: 'Eng', accountEnabled: false }];
    const r = await run(['by-department', '-o', 'json'], [
      ['GET', /getEmailActivityUserDetail/, () => ({ value: details })],
      ['GET', '/users', () => ({ value: directory })],
    ]);
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.stderr).toMatch(/2 enabled account\(s\) with a department are not in the usage report/);
    expect(r.stderr).not.toMatch(/@x\.com/);
  });
});
