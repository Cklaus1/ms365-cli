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

  it('counts unmatched accounts, ignores concealed (hashed) names, and honours an exclusion set', () => {
    const { deptOf, details } = fixture([['Sales', 7]]);
    const extra = [...details, { 'User Principal Name': 'nodept@x.com' }, { 'User Principal Name': 'A1B2C3D4E5F6A7B8C9D0E1F2' }];
    const r = rollUp(extra, deptOf, ['sendcount'], 5, new Set(['sales0@x.com', 'sales1@x.com']));
    expect(r.unmatched).toBe(2);
    expect(r.excluded).toBe(2);
    expect(r.rows[0].people).toBe(5);
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
});
