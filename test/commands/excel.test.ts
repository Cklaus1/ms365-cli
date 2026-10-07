import { describe, it, expect } from 'vitest';
import { runCli } from '../e2e/run-cli.js';

const ENV = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_TOKEN: 'fake' };
const WS = '/drives/D1/items/I1/workbook/worksheets/S1';

describe('EXCEL-1', () => {
  it('sort-range POSTs …/sort/apply with fields and matchCase', async () => {
    const r = await runCli(['excel', 'sort-range', 'D1', 'I1', 'S1', 'A1:D10', '--fields', '[{"key":0,"ascending":true}]', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/sort/apply', method: 'POST', status: 204 }],
    });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(1);
    expect(r.requests[0].method).toBe('POST');
    expect(r.requests[0].url).toBe(`https://graph.microsoft.com/v1.0${WS}/range(address='A1%3AD10')/sort/apply`);
    expect(JSON.parse(r.requests[0].body!)).toEqual({ fields: [{ key: 0, ascending: true }], matchCase: false });
    expect(() => JSON.parse(r.stdout)).not.toThrow();
  });

  it('format-range PATCHes /format/font and /format/fill separately; --font-bold is optional', async () => {
    const r = await runCli(['excel', 'format-range', 'D1', 'I1', 'S1', 'A1', '--font-color', '#FF0000', '--fill-color', '#FFFF00', '-o', 'json'], {
      env: ENV,
      routes: [
        { match: '/format/font', method: 'PATCH', body: { color: '#FF0000' } },
        { match: '/format/fill', method: 'PATCH', body: { color: '#FFFF00' } },
      ],
    });
    expect(r.code).toBe(0);
    const byUrl = Object.fromEntries(r.requests.map((q) => [q.url.slice(q.url.lastIndexOf('/format')), q]));
    expect(Object.keys(byUrl).sort()).toEqual(['/format/fill', '/format/font']);
    expect(byUrl['/format/font'].method).toBe('PATCH');
    expect(JSON.parse(byUrl['/format/font'].body!)).toEqual({ color: '#FF0000' });
    expect(JSON.parse(byUrl['/format/fill'].body!)).toEqual({ color: '#FFFF00' });
    expect(JSON.parse(r.stdout)).toEqual({ font: { color: '#FF0000' }, fill: { color: '#FFFF00' } });
  });

  it('format-range --font-bold false sends bold:false; no format option → nothing to update, no request', async () => {
    const bold = await runCli(['--dry-run', 'excel', 'format-range', 'D1', 'I1', 'S1', 'A1', '--font-bold', 'false'], { env: ENV, routes: [] });
    expect(bold.code).toBe(0);
    expect(bold.stderr).toContain('/format/font');
    expect(bold.stderr).toContain('{"bold":false}');
    expect(bold.stderr).not.toContain('/format/fill');

    const none = await runCli(['excel', 'format-range', 'D1', 'I1', 'S1', 'A1'], { env: ENV, routes: [] });
    expect(none.code).toBe(1);
    expect(none.stderr).toMatch(/nothing to update/i);
    expect(none.requests).toHaveLength(0);
  });

  it("escapes ' in sheet-qualified addresses as ''", async () => {
    const r = await runCli(['--dry-run', 'excel', 'range', 'D1', 'I1', 'S1', "O'Brien!A1:B2"], { env: ENV, routes: [] });
    expect(r.code).toBe(0);
    expect(r.stderr).toContain(`range(address='O''Brien!A1%3AB2')`);
  });
});

describe('PAGE-2 excel worksheets drains every page', () => {
  it('prints one combined JSON array from two pages', async () => {
    const r = await runCli(['excel', 'worksheets', 'D1', 'I1', '-o', 'json'], {
      env: ENV,
      routes: [
        { match: 'skiptoken=2', body: { value: [{ id: 'b' }] } },
        { match: '/workbook/worksheets', body: { value: [{ id: 'a' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/drives/D1/items/I1/workbook/worksheets?$skiptoken=2' } },
      ],
    });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout)).toEqual([{ id: 'a' }, { id: 'b' }]);
    expect(r.requests).toHaveLength(2);
  });
});

describe('CHOICE-1 excel create-chart choices', () => {
  const base = ['excel', 'create-chart', 'D1', 'I1', 'S1', '--source', 'A1:B2'];
  for (const args of [['--type', 'Columns', '--series-by', 'auto'], ['--type', 'Pie', '--series-by', 'diagonal'], ['--type', 'BoxwhiskerEx', '--series-by', 'auto']]) {
    it(`${args.join(' ')} exits 1 with Allowed choices before any request`, async () => {
      const r = await runCli([...base, ...args], { env: ENV, routes: [] });
      expect(r.code).toBe(1);
      expect(r.stderr).toContain('Allowed choices');
      expect(r.requests).toHaveLength(0);
    });
  }

  it('an ExcelApi 1.1 type and series-by are sent as given', async () => {
    const r = await runCli([...base, '--type', '3DColumnClustered', '--series-by', 'rows', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/charts/add', body: { id: 'c1' } }],
    });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.requests[0].body ?? '')).toEqual({ type: '3DColumnClustered', sourceData: 'A1:B2', seriesBy: 'rows' });
  });
});
