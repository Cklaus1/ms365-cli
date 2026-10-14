import { describe, it, expect } from 'vitest';
import { runCommand } from '../helpers/cli.js';
import { registerExcelCommands } from '../../src/commands/excel.js';

const WS = '/drives/D/items/I/workbook/worksheets/S';

describe('excel sort-range (M-4)', () => {
  it('POSTs to range(...)/sort/apply with fields and matchCase', async () => {
    const r = await runCommand(registerExcelCommands,
      ['excel', 'sort-range', 'D', 'I', 'S', 'A1:D10', '--fields', '[{"key":0,"ascending":true}]', '--match-case'], {
        routes: [['POST', /\/sort\/apply$/, () => new Response(null, { status: 204 })]],
      });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(1);
    expect(r.calls[0]).toMatchObject({ method: 'POST', path: `${WS}/range(address='A1:D10')/sort/apply` });
    expect(r.calls[0].body).toEqual({ fields: [{ key: 0, ascending: true }], matchCase: true });
    expect(r.stdout).toBe('');
    expect(r.stderr).toContain('Range sorted.');
  });

  it('rejects invalid --fields JSON with exit 2 and no request', async () => {
    const r = await runCommand(registerExcelCommands, ['excel', 'sort-range', 'D', 'I', 'S', 'A1:B2', '--fields', 'nope']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});

describe('excel format-range (M-5)', () => {
  const routes: Parameters<typeof runCommand>[2] = {
    routes: [
      ['PATCH', /\/format\/font$/, (c) => ({ ...(c.body as object) })],
      ['PATCH', /\/format\/fill$/, (c) => ({ ...(c.body as object) })],
    ],
  };

  it('sends font and fill as separate PATCHes', async () => {
    const r = await runCommand(registerExcelCommands,
      ['excel', 'format-range', 'D', 'I', 'S', 'A1:B2', '--font-bold', '--font-color', '#FF0000', '--fill-color', '#FFFF00', '-o', 'json'], routes);
    expect(r.exitCode).toBe(0);
    expect(r.calls.map((c) => [c.method, c.path])).toEqual([
      ['PATCH', `${WS}/range(address='A1:B2')/format/font`],
      ['PATCH', `${WS}/range(address='A1:B2')/format/fill`],
    ]);
    expect(r.calls[0].body).toEqual({ bold: true, color: '#FF0000' });
    expect(r.calls[1].body).toEqual({ color: '#FFFF00' });
  });

  it('fill colour alone sends only the fill PATCH (--font-bold optional)', async () => {
    const r = await runCommand(registerExcelCommands,
      ['excel', 'format-range', 'D', 'I', 'S', 'A1:B2', '--fill-color', '#00FF00'], routes);
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(1);
    expect(r.calls[0]).toMatchObject({ method: 'PATCH', path: `${WS}/range(address='A1:B2')/format/fill` });
    expect(r.calls[0].body).toEqual({ color: '#00FF00' });
  });

  it('--font-bold false sends bold:false to font only', async () => {
    const r = await runCommand(registerExcelCommands,
      ['excel', 'format-range', 'D', 'I', 'S', 'A1:B2', '--font-bold', 'false'], routes);
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(1);
    expect(r.calls[0].path).toMatch(/\/format\/font$/);
    expect(r.calls[0].body).toEqual({ bold: false });
  });

  it('no formatting option exits 2 and sends nothing', async () => {
    const r = await runCommand(registerExcelCommands, ['excel', 'format-range', 'D', 'I', 'S', 'A1:B2'], routes);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});

describe('excel range address quoting (M-24)', () => {
  it('range doubles single quotes in the address literal', async () => {
    const r = await runCommand(registerExcelCommands, ['excel', 'range', 'D', 'I', 'S', "'O'Brien'!A1"], {
      routes: [['GET', /\/range\(/, () => ({ address: 'x' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].path).toBe(`${WS}/range(address='''O''Brien''!A1')`);
  });

  it('format-range and sort-range also quote the address', async () => {
    const f = await runCommand(registerExcelCommands,
      ['excel', 'format-range', 'D', 'I', 'S', "O'B!A1", '--fill-color', '#000000'], {
        routes: [['PATCH', /\/format\/fill$/, () => ({})]],
      });
    expect(f.calls[0].path).toBe(`${WS}/range(address='O''B!A1')/format/fill`);
    const s = await runCommand(registerExcelCommands,
      ['excel', 'sort-range', 'D', 'I', 'S', "O'B!A1", '--fields', '[]'], {
        routes: [['POST', /\/sort\/apply$/, () => new Response(null, { status: 204 })]],
      });
    expect(s.calls[0].path).toBe(`${WS}/range(address='O''B!A1')/sort/apply`);
  });
});

describe('excel worksheets pages through printList (Q-8)', () => {
  const prefix = '/drives/D/items/I/workbook/worksheets';
  const routes = [['GET', prefix, (c: { url: string }) => /skiptoken=2/.test(c.url)
    ? { value: [{ id: 'b' }] }
    : { value: [{ id: 'a' }], '@odata.nextLink': `https://graph.microsoft.com/v1.0${prefix}?$skiptoken=2` }]] as never;

  it('page 1 carries nextCursor; --all fetches both pages', async () => {
    const r1 = await runCommand(registerExcelCommands, ['excel', 'worksheets', 'D', 'I', '-o', 'json'], { routes });
    expect(JSON.parse(r1.stdout).nextCursor).toEqual(expect.any(String));
    const r2 = await runCommand(registerExcelCommands, ['excel', 'worksheets', 'D', 'I', '--all', '-o', 'json'], { routes });
    expect(JSON.parse(r2.stdout).value.map((v: { id: string }) => v.id)).toEqual(['a', 'b']);
  });
});

describe('excel sort-range confirmation is silent under --dry-run (DR-1)', () => {
  it('prints no "Range sorted." when nothing was sent', async () => {
    const r = await runCommand(registerExcelCommands, ['excel', 'sort-range', 'D', 'I', 'S', 'A1:B2', '--fields', '[]'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).not.toContain('Range sorted.');
  });
});

describe('excel create-chart --series-by choices (V-19)', () => {
  it('rejects an unknown value with exit 2 before any request', async () => {
    const r = await runCommand(registerExcelCommands, ['excel', 'create-chart', 'D', 'I', 'S', '--type', 'Pie', '--source', 'A1:B2', '--series-by', 'diagonal']);
    expect(r.exitCode).toBe(2);
    expect(r.stderr).toContain('auto, columns, rows');
    expect(r.calls).toHaveLength(0);
  });
});
