import fs from 'fs';
import { Command, Option } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat, printList } from '../formatter.js';
import { addListOptions, outputOption, odataQuote, success } from '../utils.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';

export function registerExcelCommands(program: Command, getClient: () => GraphClient): void {
  const excel = program.command('excel').description('Excel workbook operations');

  const wbPath = (driveId: string, itemId: string) =>
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook`;

  const wsPath = (driveId: string, itemId: string, worksheetId: string) =>
    `${wbPath(driveId, itemId)}/worksheets/${encodeURIComponent(worksheetId)}`;

  const rangePath = (driveId: string, itemId: string, worksheetId: string, address: string) =>
    `${wsPath(driveId, itemId, worksheetId)}/range(address='${encodeURIComponent(odataQuote(address))}')`;

  addListOptions(excel
    .command('worksheets <driveId> <itemId>')
    .description('List worksheets in a workbook'))
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, opts) => {
      await printList(getClient(), `${wbPath(driveId, itemId)}/worksheets`, {}, opts);
    }));

  excel
    .command('range <driveId> <itemId> <worksheetId> <address>')
    .description('Get a cell range (e.g. "A1:D10")')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, worksheetId: string, address: string, opts) => {
      const client = getClient();
      const response = await client.request(rangePath(driveId, itemId, worksheetId, address));
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  excel
    .command('create-chart <driveId> <itemId> <worksheetId>')
    .description('Create a chart')
    .requiredOption('--type <type>', 'Chart type (e.g. ColumnClustered, Pie, Line)')
    .requiredOption('--source <range>', 'Source data range (e.g. A1:D10)')
    .addOption(new Option('--series-by <by>', 'Series grouped by').choices(['auto', 'columns', 'rows']).makeOptionMandatory())
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, worksheetId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `${wsPath(driveId, itemId, worksheetId)}/charts/add`,
        {
          method: 'POST',
          body: JSON.stringify({
            type: opts.type,
            sourceData: opts.source,
            seriesBy: opts.seriesBy,
          }),
        }
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  excel
    .command('format-range <driveId> <itemId> <worksheetId> <address>')
    .description('Format a cell range (e.g. "A1:D10")')
    .option('--font-bold [bool]', 'Bold font (true/false)')
    .option('--font-color <color>', 'Font color (e.g. #FF0000)')
    .option('--fill-color <color>', 'Fill color (e.g. #FFFF00)')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, worksheetId: string, address: string, opts) => {
      const font: Record<string, unknown> = {};
      if (opts.fontBold !== undefined) {
        font.bold = opts.fontBold === true || opts.fontBold === 'true';
      }
      if (opts.fontColor) {
        font.color = opts.fontColor;
      }
      if (Object.keys(font).length === 0 && !opts.fillColor) {
        throw new CliError('No formatting option given', {
          code: 'usage',
          exitCode: ExitCode.USAGE,
          fix: 'Pass at least one of --font-bold, --font-color, --fill-color',
        });
      }

      const client = getClient();
      const base = `${rangePath(driveId, itemId, worksheetId, address)}/format`;
      const result: Record<string, unknown> = {};
      if (Object.keys(font).length > 0) {
        const response = await client.request(`${base}/font`, { method: 'PATCH', body: JSON.stringify(font) });
        result.font = response.data;
      }
      if (opts.fillColor) {
        const response = await client.request(`${base}/fill`, {
          method: 'PATCH',
          body: JSON.stringify({ color: opts.fillColor }),
        });
        result.fill = response.data;
      }
      console.log(formatOutput(result, opts.output as OutputFormat));
    }));

  excel
    .command('sort-range <driveId> <itemId> <worksheetId> <address>')
    .description('Sort a cell range (e.g. "A1:D10")')
    .requiredOption('--fields <json>', 'Sort fields JSON (e.g. [{"key":0,"ascending":true}])')
    .option('--match-case', 'Case-sensitive sort')
    .action(handleErrors(async (driveId: string, itemId: string, worksheetId: string, address: string, opts) => {
      const client = getClient();
      let fields;
      try {
        fields = JSON.parse(opts.fields);
      } catch {
        throw new CliError('--fields must be valid JSON array', { code: 'usage', exitCode: ExitCode.USAGE });
      }

      await client.request(
        `${rangePath(driveId, itemId, worksheetId, address)}/sort/apply`,
        {
          method: 'POST',
          body: JSON.stringify({
            fields,
            matchCase: opts.matchCase || false,
          }),
        }
      );
      success('Range sorted.');
    }));

  const tablePath = (driveId: string, itemId: string, table: string) =>
    `${wbPath(driveId, itemId)}/tables/${encodeURIComponent(table)}`;

  addListOptions(excel
    .command('tables <driveId> <itemId>')
    .description('List tables in a workbook'))
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, opts) => {
      const client = getClient();
      await printList(client, `${wbPath(driveId, itemId)}/tables`, {}, opts);
    }));

  addListOptions(excel
    .command('table-rows <driveId> <itemId> <table>')
    .description('List rows of a table (table name or ID)'), { description: 'Number of rows' })
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, table: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      await printList(client, `${tablePath(driveId, itemId, table)}/rows`, { queryParams }, opts);
    }));

  excel
    .command('add-rows <driveId> <itemId> <table>')
    .description('Append rows to a table (table name or ID)')
    .option('--values <json>', 'Rows as a JSON array of arrays (e.g. [["a",1],["b",2]])')
    .option('--values-file <path>', 'Rows from a .json (array of arrays) or .csv file')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, table: string, opts) => {
      const values = readRows(opts.values, opts.valuesFile);
      const client = getClient();
      // POST tables/{id}/rows is the documented tableRowCollection add; body {values: [[...]]}.
      const response = await client.request(`${tablePath(driveId, itemId, table)}/rows`, {
        method: 'POST',
        body: JSON.stringify({ values }),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));
}

const usageError = (msg: string, fix?: string) => new CliError(msg, { code: 'usage', exitCode: ExitCode.USAGE, fix });

/** Rows for add-rows from exactly one of --values (JSON) or --values-file (.json or .csv). */
function readRows(json: string | undefined, file: string | undefined): unknown[][] {
  if ((json === undefined) === (file === undefined)) {
    throw usageError('Pass exactly one of --values or --values-file', 'e.g. --values \'[["a",1]]\' or --values-file rows.csv');
  }
  let text = json;
  if (file !== undefined) {
    try {
      text = fs.readFileSync(file, 'utf8');
    } catch {
      throw usageError(`Cannot read --values-file ${file}`);
    }
    if (/\.csv$/i.test(file)) return checkRows(parseCsv(text));
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(text!);
  } catch {
    throw usageError('--values must be valid JSON: an array of arrays');
  }
  return checkRows(parsed);
}

function checkRows(rows: unknown): unknown[][] {
  if (!Array.isArray(rows) || rows.length === 0 || !rows.every(Array.isArray)) {
    throw usageError('Rows must be a non-empty array of arrays (one inner array per row)', 'e.g. [["a",1],["b",2]]');
  }
  return rows as unknown[][];
}

/** Comma-separated rows; double quotes wrap fields that hold commas, quotes ("") or newlines. */
export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (ch === '"') quoted = false;
      else field += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === ',') { row.push(field); field = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += ch;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}
