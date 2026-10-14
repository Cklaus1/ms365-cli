import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption, odataQuote, info } from '../utils.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';

export function registerExcelCommands(program: Command, getClient: () => GraphClient): void {
  const excel = program.command('excel').description('Excel workbook operations');

  const wbPath = (driveId: string, itemId: string) =>
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook`;

  const wsPath = (driveId: string, itemId: string, worksheetId: string) =>
    `${wbPath(driveId, itemId)}/worksheets/${encodeURIComponent(worksheetId)}`;

  const rangePath = (driveId: string, itemId: string, worksheetId: string, address: string) =>
    `${wsPath(driveId, itemId, worksheetId)}/range(address='${encodeURIComponent(odataQuote(address))}')`;

  excel
    .command('worksheets <driveId> <itemId>')
    .description('List worksheets in a workbook')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, opts) => {
      const client = getClient();
      const response = await client.request(`${wbPath(driveId, itemId)}/worksheets`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
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
    .requiredOption('--series-by <by>', 'Series grouped by: auto, columns, rows')
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
      info('Range sorted.');
    }));
}
