import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption } from '../utils.js';
import { handleErrors } from '../errors.js';

export function registerExcelCommands(program: Command, getClient: () => GraphClient): void {
  const excel = program.command('excel').description('Excel workbook operations');

  const wbPath = (driveId: string, itemId: string) =>
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook`;

  const wsPath = (driveId: string, itemId: string, worksheetId: string) =>
    `${wbPath(driveId, itemId)}/worksheets/${encodeURIComponent(worksheetId)}`;

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
      const response = await client.request(
        `${wsPath(driveId, itemId, worksheetId)}/range(address='${encodeURIComponent(address)}')`
      );
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
    .requiredOption('--font-bold [bool]', 'Bold font')
    .option('--font-color <color>', 'Font color (e.g. #FF0000)')
    .option('--fill-color <color>', 'Fill color (e.g. #FFFF00)')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, worksheetId: string, address: string, opts) => {
      const client = getClient();
      const format: Record<string, unknown> = {};

      if (opts.fontBold !== undefined) {
        format.font = { bold: opts.fontBold === true || opts.fontBold === 'true' };
      }
      if (opts.fontColor) {
        format.font = { ...(format.font as Record<string, unknown> || {}), color: opts.fontColor };
      }
      if (opts.fillColor) {
        format.fill = { color: opts.fillColor };
      }

      const response = await client.request(
        `${wsPath(driveId, itemId, worksheetId)}/range(address='${encodeURIComponent(address)}')/format`,
        {
          method: 'PATCH',
          body: JSON.stringify(format),
        }
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  excel
    .command('sort-range <driveId> <itemId> <worksheetId> <address>')
    .description('Sort a cell range (e.g. "A1:D10")')
    .requiredOption('--fields <json>', 'Sort fields JSON (e.g. [{"key":0,"ascending":true}])')
    .option('--match-case', 'Case-sensitive sort')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, worksheetId: string, address: string, opts) => {
      const client = getClient();
      let fields;
      try {
        fields = JSON.parse(opts.fields);
      } catch {
        throw new Error('--fields must be valid JSON array');
      }

      const response = await client.request(
        `${wsPath(driveId, itemId, worksheetId)}/range(address='${encodeURIComponent(address)}')/sort`,
        {
          method: 'PATCH',
          body: JSON.stringify({
            fields,
            matchCase: opts.matchCase || false,
          }),
        }
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));
}
