import { Command, InvalidArgumentError, Option } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { odataQuote } from '../graph-helpers.js';
import { outputOption } from '../utils.js';
import { listRequest } from '../list.js';
import { handleErrors } from '../errors.js';

interface OutputOpts { output: string }
interface ChartOpts extends OutputOpts { type: string; source: string; seriesBy: string }
interface FormatOpts extends OutputOpts { fontBold?: boolean; fontColor?: string; fillColor?: string }
interface SortOpts extends OutputOpts { fields: string; matchCase?: boolean }

/**
 * Office JS `Excel.ChartType` members of API set ExcelApi 1.1, except `Invalid`, as their string values
 * (Graph's chartcollection-add lists only a few and ends with "etc.").
 * Source: https://learn.microsoft.com/en-us/javascript/api/excel/excel.charttype?view=excel-js-1.1
 */
const CHART_TYPES = [
  '3DArea', '3DAreaStacked', '3DAreaStacked100', '3DBarClustered', '3DBarStacked', '3DBarStacked100',
  '3DColumn', '3DColumnClustered', '3DColumnStacked', '3DColumnStacked100', '3DLine', '3DPie', '3DPieExploded',
  'Area', 'AreaStacked', 'AreaStacked100', 'BarClustered', 'BarOfPie', 'BarStacked', 'BarStacked100',
  'Bubble', 'Bubble3DEffect', 'ColumnClustered', 'ColumnStacked', 'ColumnStacked100',
  'ConeBarClustered', 'ConeBarStacked', 'ConeBarStacked100', 'ConeCol', 'ConeColClustered', 'ConeColStacked', 'ConeColStacked100',
  'CylinderBarClustered', 'CylinderBarStacked', 'CylinderBarStacked100', 'CylinderCol', 'CylinderColClustered',
  'CylinderColStacked', 'CylinderColStacked100', 'Doughnut', 'DoughnutExploded', 'Histogram',
  'Line', 'LineMarkers', 'LineMarkersStacked', 'LineMarkersStacked100', 'LineStacked', 'LineStacked100',
  'Pie', 'PieExploded', 'PieOfPie', 'PyramidBarClustered', 'PyramidBarStacked', 'PyramidBarStacked100',
  'PyramidCol', 'PyramidColClustered', 'PyramidColStacked', 'PyramidColStacked100', 'Radar', 'RadarFilled', 'RadarMarkers',
  'StockHLC', 'StockOHLC', 'StockVHLC', 'StockVOHLC', 'Surface', 'SurfaceTopView', 'SurfaceTopViewWireframe', 'SurfaceWireframe',
  'XYScatter', 'XYScatterLines', 'XYScatterLinesNoMarkers', 'XYScatterSmooth', 'XYScatterSmoothNoMarkers',
];

function parseBool(value: string): boolean {
  if (value === 'true') return true;
  if (value === 'false') return false;
  throw new InvalidArgumentError('Expected true or false.');
}

export function registerExcelCommands(program: Command, getClient: () => GraphClient): void {
  const excel = program.command('excel').description('Excel workbook operations');

  const wbPath = (driveId: string, itemId: string) =>
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook`;

  const wsPath = (driveId: string, itemId: string, worksheetId: string) =>
    `${wbPath(driveId, itemId)}/worksheets/${encodeURIComponent(worksheetId)}`;

  // The address is an OData string literal (`'` doubled), URL-encoded as a whole.
  const rangePath = (driveId: string, itemId: string, worksheetId: string, address: string) =>
    `${wsPath(driveId, itemId, worksheetId)}/range(address=${encodeURIComponent(odataQuote(address))})`;

  excel
    .command('worksheets <driveId> <itemId>')
    .description('List worksheets in a workbook')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, opts: OutputOpts) => {
      await listRequest(getClient(), `${wbPath(driveId, itemId)}/worksheets`, {}, { output: opts.output, all: true });
    }));

  excel
    .command('range <driveId> <itemId> <worksheetId> <address>')
    .description('Get a cell range (e.g. "A1:D10")')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, worksheetId: string, address: string, opts: OutputOpts) => {
      const response = await getClient().request(rangePath(driveId, itemId, worksheetId, address));
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  excel
    .command('create-chart <driveId> <itemId> <worksheetId>')
    .description('Create a chart')
    .addOption(new Option('--type <type>', 'Chart type (Excel.ChartType, e.g. ColumnClustered, Pie, Line)').choices(CHART_TYPES).makeOptionMandatory())
    .requiredOption('--source <range>', 'Source data range (e.g. A1:D10)')
    .addOption(new Option('--series-by <by>', 'Series grouped by').choices(['auto', 'columns', 'rows']).makeOptionMandatory())
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, worksheetId: string, opts: ChartOpts) => {
      const response = await getClient().request(`${wsPath(driveId, itemId, worksheetId)}/charts/add`, {
        method: 'POST',
        body: JSON.stringify({
          type: opts.type,
          sourceData: opts.source,
          seriesBy: opts.seriesBy,
        }),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  excel
    .command('format-range <driveId> <itemId> <worksheetId> <address>')
    .description('Format a cell range (e.g. "A1:D10"): font and/or fill')
    .option('--font-bold [bool]', 'Bold font (true/false; flag alone = true)', parseBool)
    .option('--font-color <color>', 'Font color (e.g. #FF0000)')
    .option('--fill-color <color>', 'Fill color (e.g. #FFFF00)')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, worksheetId: string, address: string, opts: FormatOpts) => {
      const font: Record<string, unknown> = {};
      if (opts.fontBold !== undefined) font.bold = opts.fontBold;
      if (opts.fontColor) font.color = opts.fontColor;
      if (Object.keys(font).length === 0 && !opts.fillColor) {
        throw new Error('Nothing to update: pass --font-bold, --font-color or --fill-color.');
      }

      // rangeFormat has no font/fill properties: they are separate resources, each PATCHed on its own.
      const client = getClient();
      const format = rangePath(driveId, itemId, worksheetId, address) + '/format';
      const result: Record<string, unknown> = {};
      if (Object.keys(font).length > 0) {
        result.font = (await client.request(`${format}/font`, { method: 'PATCH', body: JSON.stringify(font) })).data;
      }
      if (opts.fillColor) {
        result.fill = (await client.request(`${format}/fill`, { method: 'PATCH', body: JSON.stringify({ color: opts.fillColor }) })).data;
      }
      console.log(formatOutput(result, opts.output as OutputFormat));
    }));

  excel
    .command('sort-range <driveId> <itemId> <worksheetId> <address>')
    .description('Sort a cell range (e.g. "A1:D10")')
    .requiredOption('--fields <json>', 'Sort fields JSON (e.g. [{"key":0,"ascending":true}])')
    .option('--match-case', 'Case-sensitive sort')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, worksheetId: string, address: string, opts: SortOpts) => {
      let fields: unknown;
      try {
        fields = JSON.parse(opts.fields);
      } catch {
        throw new Error('--fields must be a valid JSON array');
      }
      if (!Array.isArray(fields)) throw new Error('--fields must be a valid JSON array');

      const response = await getClient().request(`${rangePath(driveId, itemId, worksheetId, address)}/sort/apply`, {
        method: 'POST',
        body: JSON.stringify({ fields, matchCase: opts.matchCase ?? false }),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));
}
