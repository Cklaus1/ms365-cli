import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption } from '../utils.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';

/**
 * Entity types `search query` accepts, by compatibility group. Microsoft Search only combines
 * entity types of the same group in one searchRequest: Exchange/Teams content (message, event,
 * chatMessage) cannot share a request with SharePoint/OneDrive content (driveItem, listItem, site).
 * Mixed --entity values become one POST /search/query per group; their hits are concatenated
 * in group order (ranks are per request, so they are not comparable across groups).
 */
const ENTITY_GROUPS = [
  ['message', 'event', 'chatMessage'],
  ['driveItem', 'listItem', 'site'],
] as const;
const ENTITY_TYPES: string[] = ENTITY_GROUPS.flat();

const usage = (msg: string, fix?: string) => new CliError(msg, { code: 'usage', exitCode: ExitCode.USAGE, fix });

/** The --entity list split into the fewest compatible requests, each in the order given. */
function entityGroups(csv: string): string[][] {
  const types = [...new Set(csv.split(',').map(t => t.trim()).filter(Boolean))];
  const unknown = types.filter(t => !ENTITY_TYPES.includes(t));
  if (unknown.length || types.length === 0) {
    throw usage(`Unknown entity type: ${unknown.join(', ') || '(none)'}`, `Use any of: ${ENTITY_TYPES.join(', ')}`);
  }
  return ENTITY_GROUPS
    .map(group => types.filter(t => (group as readonly string[]).includes(t)))
    .filter(group => group.length > 0);
}

function nonNegativeInt(value: string, flag: string): number {
  const n = Number(value);
  if (!Number.isInteger(n) || n < 0) throw usage(`${flag} must be a non-negative integer`);
  return n;
}

interface SearchHit {
  entityType: string;
  id: string;
  summary?: string;
  rank?: number;
  title?: string;
  webUrl?: string;
}

/** searchResponse → flat hits: value[].hitsContainers[].hits[] { hitId, rank, summary, resource }. */
function flattenHits(data: unknown): SearchHit[] {
  type Hit = { hitId: string; rank?: number; summary?: string; resource?: Record<string, unknown> };
  const responses = (data as { value?: Array<{ hitsContainers?: Array<{ hits?: Hit[] }> }> })?.value ?? [];
  return responses.flatMap(r => r.hitsContainers ?? []).flatMap(c => c.hits ?? []).map(hit => {
    const res = hit.resource ?? {};
    const pick = (...keys: string[]) => keys.map(k => res[k]).find(v => typeof v === 'string') as string | undefined;
    return {
      entityType: String(res['@odata.type'] ?? '').replace('#microsoft.graph.', ''),
      id: hit.hitId,
      summary: hit.summary,
      rank: hit.rank,
      title: pick('subject', 'name', 'displayName', 'title'),
      webUrl: pick('webUrl', 'webLink'),
    };
  });
}

export function registerSearchCommands(program: Command, getClient: () => GraphClient): void {
  const search = program.command('search').description('Microsoft 365 search');

  search
    .command('query <query>')
    .description('Search mail, events, chats, files, sites and list items in one query (hits merged into one list)')
    .option('--entity <types>', `Comma-separated entity types: ${ENTITY_TYPES.join(', ')}`, 'message,event,driveItem')
    .option('--size <n>', 'Number of hits per request', '10')
    .option('--from <n>', 'Zero-based offset of the first hit, per request', '0')
    .addOption(outputOption())
    .action(handleErrors(async (query: string, opts) => {
      const groups = entityGroups(opts.entity);
      const size = nonNegativeInt(opts.size, '--size');
      const from = nonNegativeInt(opts.from, '--from');
      const client = getClient();
      const hits: SearchHit[] = [];
      for (const entityTypes of groups) {
        const response = await client.request('/search/query', {
          method: 'POST',
          body: JSON.stringify({ requests: [{ entityTypes, query: { queryString: query }, from, size }] }),
        });
        hits.push(...flattenHits(response.data));
      }
      console.log(formatOutput({ value: hits }, opts.output as OutputFormat));
    }));

  search
    .command('people [query]')
    .description('Search people')
    .option('--top <n>', 'Number of results', '10')
    .addOption(outputOption())
    .action(handleErrors(async (query: string | undefined, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (query) queryParams['$search'] = `"${query}"`;

      const response = await client.request('/me/people', { queryParams });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));
}
