import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { graph } from '../command-tree.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { saveOrStream } from '../files.js';
import { listRequest, paginationOptions, type ListOptions } from '../list.js';
import { info, outputOption } from '../utils.js';
import { handleErrors } from '../errors.js';

interface PageOpts { out?: string; force?: boolean }
interface CreatePageOpts { title: string; content: string; output: string }

export function registerOneNoteCommands(program: Command, getClient: () => GraphClient): void {
  const onenote = program.command('onenote').description('OneNote notebooks');

  paginationOptions(graph(onenote.command('notebooks'), ['GET /me/onenote/notebooks'])
    .description('List notebooks'))
    .addOption(outputOption())
    .action(handleErrors(async (opts: ListOptions) => {
      await listRequest(getClient(), '/me/onenote/notebooks', {}, opts);
    }));

  paginationOptions(graph(onenote.command('sections <notebookId>'), ['GET /me/onenote/notebooks/{notebook-id}/sections'])
    .description('List sections in a notebook'))
    .addOption(outputOption())
    .action(handleErrors(async (notebookId: string, opts: ListOptions) => {
      await listRequest(getClient(), `/me/onenote/notebooks/${encodeURIComponent(notebookId)}/sections`, {}, opts);
    }));

  paginationOptions(graph(onenote.command('pages <sectionId>'), ['GET /me/onenote/sections/{section-id}/pages'])
    .description('List pages in a section (Graph returns 20 per page by default)'))
    .addOption(outputOption())
    .action(handleErrors(async (sectionId: string, opts: ListOptions) => {
      await listRequest(getClient(), `/me/onenote/sections/${encodeURIComponent(sectionId)}/pages`, {}, opts);
    }));

  graph(onenote.command('page <pageId>'), ['GET /me/onenote/pages/{page-id}/content'])
    .description('Get page content as raw HTML (to stdout or --out)')
    .option('--out <path>', 'Save the HTML to a file instead of stdout')
    .option('--force', 'Overwrite an existing --out file')
    .action(handleErrors(async (pageId: string, opts: PageOpts) => {
      const response = await getClient().requestBinary(`/me/onenote/pages/${encodeURIComponent(pageId)}/content`);
      const saved = await saveOrStream(response.body, { out: opts.out, force: opts.force });
      if (saved.path) info(`Saved to ${saved.path} (${saved.bytes} bytes)`);
    }));

  graph(onenote.command('create-page <sectionId>'), ['POST /me/onenote/sections/{section-id}/pages'])
    .description('Create a page in a section')
    .requiredOption('--title <title>', 'Page title')
    .requiredOption('--content <html>', 'Page content (HTML)')
    .addOption(outputOption())
    .action(handleErrors(async (sectionId: string, opts: CreatePageOpts) => {
      const safeTitle = opts.title.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      const html = `<!DOCTYPE html><html><head><title>${safeTitle}</title></head><body>${opts.content}</body></html>`;
      const response = await getClient().request(`/me/onenote/sections/${encodeURIComponent(sectionId)}/pages`, {
        method: 'POST',
        body: html,
        headers: { 'Content-Type': 'text/html' },
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));
}
