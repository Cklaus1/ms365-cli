import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { outputOption } from '../utils.js';
import { handleErrors } from '../errors.js';

export function registerOneNoteCommands(program: Command, getClient: () => GraphClient): void {
  const onenote = program.command('onenote').description('OneNote notebooks');

  onenote
    .command('notebooks')
    .description('List notebooks')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/onenote/notebooks');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  onenote
    .command('sections <notebookId>')
    .description('List sections in a notebook')
    .addOption(outputOption())
    .action(handleErrors(async (notebookId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/me/onenote/notebooks/${encodeURIComponent(notebookId)}/sections`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  onenote
    .command('pages <sectionId>')
    .description('List pages in a section')
    .addOption(outputOption())
    .action(handleErrors(async (sectionId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/me/onenote/sections/${encodeURIComponent(sectionId)}/pages`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  onenote
    .command('page <pageId>')
    .description('Get page content')
    .action(handleErrors(async (pageId: string) => {
      const client = getClient();
      const response = await client.request(
        `/me/onenote/pages/${encodeURIComponent(pageId)}/content`
      );
      console.log(formatOutput(response.data, 'json'));
    }));

  onenote
    .command('create-page <sectionId>')
    .description('Create a page in a section')
    .requiredOption('--title <title>', 'Page title')
    .requiredOption('--content <html>', 'Page content (HTML)')
    .addOption(outputOption())
    .action(handleErrors(async (sectionId: string, opts) => {
      const client = getClient();
      const safeTitle = opts.title.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      const html = `<!DOCTYPE html><html><head><title>${safeTitle}</title></head><body>${opts.content}</body></html>`;
      const response = await client.request(
        `/me/onenote/sections/${encodeURIComponent(sectionId)}/pages`,
        {
          method: 'POST',
          body: html,
          headers: { 'Content-Type': 'text/html' },
        }
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));
}
