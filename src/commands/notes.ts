import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat, printList } from '../formatter.js';
import { outputOption, confirm, info } from '../utils.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';

/**
 * Outlook quick-capture notes (spec G-8): the `note` resource under /me/notes,
 * { id, subject, body: { contentType, content }, createdDateTime, lastModifiedDateTime }.
 * Built to the documented shape; not yet exercised against a live tenant.
 */
export function registerNotesCommands(program: Command, getClient: () => GraphClient): void {
  const notes = program.command('notes').description('Outlook quick-capture notes');

  const notePath = (id: string) => `/me/notes/${encodeURIComponent(id)}`;

  notes
    .command('list')
    .description('List notes')
    .option('--top <n>', 'Number of notes')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      await printList(client, '/me/notes', { queryParams }, opts);
    }));

  notes
    .command('get <id>')
    .description('Get a note')
    .addOption(outputOption())
    .action(handleErrors(async (id: string, opts) => {
      const client = getClient();
      const response = await client.request(notePath(id));
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  notes
    .command('create')
    .description('Create a note')
    .requiredOption('--body <text>', 'Note text')
    .option('--subject <subject>', 'Note subject')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const note: Record<string, unknown> = {};
      if (opts.subject) note.subject = opts.subject;
      note.body = { contentType: 'text', content: opts.body };
      const response = await client.request('/me/notes', { method: 'POST', body: JSON.stringify(note) });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  notes
    .command('update <id>')
    .description('Update a note')
    .option('--body <text>', 'New note text')
    .option('--subject <subject>', 'New subject')
    .addOption(outputOption())
    .action(handleErrors(async (id: string, opts) => {
      const patch: Record<string, unknown> = {};
      if (opts.subject !== undefined) patch.subject = opts.subject;
      if (opts.body !== undefined) patch.body = { contentType: 'text', content: opts.body };
      if (Object.keys(patch).length === 0) {
        throw new CliError('Nothing to update', { code: 'usage', exitCode: ExitCode.USAGE, fix: 'Pass --body and/or --subject' });
      }
      const client = getClient();
      const response = await client.request(notePath(id), { method: 'PATCH', body: JSON.stringify(patch) });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  notes
    .command('delete <id>')
    .description('Delete a note')
    .action(handleErrors(async (id: string) => {
      if (!await confirm(`Delete note ${id.substring(0, 16)}...?`)) return;
      const client = getClient();
      await client.request(notePath(id), { method: 'DELETE' });
      info('Note deleted.');
    }));
}
