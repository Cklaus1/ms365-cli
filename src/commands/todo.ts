import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { handleErrors } from '../errors.js';
import { parseDate, outputOption, confirm, info } from '../utils.js';

export function registerTodoCommands(program: Command, getClient: () => GraphClient): void {
  const todo = program.command('todo').description('Microsoft To Do tasks');

  todo
    .command('lists')
    .description('List task lists')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/todo/lists');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  todo
    .command('tasks <listId>')
    .description('List tasks in a list')
    .option('--top <n>', 'Number of tasks')
    .option('--filter <filter>', 'OData filter')
    .option('--orderby <field>', 'Order results by field')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (listId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.orderby) queryParams['$orderby'] = opts.orderby;

      const endpoint = `/me/todo/lists/${encodeURIComponent(listId)}/tasks`;
      const response = opts.all
        ? await client.requestAllPages(endpoint, { queryParams }, opts.pageLimit ? parseInt(opts.pageLimit) : undefined, opts.pageDelay ? parseInt(opts.pageDelay) : undefined)
        : await client.request(endpoint, { queryParams });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  todo
    .command('get <listId> <taskId>')
    .description('Get a specific task')
    .addOption(outputOption())
    .action(handleErrors(async (listId: string, taskId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  todo
    .command('create <listId>')
    .description('Create a task')
    .requiredOption('--title <title>', 'Task title')
    .option('--body <body>', 'Task body/notes')
    .option('--due <date>', 'Due date (YYYY-MM-DD)')
    .option('--importance <level>', 'Importance: low, normal, high', 'normal')
    .addOption(outputOption())
    .action(handleErrors(async (listId: string, opts) => {
      const client = getClient();
      const task: Record<string, unknown> = {
        title: opts.title,
        importance: opts.importance,
      };
      if (opts.body) {
        task.body = { contentType: 'text', content: opts.body };
      }
      if (opts.due) {
        task.dueDateTime = { dateTime: parseDate(opts.due) + 'T00:00:00', timeZone: 'UTC' };
      }

      const response = await client.request(
        `/me/todo/lists/${encodeURIComponent(listId)}/tasks`,
        { method: 'POST', body: JSON.stringify(task) }
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  todo
    .command('update <listId> <taskId>')
    .description('Update a task')
    .option('--title <title>', 'Task title')
    .option('--body <body>', 'Task body')
    .option('--status <status>', 'Status: notStarted, inProgress, completed, waitingOnOthers, deferred')
    .option('--importance <level>', 'Importance: low, normal, high')
    .option('--due <date>', 'Due date (YYYY-MM-DD)')
    .addOption(outputOption())
    .action(handleErrors(async (listId: string, taskId: string, opts) => {
      const client = getClient();
      const patch: Record<string, unknown> = {};
      if (opts.title) patch.title = opts.title;
      if (opts.status) patch.status = opts.status;
      if (opts.importance) patch.importance = opts.importance;
      if (opts.body) patch.body = { contentType: 'text', content: opts.body };
      if (opts.due) patch.dueDateTime = { dateTime: parseDate(opts.due) + 'T00:00:00', timeZone: 'UTC' };

      const response = await client.request(
        `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`,
        { method: 'PATCH', body: JSON.stringify(patch) }
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  todo
    .command('delete <listId> <taskId>')
    .description('Delete a task')
    .action(handleErrors(async (listId: string, taskId: string) => {
      if (!await confirm(`Delete task ${taskId.substring(0, 16)}...?`)) return;
      const client = getClient();
      await client.request(
        `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`,
        { method: 'DELETE' }
      );
      info('Task deleted.');
    }));

  // ── List management ─────────────────────────────────────────
  todo
    .command('create-list')
    .description('Create a new task list')
    .requiredOption('--name <name>', 'List display name')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/todo/lists', {
        method: 'POST',
        body: JSON.stringify({ displayName: opts.name }),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  todo
    .command('update-list <listId>')
    .description('Rename a task list')
    .requiredOption('--name <name>', 'New display name')
    .addOption(outputOption())
    .action(handleErrors(async (listId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/me/todo/lists/${encodeURIComponent(listId)}`,
        { method: 'PATCH', body: JSON.stringify({ displayName: opts.name }) }
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

}
