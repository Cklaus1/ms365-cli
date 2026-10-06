import { Command, Option } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { handleErrors } from '../errors.js';
import { parseDate, outputOption, confirm, info } from '../utils.js';
import { paginationOptions, listRequest, type ListOptions } from '../list.js';

const IMPORTANCE = ['low', 'normal', 'high'];
const STATUS = ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'];

interface TasksOptions extends ListOptions {
  filter?: string;
  orderby?: string;
}

interface UpdateOptions {
  title?: string;
  body?: string;
  status?: string;
  importance?: string;
  due?: string;
  output: string;
}

export function registerTodoCommands(program: Command, getClient: () => GraphClient): void {
  const todo = program.command('todo').description('Microsoft To Do tasks');

  paginationOptions(
    todo
      .command('lists')
      .description('List task lists'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (opts: ListOptions) => {
      await listRequest(getClient(), '/me/todo/lists', {}, opts);
    }));

  paginationOptions(
    todo
      .command('tasks <listId>')
      .description('List tasks in a list')
      .option('--filter <filter>', 'OData filter')
      .option('--orderby <field>', 'Order results by field'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (listId: string, opts: TasksOptions) => {
      const queryParams: Record<string, string> = {};
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.orderby) queryParams['$orderby'] = opts.orderby;
      await listRequest(getClient(), `/me/todo/lists/${encodeURIComponent(listId)}/tasks`, { queryParams }, opts);
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
    .addOption(new Option('--importance <level>', 'Importance').choices(IMPORTANCE).default('normal'))
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
    .addOption(new Option('--status <status>', 'Status').choices(STATUS))
    .addOption(new Option('--importance <level>', 'Importance').choices(IMPORTANCE))
    .option('--due <date>', 'Due date (YYYY-MM-DD)')
    .addOption(outputOption())
    .action(handleErrors(async (listId: string, taskId: string, opts: UpdateOptions) => {
      const patch: Record<string, unknown> = {};
      if (opts.title) patch.title = opts.title;
      if (opts.status) patch.status = opts.status;
      if (opts.importance) patch.importance = opts.importance;
      if (opts.body) patch.body = { contentType: 'text', content: opts.body };
      if (opts.due) patch.dueDateTime = { dateTime: parseDate(opts.due) + 'T00:00:00', timeZone: 'UTC' };
      if (Object.keys(patch).length === 0) {
        throw new Error('Nothing to update: give at least one of --title, --body, --status, --importance, --due.');
      }

      const response = await getClient().request(
        `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`,
        { method: 'PATCH', body: JSON.stringify(patch) }
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  todo
    .command('delete <listId> <taskId>')
    .description('Delete a task')
    .action(handleErrors(async (listId: string, taskId: string) => {
      await confirm(`Delete task ${taskId.substring(0, 16)}...?`);
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
