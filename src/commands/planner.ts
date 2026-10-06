import { Command, Option } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { handleErrors } from '../errors.js';
import { parseDate, outputOption, confirm, info } from '../utils.js';

interface OutputOpts { output: string }
interface EtagOpts { etag?: string }
interface CreateTaskOpts extends OutputOpts { plan: string; title: string; bucket?: string; due?: string; assignedTo?: string }
interface UpdateTaskOpts extends OutputOpts, EtagOpts { title?: string; due?: string; percent?: string; bucket?: string }
interface PlanDetailsOpts extends OutputOpts, EtagOpts { categories?: string }
interface TaskDetailsOpts extends OutputOpts, EtagOpts { description?: string }
interface UpdatePlanOpts extends OutputOpts, EtagOpts { title?: string }
interface UpdateBucketOpts extends OutputOpts, EtagOpts { name?: string }
interface CreatePlanOpts extends OutputOpts { title: string; group: string }
interface CreateBucketOpts extends OutputOpts { plan: string; name: string }

/**
 * The If-Match value for a write to `endpoint`. `--read-only` is checked first, so a blocked write
 * never makes the lookup request. Otherwise: `--etag`, a placeholder in dry-run (nothing is sent),
 * or the resource's `@odata.etag` read through the client.
 */
async function ifMatch(client: GraphClient, method: string, endpoint: string, etag: string | undefined): Promise<string> {
  client.assertAllowed(method, endpoint);
  if (etag) return etag;
  if (client.isDryRun) return '<dry-run>';
  const { data } = await client.request(endpoint);
  const value = typeof data === 'object' && data !== null && '@odata.etag' in data ? data['@odata.etag'] : undefined;
  if (typeof value !== 'string') throw new Error(`No @odata.etag on ${endpoint}. Provide --etag.`);
  return value;
}

/** PATCH `body` with If-Match and print the result; an empty body is rejected before any request. */
async function patch(client: GraphClient, endpoint: string, body: Record<string, unknown>, opts: OutputOpts & EtagOpts): Promise<void> {
  if (Object.keys(body).length === 0) throw new Error('Nothing to update: pass at least one field option.');
  const headers = { 'If-Match': await ifMatch(client, 'PATCH', endpoint, opts.etag) };
  const response = await client.request(endpoint, { method: 'PATCH', body: JSON.stringify(body), headers });
  console.log(formatOutput(response.data, opts.output as OutputFormat));
}

async function remove(client: GraphClient, endpoint: string, prompt: string, opts: EtagOpts): Promise<void> {
  client.assertAllowed('DELETE', endpoint);
  await confirm(prompt);
  const headers = { 'If-Match': await ifMatch(client, 'DELETE', endpoint, opts.etag) };
  await client.request(endpoint, { method: 'DELETE', headers });
}

/** A due date (YYYY-MM-DD) as the start of that day in the local time zone, so Planner shows that day. */
function dueDateTime(date: string): string {
  const [year, month, day] = parseDate(date).split('-').map(Number);
  return new Date(year, month - 1, day).toISOString();
}

export function registerPlannerCommands(program: Command, getClient: () => GraphClient): void {
  const planner = program.command('planner').description('Microsoft Planner tasks');
  const taskPath = (id: string) => `/planner/tasks/${encodeURIComponent(id)}`;
  const planPath = (id: string) => `/planner/plans/${encodeURIComponent(id)}`;
  const bucketPath = (id: string) => `/planner/buckets/${encodeURIComponent(id)}`;

  const get = (name: string, description: string, endpoint: (id: string) => string) =>
    planner
      .command(name)
      .description(description)
      .addOption(outputOption())
      .action(handleErrors(async (id: string, opts: OutputOpts) => {
        const response = await getClient().request(endpoint(id));
        console.log(formatOutput(response.data, opts.output as OutputFormat));
      }));

  planner
    .command('plans')
    .description('List my planner plans')
    .addOption(outputOption())
    .action(handleErrors(async (opts: OutputOpts) => {
      const response = await getClient().request('/me/planner/plans');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  get('plan <planId>', 'Get plan details', planPath);
  get('plan-details <planId>', 'Get plan extended details (category labels, etc.)', (id) => `${planPath(id)}/details`);

  planner
    .command('my-tasks')
    .description('List my planner tasks')
    .addOption(outputOption())
    .action(handleErrors(async (opts: OutputOpts) => {
      const response = await getClient().request('/me/planner/tasks');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  get('tasks <planId>', 'List tasks in a plan', (id) => `${planPath(id)}/tasks`);
  get('task <taskId>', 'Get task details', taskPath);
  get('task-details <taskId>', 'Get task extended details (description, checklist, etc.)', (id) => `${taskPath(id)}/details`);
  get('buckets <planId>', 'List plan buckets', (id) => `${planPath(id)}/buckets`);
  get('bucket <bucketId>', 'Get bucket details', bucketPath);
  get('bucket-tasks <bucketId>', 'List tasks in a bucket', (id) => `${bucketPath(id)}/tasks`);

  planner
    .command('create-task')
    .description('Create a planner task')
    .requiredOption('--plan <planId>', 'Plan ID')
    .requiredOption('--title <title>', 'Task title')
    .option('--bucket <bucketId>', 'Bucket ID')
    .option('--due <date>', 'Due date (YYYY-MM-DD, local time zone)')
    .option('--assigned-to <userId>', 'Assign to user ID')
    .addOption(outputOption())
    .action(handleErrors(async (opts: CreateTaskOpts) => {
      const task: Record<string, unknown> = {
        planId: opts.plan,
        title: opts.title,
      };
      if (opts.bucket) task.bucketId = opts.bucket;
      if (opts.due) task.dueDateTime = dueDateTime(opts.due);
      if (opts.assignedTo) {
        task.assignments = { [opts.assignedTo]: { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' } };
      }

      const response = await getClient().request('/planner/tasks', {
        method: 'POST',
        body: JSON.stringify(task),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('update-task <taskId>')
    .description('Update a planner task')
    .option('--title <title>', 'Task title')
    .option('--due <date>', 'Due date (YYYY-MM-DD, local time zone)')
    .addOption(new Option('--percent <n>', 'Percent complete').choices(['0', '50', '100']))
    .option('--bucket <bucketId>', 'Move to bucket')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .addOption(outputOption())
    .action(handleErrors(async (taskId: string, opts: UpdateTaskOpts) => {
      const body: Record<string, unknown> = {};
      if (opts.title) body.title = opts.title;
      if (opts.due) body.dueDateTime = dueDateTime(opts.due);
      if (opts.percent !== undefined) body.percentComplete = Number(opts.percent);
      if (opts.bucket) body.bucketId = opts.bucket;
      await patch(getClient(), taskPath(taskId), body, opts);
    }));

  planner
    .command('update-plan-details <planId>')
    .description('Update plan details (category labels)')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .option('--categories <json>', 'Category descriptions as JSON (e.g. {"category1":"Urgent"})')
    .addOption(outputOption())
    .action(handleErrors(async (planId: string, opts: PlanDetailsOpts) => {
      const body: Record<string, unknown> = {};
      if (opts.categories) {
        let categories: unknown;
        try {
          categories = JSON.parse(opts.categories);
        } catch {
          throw new Error('--categories must be a JSON object, e.g. {"category1":"Urgent"}');
        }
        if (typeof categories !== 'object' || categories === null || Array.isArray(categories)) {
          throw new Error('--categories must be a JSON object, e.g. {"category1":"Urgent"}');
        }
        body.categoryDescriptions = categories;
      }
      await patch(getClient(), `${planPath(planId)}/details`, body, opts);
    }));

  planner
    .command('update-task-details <taskId>')
    .description('Update task details (description)')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .option('--description <text>', 'Task description')
    .addOption(outputOption())
    .action(handleErrors(async (taskId: string, opts: TaskDetailsOpts) => {
      const body: Record<string, unknown> = {};
      if (opts.description) body.description = opts.description;
      await patch(getClient(), `${taskPath(taskId)}/details`, body, opts);
    }));

  planner
    .command('delete-task <taskId>')
    .description('Delete a planner task')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .action(handleErrors(async (taskId: string, opts: EtagOpts) => {
      await remove(getClient(), taskPath(taskId), `Delete planner task ${taskId.substring(0, 16)}...?`, opts);
      info('Task deleted.');
    }));

  planner
    .command('delete-plan <planId>')
    .description('Delete a planner plan')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .action(handleErrors(async (planId: string, opts: EtagOpts) => {
      await remove(getClient(), planPath(planId), `Delete planner plan ${planId.substring(0, 16)}...?`, opts);
      info('Plan deleted.');
    }));

  planner
    .command('delete-bucket <bucketId>')
    .description('Delete a planner bucket')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .action(handleErrors(async (bucketId: string, opts: EtagOpts) => {
      await remove(getClient(), bucketPath(bucketId), `Delete planner bucket ${bucketId.substring(0, 16)}...?`, opts);
      info('Bucket deleted.');
    }));

  planner
    .command('update-plan <planId>')
    .description('Update a planner plan')
    .option('--title <title>', 'Plan title')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .addOption(outputOption())
    .action(handleErrors(async (planId: string, opts: UpdatePlanOpts) => {
      const body: Record<string, unknown> = {};
      if (opts.title) body.title = opts.title;
      await patch(getClient(), planPath(planId), body, opts);
    }));

  planner
    .command('update-bucket <bucketId>')
    .description('Update a planner bucket')
    .option('--name <name>', 'Bucket name')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .addOption(outputOption())
    .action(handleErrors(async (bucketId: string, opts: UpdateBucketOpts) => {
      const body: Record<string, unknown> = {};
      if (opts.name) body.name = opts.name;
      await patch(getClient(), bucketPath(bucketId), body, opts);
    }));

  planner
    .command('create-plan')
    .description('Create a planner plan in a Microsoft 365 group')
    .requiredOption('--title <title>', 'Plan title')
    .requiredOption('--group <groupId>', 'Group/Team ID (plan container)')
    .addOption(outputOption())
    .action(handleErrors(async (opts: CreatePlanOpts) => {
      const response = await getClient().request('/planner/plans', {
        method: 'POST',
        body: JSON.stringify({
          title: opts.title,
          container: { containerId: opts.group, type: 'group' },
        }),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('create-bucket')
    .description('Create a bucket in a plan')
    .requiredOption('--plan <planId>', 'Plan ID')
    .requiredOption('--name <name>', 'Bucket name')
    .addOption(outputOption())
    .action(handleErrors(async (opts: CreateBucketOpts) => {
      const response = await getClient().request('/planner/buckets', {
        method: 'POST',
        body: JSON.stringify({
          name: opts.name,
          planId: opts.plan,
        }),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));
}
