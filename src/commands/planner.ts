import { Command } from 'commander';
import { GraphClient, type GraphRequestOptions } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { handleErrors, CliError } from '../errors.js';
import { outputOption, confirm, info, parseTime } from '../utils.js';

const PREMIUM_HINT = "Premium plans (Project for the web) aren't available via the Graph Planner API";

/** Planner request; a 403/404 on a plan or task may mean a Premium plan, so say so. */
async function plannerRequest(client: GraphClient, endpoint: string, options?: GraphRequestOptions) {
  try {
    return await client.request(endpoint, options);
  } catch (err) {
    if (err instanceof CliError && (err.status === 403 || err.status === 404) && /^\/planner\/(plans|tasks)\b/.test(endpoint)) {
      throw new CliError(err.message, {
        code: err.code, exitCode: err.exitCode, status: err.status,
        graphCode: err.graphCode, requestId: err.requestId, fix: PREMIUM_HINT,
      });
    }
    throw err;
  }
}

/** Fetch a Planner resource's ETag (body @odata.etag, else the ETag header) for If-Match. */
async function autoFetchETag(client: GraphClient, endpoint: string): Promise<string> {
  const response = await plannerRequest(client, endpoint);
  const data = response.data as Record<string, unknown> | undefined;
  if (data?.dryRun) return 'W/"dry-run"';
  const etag = data?.['@odata.etag'] ?? response.headers?.get('ETag');
  if (typeof etag === 'string' && etag) return etag;
  throw new CliError(`Could not auto-fetch ETag for ${endpoint}. Provide --etag manually.`);
}

/** Planner dueDateTime is a UTC instant; resolve the user's input (date = local midnight). */
function plannerDue(input: string, tz?: string): string {
  return parseTime(input, { tz }).date.toISOString();
}

export function registerPlannerCommands(program: Command, getClient: () => GraphClient): void {
  const planner = program.command('planner').description('Microsoft Planner tasks');

  planner
    .command('plans')
    .description('List my planner plans')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/planner/plans');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('plan <planId>')
    .description('Get plan details')
    .addOption(outputOption())
    .action(handleErrors(async (planId: string, opts) => {
      const client = getClient();
      const response = await plannerRequest(client, `/planner/plans/${encodeURIComponent(planId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('plan-details <planId>')
    .description('Get plan extended details (category labels, etc.)')
    .addOption(outputOption())
    .action(handleErrors(async (planId: string, opts) => {
      const client = getClient();
      const response = await plannerRequest(client, `/planner/plans/${encodeURIComponent(planId)}/details`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('my-tasks')
    .description('List my planner tasks')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/planner/tasks');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('tasks <planId>')
    .description('List tasks in a plan')
    .addOption(outputOption())
    .action(handleErrors(async (planId: string, opts) => {
      const client = getClient();
      const response = await plannerRequest(client, `/planner/plans/${encodeURIComponent(planId)}/tasks`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('task <taskId>')
    .description('Get task details')
    .addOption(outputOption())
    .action(handleErrors(async (taskId: string, opts) => {
      const client = getClient();
      const response = await plannerRequest(client, `/planner/tasks/${encodeURIComponent(taskId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('task-details <taskId>')
    .description('Get task extended details (description, checklist, etc.)')
    .addOption(outputOption())
    .action(handleErrors(async (taskId: string, opts) => {
      const client = getClient();
      const response = await plannerRequest(client, `/planner/tasks/${encodeURIComponent(taskId)}/details`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('buckets <planId>')
    .description('List plan buckets')
    .addOption(outputOption())
    .action(handleErrors(async (planId: string, opts) => {
      const client = getClient();
      const response = await plannerRequest(client, `/planner/plans/${encodeURIComponent(planId)}/buckets`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('bucket <bucketId>')
    .description('Get bucket details')
    .addOption(outputOption())
    .action(handleErrors(async (bucketId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/planner/buckets/${encodeURIComponent(bucketId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('bucket-tasks <bucketId>')
    .description('List tasks in a bucket')
    .addOption(outputOption())
    .action(handleErrors(async (bucketId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/planner/buckets/${encodeURIComponent(bucketId)}/tasks`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('create-task')
    .description('Create a planner task')
    .requiredOption('--plan <planId>', 'Plan ID')
    .requiredOption('--title <title>', 'Task title')
    .option('--bucket <bucketId>', 'Bucket ID')
    .option('--due <date>', 'Due date (YYYY-MM-DD, "tomorrow 17:00", "next fri", "+3d")')
    .option('--timezone <tz>', 'IANA time zone for --due (default: MS365_CLI_TIMEZONE or local)')
    .option('--assigned-to <userId>', 'Assign to user ID')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const task: Record<string, unknown> = {
        planId: opts.plan,
        title: opts.title,
      };
      if (opts.bucket) task.bucketId = opts.bucket;
      if (opts.due) task.dueDateTime = plannerDue(opts.due, opts.timezone);
      if (opts.assignedTo) {
        task.assignments = { [opts.assignedTo]: { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' } };
      }

      const response = await plannerRequest(client, '/planner/tasks', {
        method: 'POST',
        body: JSON.stringify(task),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('update-task <taskId>')
    .description('Update a planner task')
    .option('--title <title>', 'Task title')
    .option('--due <date>', 'Due date (YYYY-MM-DD, "tomorrow 17:00", "next fri", "+3d")')
    .option('--timezone <tz>', 'IANA time zone for --due (default: MS365_CLI_TIMEZONE or local)')
    .option('--percent <n>', 'Percent complete (0, 50, 100)')
    .option('--bucket <bucketId>', 'Move to bucket')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .addOption(outputOption())
    .action(handleErrors(async (taskId: string, opts) => {
      const client = getClient();
      const patch: Record<string, unknown> = {};
      if (opts.title) patch.title = opts.title;
      if (opts.due) patch.dueDateTime = plannerDue(opts.due, opts.timezone);
      if (opts.percent !== undefined) {
        const pct = parseInt(opts.percent);
        if (![0, 50, 100].includes(pct)) {
          throw new Error('--percent must be 0, 50, or 100');
        }
        patch.percentComplete = pct;
      }
      if (opts.bucket) patch.bucketId = opts.bucket;

      const response = await plannerRequest(client, `/planner/tasks/${encodeURIComponent(taskId)}`, {
        method: 'PATCH',
        body: JSON.stringify(patch),
        headers: { 'If-Match': opts.etag || await autoFetchETag(client, `/planner/tasks/${encodeURIComponent(taskId)}`) },
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('update-plan-details <planId>')
    .description('Update plan details (category labels)')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .option('--categories <json>', 'Category descriptions as JSON')
    .addOption(outputOption())
    .action(handleErrors(async (planId: string, opts) => {
      const client = getClient();
      const patch: Record<string, unknown> = {};
      if (opts.categories) {
        patch.categoryDescriptions = JSON.parse(opts.categories);
      }

      const response = await plannerRequest(client, `/planner/plans/${encodeURIComponent(planId)}/details`, {
        method: 'PATCH',
        body: JSON.stringify(patch),
        headers: { 'If-Match': opts.etag || await autoFetchETag(client, `/planner/plans/${encodeURIComponent(planId)}/details`) },
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('update-task-details <taskId>')
    .description('Update task details (description, checklist)')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .option('--description <text>', 'Task description')
    .addOption(outputOption())
    .action(handleErrors(async (taskId: string, opts) => {
      const client = getClient();
      const patch: Record<string, unknown> = {};
      if (opts.description) patch.description = opts.description;

      const response = await plannerRequest(client, `/planner/tasks/${encodeURIComponent(taskId)}/details`, {
        method: 'PATCH',
        body: JSON.stringify(patch),
        headers: { 'If-Match': opts.etag || await autoFetchETag(client, `/planner/tasks/${encodeURIComponent(taskId)}/details`) },
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('delete-task <taskId>')
    .description('Delete a planner task')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .action(handleErrors(async (taskId: string, opts) => {
      if (!await confirm(`Delete planner task ${taskId.substring(0, 16)}...?`)) return;
      const client = getClient();
      const etag = opts.etag || await autoFetchETag(client, `/planner/tasks/${encodeURIComponent(taskId)}`);
      await plannerRequest(client, `/planner/tasks/${encodeURIComponent(taskId)}`, {
        method: 'DELETE',
        headers: { 'If-Match': etag },
      });
      info('Task deleted.');
    }));

  planner
    .command('delete-plan <planId>')
    .description('Delete a planner plan')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .action(handleErrors(async (planId: string, opts) => {
      if (!await confirm(`Delete planner plan ${planId.substring(0, 16)}...?`)) return;
      const client = getClient();
      const etag = opts.etag || await autoFetchETag(client, `/planner/plans/${encodeURIComponent(planId)}`);
      await plannerRequest(client, `/planner/plans/${encodeURIComponent(planId)}`, {
        method: 'DELETE',
        headers: { 'If-Match': etag },
      });
      info('Plan deleted.');
    }));

  planner
    .command('delete-bucket <bucketId>')
    .description('Delete a planner bucket')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .action(handleErrors(async (bucketId: string, opts) => {
      if (!await confirm(`Delete planner bucket ${bucketId.substring(0, 16)}...?`)) return;
      const client = getClient();
      const etag = opts.etag || await autoFetchETag(client, `/planner/buckets/${encodeURIComponent(bucketId)}`);
      await client.request(`/planner/buckets/${encodeURIComponent(bucketId)}`, {
        method: 'DELETE',
        headers: { 'If-Match': etag },
      });
      info('Bucket deleted.');
    }));

  planner
    .command('update-plan <planId>')
    .description('Update a planner plan')
    .option('--title <title>', 'Plan title')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .addOption(outputOption())
    .action(handleErrors(async (planId: string, opts) => {
      const client = getClient();
      const patch: Record<string, unknown> = {};
      if (opts.title) patch.title = opts.title;

      const response = await plannerRequest(client, `/planner/plans/${encodeURIComponent(planId)}`, {
        method: 'PATCH',
        body: JSON.stringify(patch),
        headers: { 'If-Match': opts.etag || await autoFetchETag(client, `/planner/plans/${encodeURIComponent(planId)}`) },
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('update-bucket <bucketId>')
    .description('Update a planner bucket')
    .option('--name <name>', 'Bucket name')
    .option('--etag <etag>', 'ETag for concurrency (auto-fetched if omitted)')
    .addOption(outputOption())
    .action(handleErrors(async (bucketId: string, opts) => {
      const client = getClient();
      const patch: Record<string, unknown> = {};
      if (opts.name) patch.name = opts.name;

      const response = await client.request(`/planner/buckets/${encodeURIComponent(bucketId)}`, {
        method: 'PATCH',
        body: JSON.stringify(patch),
        headers: { 'If-Match': opts.etag || await autoFetchETag(client, `/planner/buckets/${encodeURIComponent(bucketId)}`) },
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  planner
    .command('create-plan')
    .description('Create a planner plan')
    .requiredOption('--title <title>', 'Plan title')
    .requiredOption('--group <groupId>', 'Group/Team ID (owner)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/planner/plans', {
        method: 'POST',
        body: JSON.stringify({
          title: opts.title,
          container: { url: `https://graph.microsoft.com/v1.0/groups/${encodeURIComponent(opts.group)}` },
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
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/planner/buckets', {
        method: 'POST',
        body: JSON.stringify({
          name: opts.name,
          planId: opts.plan,
        }),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));
}
