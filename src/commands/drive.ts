import { Command, InvalidArgumentError, Option } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat, printList } from '../formatter.js';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import { detectMimeType } from '../mime.js';
import { formatSize, addListOptions, outputOption, confirm, info, success, encodePath, odataQuote, safeOutputPath } from '../utils.js';
import fs from 'fs';
import path from 'path';
import { Readable } from 'stream';
import { once } from 'events';
import { pipeline } from 'stream/promises';
import type { ReadableStream as WebReadableStream } from 'stream/web';

/** Fetch a /content endpoint and stream its body to `target` (a file path) or stdout. */
async function download(client: GraphClient, endpoint: string, target: string | null): Promise<void> {
  const response = await client.request(endpoint, { responseType: 'stream' });
  if (client.isPreview(response)) {
    console.log(formatOutput(response.data, 'json'));
    return;
  }
  // A body-less answer (204) saves an empty file.
  const body = response.stream ? Readable.fromWeb(response.stream as WebReadableStream<Uint8Array>) : Readable.from([]);
  if (target === null) {
    for await (const chunk of body) {
      if (!process.stdout.write(chunk)) await once(process.stdout, 'drain');
    }
    return;
  }
  try {
    await pipeline(body, fs.createWriteStream(target));
  } catch (err) {
    fs.rmSync(target, { force: true });
    throw err;
  }
  info(`Downloaded to ${target} (${formatSize(fs.statSync(target).size)})`);
}

/** Graph shares API: a sharing URL is addressed as "u!" + unpadded base64url of the URL. */
function shareToken(input: string): string {
  return /^https?:\/\//i.test(input)
    ? 'u!' + Buffer.from(input).toString('base64url')
    : encodeURIComponent(input);
}

/** The --password-stdin secret. An MCP-spawned child has stdin on /dev/null, and an empty password is never meant (X-15). */
async function readPasswordStdin(): Promise<string> {
  const usage = (msg: string) => new CliError(msg, {
    code: 'usage', exitCode: ExitCode.USAGE, fix: 'Pipe the password in, e.g. printf %s "$PW" | ms365 drive share … --password-stdin',
  });
  if (process.env.MS365_CLI_MCP) throw usage('--password-stdin is not available over MCP (stdin is not readable)');
  const chunks: Buffer[] = [];
  for await (const chunk of process.stdin) chunks.push(Buffer.from(chunk));
  const password = Buffer.concat(chunks).toString('utf8').replace(/\r?\n$/, '');
  if (!password) throw usage('--password-stdin read an empty password from stdin');
  return password;
}

export function registerDriveCommands(program: Command, getClient: () => GraphClient): void {
  const drive = program.command('drive').description('OneDrive file operations');

  addListOptions(drive
    .command('list')
    .description('List available drives'))
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      await printList(getClient(), '/me/drives', {}, opts);
    }));

  drive
    .command('root <driveId>')
    .description('Get root folder of a drive')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/drives/${encodeURIComponent(driveId)}/root`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  addListOptions(drive
    .command('files <driveId> <folderId>')
    .description('List files in a folder')
    .option('--filter <filter>', 'OData filter')
    .option('--select <fields>', 'Fields to return'), { description: 'Number of items' })
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, folderId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.select) queryParams['$select'] = opts.select;

      const endpoint = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folderId)}/children`;
      await printList(client, endpoint, { queryParams }, opts);
    }));

  drive
    .command('download <driveId> <itemId> <childId>')
    .description('Download a file')
    .option('--out <path>', 'Output file path')
    .option('--force', 'Overwrite an existing file')
    .action(handleErrors(async (driveId: string, itemId: string, childId: string, opts) => {
      const target = safeOutputPath(`download_${childId}`, { out: opts.out, force: opts.force });
      const endpoint = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children/${encodeURIComponent(childId)}/content`;
      await download(getClient(), endpoint, target);
    }));

  drive
    .command('upload <driveId> <parentId> <filePath>')
    .description('Upload a file (auto-detects MIME type, streams large files)')
    .option('--content-type <mime>', 'Override MIME type')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, parentId: string, filePath: string, opts) => {
      const resolvedPath = path.resolve(filePath);
      if (!fs.existsSync(resolvedPath)) {
        throw new CliError(`File not found: ${resolvedPath}`, { code: 'usage', exitCode: ExitCode.USAGE });
      }

      const client = getClient();
      const fileName = path.basename(resolvedPath);
      const contentType = opts.contentType || detectMimeType(fileName);
      const stat = fs.statSync(resolvedPath);

      console.error(`Uploading ${fileName} (${formatSize(stat.size)}, ${contentType})`);

      // Use colon-based path syntax — colons are delimiters, not encoded
      const endpoint = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}:/${encodeURIComponent(fileName)}:/content`;
      const response = await client.uploadFile(endpoint, resolvedPath, contentType);

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  drive
    .command('delete <driveId> <itemId>')
    .description('Delete a drive item')
    .action(handleErrors(async (driveId: string, itemId: string) => {
      if (!await confirm(`Delete drive item ${itemId.substring(0, 16)}...?`)) return;
      const client = getClient();
      await client.request(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`, {
        method: 'DELETE',
      });
      success('Item deleted.');
    }));

  drive
    .command('download-shared <shareId>')
    .description('Download a file via share ID or sharing link URL')
    .option('--out <path>', 'Output file path')
    .option('--force', 'Overwrite an existing file')
    .action(handleErrors(async (shareId: string, opts) => {
      const target = safeOutputPath('shared_download', { out: opts.out, force: opts.force });
      await download(getClient(), `/shares/${shareToken(shareId)}/driveItem/content`, target);
    }));
  // ── Search ───────────────────────────────────────────────────
  addListOptions(drive
    .command('search <query>')
    .description('Search files across OneDrive')
    .option('--drive <driveId>', 'Limit to specific drive'), { description: 'Number of results', default: '25' })
    .addOption(outputOption())
    .action(handleErrors(async (query: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = { '$top': opts.top };

      // The term sits inside an OData function call in the path: escape quotes, then percent-encode.
      const safeQuery = encodeURIComponent(odataQuote(query));
      const endpoint = opts.drive
        ? `/drives/${encodeURIComponent(opts.drive)}/root/search(q='${safeQuery}')`
        : `/me/drive/root/search(q='${safeQuery}')`;

      await printList(client, endpoint, { queryParams }, opts);
    }));

  // ── Create folder ───────────────────────────────────────────
  drive
    .command('mkdir <driveId> <parentId> <name>')
    .description('Create a folder')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, parentId: string, name: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}/children`,
        {
          method: 'POST',
          body: JSON.stringify({
            name,
            folder: {},
            '@microsoft.graph.conflictBehavior': 'rename',
          }),
        }
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // ── Sharing links ───────────────────────────────────────────
  drive
    .command('share <driveId> <itemId>')
    .description('Create a sharing link for a file or folder')
    .addOption(new Option('--type <type>', 'Link type').choices(['view', 'edit', 'embed']).default('view'))
    .addOption(new Option('--scope <scope>', 'Link scope').choices(['anonymous', 'organization', 'users']).default('organization'))
    .option('--password-stdin', 'Password-protect the link; the password is read from stdin (never from argv)')
    .option('--expiry <datetime>', 'Expiration datetime (ISO 8601)')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, opts) => {
      const client = getClient();
      const body: Record<string, unknown> = {
        type: opts.type,
        scope: opts.scope,
      };
      if (opts.passwordStdin) {
        body.password = await readPasswordStdin();
      }
      if (opts.expiry) body.expirationDateTime = opts.expiry;

      const response = await client.request(
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/createLink`,
        { method: 'POST', body: JSON.stringify(body) }
      );

      const data = response.data as Record<string, unknown>;
      const link = data.link as Record<string, unknown>;
      if (link?.webUrl && opts.output === 'text') {
        console.log(`Sharing link (${opts.type}, ${opts.scope}):`);
        console.log(`  ${link.webUrl}`);
      } else {
        console.log(formatOutput(response.data, opts.output as OutputFormat));
      }
    }));

  // ── Permissions ─────────────────────────────────────────────
  addListOptions(drive
    .command('permissions <driveId> <itemId>')
    .description('List sharing permissions on a file or folder'))
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, opts) => {
      const endpoint = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/permissions`;
      await printList(getClient(), endpoint, {}, opts);
    }));

  // ── Quota / storage usage ───────────────────────────────────
  drive
    .command('quota')
    .description('Show OneDrive storage usage and quota')
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/drive', {
        queryParams: { '$select': 'quota' },
      });
      const data = response.data as Record<string, unknown>;
      const quota = data.quota as Record<string, unknown>;

      if (opts.output === 'json' || opts.output === 'yaml') {
        console.log(formatOutput(data, opts.output as OutputFormat));
      } else {
        const used = quota?.used as number || 0;
        const total = quota?.total as number || 0;
        const remaining = quota?.remaining as number || 0;
        const deleted = quota?.deleted as number || 0;
        const pct = total > 0 ? Math.round((used / total) * 100) : 0;

        console.log('OneDrive Storage:');
        console.log(`  Used:      ${formatSize(used)} (${pct}%)`);
        console.log(`  Remaining: ${formatSize(remaining)}`);
        console.log(`  Total:     ${formatSize(total)}`);
        if (deleted > 0) console.log(`  In trash:  ${formatSize(deleted)}`);
      }
    }));

  // ── Get item info ───────────────────────────────────────────
  drive
    .command('info <driveId> <itemId>')
    .description('Get metadata for a file or folder')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // ── Path-based shortcuts ──────────────────────────────────
  addListOptions(drive
    .command('ls [path]')
    .description('List files by path (e.g. "/Documents", "/Documents/Reports")'), { description: 'Number of items' })
    .addOption(outputOption())
    .action(handleErrors(async (drivePath: string | undefined, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;

      const endpoint = drivePath
        ? `/me/drive/root:/${encodePath(drivePath.replace(/^\//, ''))}:/children`
        : '/me/drive/root/children';
      await printList(client, endpoint, { queryParams }, opts);
    }));

  drive
    .command('cat <path>')
    .description('Download a file by path (e.g. "/Documents/notes.md")')
    .option('--out <file>', 'Save to file instead of stdout')
    .option('--force', 'Overwrite an existing --out file')
    .action(handleErrors(async (drivePath: string, opts) => {
      const target = opts.out ? safeOutputPath('', { out: opts.out, force: opts.force }) : null;
      const cleanPath = encodePath(drivePath.replace(/^\//, ''));
      await download(getClient(), `/me/drive/root:/${cleanPath}:/content`, target);
    }));

  drive
    .command('cp <localFile> <remotePath>')
    .description('Upload a local file by path (e.g. ./report.pdf /Documents/report.pdf); to copy a drive item server-side, see "drive copy"')
    .option('--content-type <mime>', 'Override MIME type')
    .addOption(outputOption())
    .action(handleErrors(async (localFile: string, remotePath: string, opts) => {
      const resolvedPath = path.resolve(localFile);
      if (!fs.existsSync(resolvedPath)) {
        throw new CliError(`File not found: ${resolvedPath}`, { code: 'usage', exitCode: ExitCode.USAGE });
      }
      const client = getClient();
      const contentType = opts.contentType || detectMimeType(path.basename(resolvedPath));
      const cleanPath = encodePath(remotePath.replace(/^\//, ''));
      const endpoint = `/me/drive/root:/${cleanPath}:/content`;
      const response = await client.uploadFile(endpoint, resolvedPath, contentType);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // ── Server-side copy ────────────────────────────────────────
  drive
    .command('copy <driveId> <itemId>')
    .description('Copy a drive item server-side (Graph does the copy; an existing name fails, nothing is overwritten). To upload a local file, see "drive cp"')
    .requiredOption('--to-parent <id>', 'Destination folder ID')
    .option('--to-drive <id>', 'Destination drive ID (default: the source drive)')
    .option('--name <name>', 'Name of the copy (default: the original name)')
    .option('--wait', 'Wait for the copy to finish, then print the new item')
    .option('--timeout <s>', 'Seconds --wait waits before giving up', (v: string) => {
      if (!/^\d+$/.test(v) || Number(v) < 1) throw new InvalidArgumentError('must be an integer ≥ 1.');
      return v;
    }, '300')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, opts) => {
      const client = getClient();
      const format = opts.output as OutputFormat;
      const toDrive: string = opts.toDrive ?? driveId;
      const body: Record<string, unknown> = { parentReference: { driveId: toDrive, id: opts.toParent } };
      if (opts.name) body.name = opts.name;
      // 202 Accepted + Location: a monitor URL for the async job †.
      const response = await client.request(
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/copy`,
        { method: 'POST', body: JSON.stringify(body) }
      );
      if (client.isPreview(response)) {
        console.log(formatOutput(response.data, format));
        return;
      }
      const monitorUrl = response.headers?.get('location');
      if (!monitorUrl) {
        throw new CliError('Graph accepted the copy but returned no monitor URL (Location header)', { code: 'copy_failed' });
      }
      if (!opts.wait) {
        console.log(formatOutput({ status: 'accepted', monitorUrl }, format));
        return;
      }
      const result = await client.pollMonitor(monitorUrl, Number(opts.timeout) * 1000);
      if (result.status === 'failed') {
        const error = result.body.error as { code?: string; message?: string } | undefined;
        throw new CliError(`Copy failed: ${error?.message ?? error?.code ?? 'the monitor reported status "failed"'}`, { code: 'copy_failed' });
      }
      if (!result.resourceId) {
        throw new CliError('Copy completed, but the monitor did not name the new item', {
          code: 'copy_failed', fix: `List the destination: ms365 drive files ${toDrive} ${opts.toParent}`,
        });
      }
      const item = await client.request(`/drives/${encodeURIComponent(toDrive)}/items/${encodeURIComponent(result.resourceId)}`);
      console.log(formatOutput(item.data, format));
    }));
}
