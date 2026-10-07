import { Command, Option } from 'commander';
import { GraphClient, type ConflictBehavior } from '../graph-client.js';
import { graph } from '../command-tree.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { handleErrors } from '../errors.js';
import { detectMimeType } from '../mime.js';
import { saveOrStream } from '../files.js';
import { drivePath, field, odataQuote } from '../graph-helpers.js';
import { listRequest, paginationOptions, type ListOptions } from '../list.js';
import { formatSize, outputOption, confirm, done, info } from '../utils.js';
import fs from 'fs';
import path from 'path';

interface OutputOpts { output: string }
interface SaveOpts { out?: string; force?: boolean }
interface UploadOpts extends OutputOpts { contentType?: string; conflict: ConflictBehavior }
interface FilesOpts extends ListOptions { filter?: string; select?: string }
interface SearchOpts extends ListOptions { drive?: string }
interface ShareOpts extends OutputOpts { type: string; scope: string; password?: string; expiry?: string }

function conflictOption(defaultValue: ConflictBehavior): Option {
  return new Option('--conflict <behavior>', 'If the name exists: rename, replace, fail')
    .choices(['rename', 'replace', 'fail'])
    .default(defaultValue);
}

function saveOptions(cmd: Command, outHelp: string): Command {
  return cmd
    .option('--out <path>', outHelp)
    .option('--force', 'Overwrite an existing --out file');
}

/** Download `endpoint` unparsed to `--out` (no overwrite unless --force) or to stdout. */
async function download(client: GraphClient, endpoint: string, opts: SaveOpts): Promise<void> {
  const response = await client.requestBinary(endpoint);
  const saved = await saveOrStream(response.body, { out: opts.out, force: opts.force });
  if (saved.path) info(`Saved to ${path.resolve(saved.path)} (${formatSize(saved.bytes)})`);
}

/** Upload a local file to `itemPath` (colon path without trailing colon) and print the stored driveItem. */
async function upload(client: GraphClient, itemPath: string, localPath: string, opts: UploadOpts): Promise<void> {
  const contentType = opts.contentType || detectMimeType(path.basename(localPath));
  if (!client.isDryRun) info(`Uploading ${path.basename(localPath)} (${formatSize(fs.statSync(localPath).size)}, ${contentType})`);
  const response = await client.uploadDriveItem(itemPath, localPath, { conflict: opts.conflict, contentType });
  console.log(formatOutput(response.data, opts.output as OutputFormat));
}

function existingFile(filePath: string): string {
  const resolved = path.resolve(filePath);
  if (!fs.existsSync(resolved)) throw new Error(`File not found: ${resolved}`);
  return resolved;
}

export function registerDriveCommands(program: Command, getClient: () => GraphClient): void {
  const drive = program.command('drive').description('OneDrive file operations');
  const itemRoot = (driveId: string, itemId: string) =>
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`;

  paginationOptions(graph(drive.command('list'), ['GET /me/drives'])
    .description('List available drives'))
    .addOption(outputOption())
    .action(handleErrors(async (opts: ListOptions) => {
      await listRequest(getClient(), '/me/drives', {}, opts);
    }));

  graph(drive.command('root <driveId>'), ['GET /drives/{drive-id}/root'])
    .description('Get root folder of a drive')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, opts: OutputOpts) => {
      const response = await getClient().request(`/drives/${encodeURIComponent(driveId)}/root`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  paginationOptions(graph(drive.command('files <driveId> <folderId>'), ['GET /drives/{drive-id}/items/{item-id}/children'])
    .description('List files in a folder')
    .option('--filter <filter>', 'OData filter')
    .option('--select <fields>', 'Fields to return'))
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, folderId: string, opts: FilesOpts) => {
      const queryParams: Record<string, string> = {};
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.select) queryParams['$select'] = opts.select;
      await listRequest(getClient(), `${itemRoot(driveId, folderId)}/children`, { queryParams }, opts);
    }));

  saveOptions(graph(drive.command('download <driveId> <itemId>'), ['GET /drives/{drive-id}/items/{item-id}/content'])
    .description('Download a file (to --out, or raw bytes to stdout)'), 'Output file path (default: stdout)')
    .allowExcessArguments(false)
    .action(handleErrors(async (driveId: string, itemId: string, opts: SaveOpts) => {
      await download(getClient(), `${itemRoot(driveId, itemId)}/content`, opts);
    }));

  graph(drive.command('upload <driveId> <parentId> <filePath>'), [
    'PUT /drives/{drive-id}/items/{parent-id}:/{filename}:/content',
    'POST /drives/{drive-id}/items/{parent-id}:/{filename}:/createUploadSession',
  ])
    .description('Upload a file into a folder (auto-detects MIME type, streams large files)')
    .option('--content-type <mime>', 'Override MIME type')
    .addOption(conflictOption('rename'))
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, parentId: string, filePath: string, opts: UploadOpts) => {
      const localPath = existingFile(filePath);
      const itemPath = drivePath(itemRoot(driveId, parentId), path.basename(localPath));
      await upload(getClient(), itemPath, localPath, opts);
    }));

  graph(drive.command('delete <driveId> <itemId>'), ['DELETE /drives/{drive-id}/items/{item-id}'])
    .description('Delete a drive item')
    .action(handleErrors(async (driveId: string, itemId: string) => {
      await confirm(`Delete drive item ${itemId.substring(0, 16)}...?`);
      done(await getClient().request(itemRoot(driveId, itemId), { method: 'DELETE' }), 'Item deleted.');
    }));

  saveOptions(graph(drive.command('download-shared <shareId>'), ['GET /shares/{share-id}/driveItem/content'])
    .description('Download a file via sharing link (to --out, or raw bytes to stdout)'), 'Output file path (default: stdout)')
    .action(handleErrors(async (shareId: string, opts: SaveOpts) => {
      await download(getClient(), `/shares/${encodeURIComponent(shareId)}/driveItem/content`, opts);
    }));

  // ── Search ───────────────────────────────────────────────────
  paginationOptions(graph(drive.command('search <query>'), [
    "GET /me/drive/root/search(q='{query}')",
    "GET /drives/{drive-id}/root/search(q='{query}')",
  ])
    .description('Search files across OneDrive')
    .option('--drive <driveId>', 'Limit to specific drive'), 25)
    .addOption(outputOption())
    .action(handleErrors(async (query: string, opts: SearchOpts) => {
      // The query is an OData string literal inside the path: quote it, then URL-encode so #, ?, % survive.
      const fn = `search(q=${encodeURIComponent(odataQuote(query))})`;
      const endpoint = opts.drive ? `/drives/${encodeURIComponent(opts.drive)}/root/${fn}` : `/me/drive/root/${fn}`;
      await listRequest(getClient(), endpoint, {}, opts);
    }));

  // ── Create folder ───────────────────────────────────────────
  graph(drive.command('mkdir <driveId> <parentId> <name>'), ['POST /drives/{drive-id}/items/{item-id}/children'])
    .description('Create a folder')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, parentId: string, name: string, opts: OutputOpts) => {
      const response = await getClient().request(`${itemRoot(driveId, parentId)}/children`, {
        method: 'POST',
        body: JSON.stringify({
          name,
          folder: {},
          '@microsoft.graph.conflictBehavior': 'rename',
        }),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // ── Sharing links ───────────────────────────────────────────
  graph(drive.command('share <driveId> <itemId>'), ['POST /drives/{drive-id}/items/{item-id}/createLink'])
    .description('Create a sharing link for a file or folder')
    .addOption(new Option('--type <type>', 'Link type').choices(['view', 'edit', 'embed']).default('view'))
    .addOption(new Option('--scope <scope>', 'Link scope').choices(['anonymous', 'organization', 'users']).default('organization'))
    .option('--password <password>', 'Password-protect the link')
    .option('--expiry <datetime>', 'Expiration datetime (ISO 8601)')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, opts: ShareOpts) => {
      const body: Record<string, unknown> = {
        type: opts.type,
        scope: opts.scope,
      };
      if (opts.password) body.password = opts.password;
      if (opts.expiry) body.expirationDateTime = opts.expiry;

      const response = await getClient().request(`${itemRoot(driveId, itemId)}/createLink`, {
        method: 'POST',
        body: JSON.stringify(body),
      });

      const webUrl = field(field(response.data, 'link'), 'webUrl');
      if (typeof webUrl === 'string' && (opts.output === 'text' || opts.output === 'table')) {
        console.log(`Sharing link (${opts.type}, ${opts.scope}):`);
        console.log(`  ${webUrl}`);
      } else {
        console.log(formatOutput(response.data, opts.output as OutputFormat));
      }
    }));

  // ── Permissions ─────────────────────────────────────────────
  graph(drive.command('permissions <driveId> <itemId>'), ['GET /drives/{drive-id}/items/{item-id}/permissions'])
    .description('List sharing permissions on a file or folder')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, opts: OutputOpts) => {
      await listRequest(getClient(), `${itemRoot(driveId, itemId)}/permissions`, {}, { output: opts.output, all: true });
    }));

  // ── Quota / storage usage ───────────────────────────────────
  graph(drive.command('quota'), ['GET /me/drive'])
    .description('Show OneDrive storage usage and quota')
    .addOption(outputOption('text'))
    .action(handleErrors(async (opts: OutputOpts) => {
      const response = await getClient().request('/me/drive', {
        queryParams: { '$select': 'quota' },
      });
      const quota = field(response.data, 'quota');

      if (opts.output !== 'text' || typeof quota !== 'object' || quota === null) {
        console.log(formatOutput(response.data, opts.output as OutputFormat));
        return;
      }
      const bytes = (key: string) => {
        const v = field(quota, key);
        return typeof v === 'number' ? v : 0;
      };
      const used = bytes('used');
      const total = bytes('total');
      const pct = total > 0 ? Math.round((used / total) * 100) : 0;
      console.log('OneDrive Storage:');
      console.log(`  Used:      ${formatSize(used)} (${pct}%)`);
      console.log(`  Remaining: ${formatSize(bytes('remaining'))}`);
      console.log(`  Total:     ${formatSize(total)}`);
      if (bytes('deleted') > 0) console.log(`  In trash:  ${formatSize(bytes('deleted'))}`);
    }));

  // ── Get item info ───────────────────────────────────────────
  graph(drive.command('info <driveId> <itemId>'), ['GET /drives/{drive-id}/items/{item-id}'])
    .description('Get metadata for a file or folder')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, opts: OutputOpts) => {
      const response = await getClient().request(itemRoot(driveId, itemId));
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  // ── Path-based shortcuts ──────────────────────────────────
  paginationOptions(graph(drive.command('ls [path]'), ['GET /me/drive/root/children', 'GET /me/drive/root:/{path}:/children'])
    .description('List files by path (e.g. "/Documents", "/Documents/Reports")'))
    .addOption(outputOption())
    .action(handleErrors(async (userPath: string | undefined, opts: ListOptions) => {
      await listRequest(getClient(), drivePath('/me/drive/root', userPath ?? '', '/children'), {}, opts);
    }));

  saveOptions(graph(drive.command('cat <path>'), ['GET /me/drive/root:/{path}:/content'])
    .description('Download a file by path (e.g. "/Documents/notes.md"); raw bytes to stdout'), 'Save to file instead of stdout')
    .action(handleErrors(async (userPath: string, opts: SaveOpts) => {
      await download(getClient(), drivePath('/me/drive/root', userPath, '/content'), opts);
    }));

  graph(drive.command('cp <localFile> <remotePath>'), [
    'PUT /me/drive/root:/{path}:/content',
    'POST /me/drive/root:/{path}:/createUploadSession',
  ])
    .description('Upload a file by path (e.g. ./report.pdf /Documents/report.pdf; a path ending in / is a folder)')
    .option('--content-type <mime>', 'Override MIME type')
    .addOption(conflictOption('replace'))
    .addOption(outputOption())
    .action(handleErrors(async (localFile: string, remotePath: string, opts: UploadOpts) => {
      const localPath = existingFile(localFile);
      const target = remotePath.endsWith('/') || remotePath === '' ? `${remotePath}/${path.basename(localPath)}` : remotePath;
      await upload(getClient(), drivePath('/me/drive/root', target), localPath, opts);
    }));
}
