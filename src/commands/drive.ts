import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { handleErrors } from '../errors.js';
import { detectMimeType } from '../mime.js';
import { formatSize, outputOption, confirm, info } from '../utils.js';
import fs from 'fs';
import path from 'path';

export function registerDriveCommands(program: Command, getClient: () => GraphClient): void {
  const drive = program.command('drive').description('OneDrive file operations');

  drive
    .command('list')
    .description('List available drives')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const response = await client.request('/me/drives');
      console.log(formatOutput(response.data, opts.output as OutputFormat));
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

  drive
    .command('files <driveId> <folderId>')
    .description('List files in a folder')
    .option('--top <n>', 'Number of items')
    .option('--filter <filter>', 'OData filter')
    .option('--select <fields>', 'Fields to return')
    .option('--all', 'Fetch all pages')
    .option('--page-limit <n>', 'Max pages when using --all')
    .option('--page-delay <ms>', 'Delay between pages in ms')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, folderId: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.select) queryParams['$select'] = opts.select;

      const endpoint = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folderId)}/children`;
      const response = opts.all
        ? await client.requestAllPages(
            endpoint,
            { queryParams },
            opts.pageLimit ? parseInt(opts.pageLimit) : undefined,
            opts.pageDelay ? parseInt(opts.pageDelay) : undefined,
          )
        : await client.request(endpoint, { queryParams });

      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  drive
    .command('download <driveId> <itemId> <childId>')
    .description('Download a file')
    .option('--out <path>', 'Output file path')
    .action(handleErrors(async (driveId: string, itemId: string, childId: string, opts) => {
      const client = getClient();
      const endpoint = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children/${encodeURIComponent(childId)}/content`;
      const response = await client.request(endpoint);
      const data = response.data as Record<string, unknown>;

      if (data.contentBytes) {
        const outPath = path.resolve(opts.out || `download_${childId}`);
        fs.writeFileSync(outPath, Buffer.from(data.contentBytes as string, 'base64'));
        console.log(`Downloaded to ${outPath} (${data.size} bytes)`);
      } else {
        console.log(formatOutput(response.data, 'json'));
      }
    }));

  drive
    .command('upload <driveId> <parentId> <filePath>')
    .description('Upload a file (auto-detects MIME type, streams large files)')
    .option('--content-type <mime>', 'Override MIME type')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, parentId: string, filePath: string, opts) => {
      const resolvedPath = path.resolve(filePath);
      if (!fs.existsSync(resolvedPath)) {
        throw new Error(`File not found: ${resolvedPath}`);
      }

      const client = getClient();
      const fileName = path.basename(resolvedPath);
      const contentType = opts.contentType || detectMimeType(fileName);
      const stat = fs.statSync(resolvedPath);

      console.error(`Uploading ${fileName} (${formatSize(stat.size)}, ${contentType})`);

      // Use colon-based path syntax — colons are delimiters, not encoded
      const endpoint = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}:/${fileName}:/content`;
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
      info('Item deleted.');
    }));

  drive
    .command('download-shared <shareId>')
    .description('Download a file via sharing link')
    .option('--out <path>', 'Output file path')
    .action(handleErrors(async (shareId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/shares/${encodeURIComponent(shareId)}/driveItem/content`
      );
      const data = response.data as Record<string, unknown>;

      if (data.contentBytes) {
        const outPath = path.resolve(opts.out || `shared_download`);
        fs.writeFileSync(outPath, Buffer.from(data.contentBytes as string, 'base64'));
        console.log(`Downloaded to ${outPath} (${data.size} bytes)`);
      } else {
        console.log(formatOutput(response.data, 'json'));
      }
    }));
  // ── Search ───────────────────────────────────────────────────
  drive
    .command('search <query>')
    .description('Search files across OneDrive')
    .option('--drive <driveId>', 'Limit to specific drive')
    .option('--top <n>', 'Number of results', '25')
    .addOption(outputOption())
    .action(handleErrors(async (query: string, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = { '$top': opts.top };

      // Don't encodeURIComponent the query — it's part of the OData function URL,
      // and buildUrl will encode query params separately. Single quotes are escaped.
      const safeQuery = query.replace(/'/g, "''");
      const endpoint = opts.drive
        ? `/drives/${encodeURIComponent(opts.drive)}/root/search(q='${safeQuery}')`
        : `/me/drive/root/search(q='${safeQuery}')`;

      const response = await client.request(endpoint, { queryParams });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
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
    .option('--type <type>', 'Link type: view, edit, embed', 'view')
    .option('--scope <scope>', 'Link scope: anonymous, organization, users', 'organization')
    .option('--password <password>', 'Password-protect the link')
    .option('--expiry <datetime>', 'Expiration datetime (ISO 8601)')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, opts) => {
      const client = getClient();
      const body: Record<string, unknown> = {
        type: opts.type,
        scope: opts.scope,
      };
      if (opts.password) body.password = opts.password;
      if (opts.expiry) body.expirationDateTime = opts.expiry;

      const response = await client.request(
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/createLink`,
        { method: 'POST', body: JSON.stringify(body) }
      );

      const data = response.data as Record<string, unknown>;
      const link = data.link as Record<string, unknown>;
      if (link?.webUrl && opts.output !== 'json') {
        console.log(`Sharing link (${opts.type}, ${opts.scope}):`);
        console.log(`  ${link.webUrl}`);
      } else {
        console.log(formatOutput(response.data, opts.output as OutputFormat));
      }
    }));

  // ── Permissions ─────────────────────────────────────────────
  drive
    .command('permissions <driveId> <itemId>')
    .description('List sharing permissions on a file or folder')
    .addOption(outputOption())
    .action(handleErrors(async (driveId: string, itemId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/permissions`
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
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
  drive
    .command('ls [path]')
    .description('List files by path (e.g. "/Documents", "/Documents/Reports")')
    .option('--top <n>', 'Number of items')
    .addOption(outputOption())
    .action(handleErrors(async (drivePath: string | undefined, opts) => {
      const client = getClient();
      const queryParams: Record<string, string> = {};
      if (opts.top) queryParams['$top'] = opts.top;

      const endpoint = drivePath
        ? `/me/drive/root:/${drivePath.replace(/^\//, '')}:/children`
        : '/me/drive/root/children';
      const response = await client.request(endpoint, { queryParams });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  drive
    .command('cat <path>')
    .description('Download a file by path (e.g. "/Documents/notes.md")')
    .option('--out <file>', 'Save to file instead of stdout')
    .action(handleErrors(async (drivePath: string, opts) => {
      const client = getClient();
      const cleanPath = drivePath.replace(/^\//, '');
      const response = await client.request(`/me/drive/root:/${cleanPath}:/content`);
      const data = response.data as Record<string, unknown>;

      if (data.contentBytes) {
        const buf = Buffer.from(data.contentBytes as string, 'base64');
        if (opts.out) {
          const outPath = (await import('path')).resolve(opts.out);
          (await import('fs')).writeFileSync(outPath, buf);
          info(`Saved to ${outPath} (${formatSize(buf.length)})`);
        } else {
          process.stdout.write(buf);
        }
      } else {
        console.log(formatOutput(response.data, 'json'));
      }
    }));

  drive
    .command('cp <localFile> <remotePath>')
    .description('Upload a file by path (e.g. ./report.pdf /Documents/report.pdf)')
    .option('--content-type <mime>', 'Override MIME type')
    .addOption(outputOption())
    .action(handleErrors(async (localFile: string, remotePath: string, opts) => {
      const resolvedPath = path.resolve(localFile);
      if (!fs.existsSync(resolvedPath)) {
        throw new Error(`File not found: ${resolvedPath}`);
      }
      const client = getClient();
      const contentType = opts.contentType || detectMimeType(path.basename(resolvedPath));
      const cleanPath = remotePath.replace(/^\//, '');
      const endpoint = `/me/drive/root:/${cleanPath}:/content`;
      const response = await client.uploadFile(endpoint, resolvedPath, contentType);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));
}

