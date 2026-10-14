/**
 * Regenerate the README command reference from `ms365 describe --json`.
 *   tsx scripts/gen-docs.ts             rewrite the generated region
 *   tsx scripts/gen-docs.ts --check     exit 1 with a diff summary when the README is stale
 *   --readme <path>                     README to read/write (default: README.md at the repo root)
 */
import fs from 'fs';
import path from 'path';
import { execFileSync } from 'child_process';
import { fileURLToPath } from 'url';
import { parseArgs } from 'util';
import { type Catalog, replaceGeneratedRegion, checkGeneratedRegion } from '../src/docs-gen.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { values } = parseArgs({ options: { check: { type: 'boolean' }, readme: { type: 'string' } } });
const readmePath = path.resolve(values.readme ?? path.join(root, 'README.md'));

// Run the real CLI (piped, so not a TTY) so the reference matches the registered command tree exactly.
const catalog = JSON.parse(execFileSync(process.execPath, ['--import', 'tsx', path.join(root, 'src', 'index.ts'), 'describe', '--json'], {
  cwd: root,
  encoding: 'utf8',
  maxBuffer: 64 * 1024 * 1024,
})) as Catalog;
const readme = fs.readFileSync(readmePath, 'utf8');

if (values.check) {
  const { ok, summary } = checkGeneratedRegion(readme, catalog);
  if (!ok) {
    process.stderr.write(`${path.relative(root, readmePath) || readmePath} command reference is stale. Run \`npm run docs\` and commit the result.\n${summary}\n`);
    process.exit(1);
  }
  process.stderr.write('README command reference is up to date.\n');
} else {
  fs.writeFileSync(readmePath, replaceGeneratedRegion(readme, catalog));
  process.stderr.write(`Wrote ${catalog.commands.length} commands to ${readmePath}\n`);
}
