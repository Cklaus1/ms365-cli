// Child process for the ledger race gate (R-6): claim a key, print the outcome ("new" or the error code), and hold
// the claim until <goFile> exists. Without <goFile> the claim is held until the process is killed.
import fs from 'fs';
import { withIdempotency } from '../../src/idempotency.js';
import { CliError } from '../../src/errors.js';

const [key, goFile] = process.argv.slice(2);
try {
  await withIdempotency(key, 'mail send', 'fp', async () => {
    process.stdout.write('new\n');
    while (!goFile || !fs.existsSync(goFile)) await new Promise(r => setTimeout(r, 25));
    return { sentAt: 'T' };
  });
} catch (err) {
  process.stdout.write(`${err instanceof CliError ? err.code : String(err)}\n`);
}
