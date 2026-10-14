import fs from 'fs';
import { Command } from 'commander';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import { outputOption } from '../utils.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { auditLogPath, readAuditTail, parseLines } from '../audit.js';

/** `ms365 audit` — read the local log of non-GET Graph calls. Offline: no sign-in needed. */
export function registerAuditCommands(program: Command): void {
  const audit = program.command('audit').description('Inspect the local audit log of every write sent to Graph');

  audit
    .command('tail')
    .description('Print the most recent audit entries')
    .option('-n, --lines <n>', 'Number of entries', '20')
    .option('--follow', 'Keep printing new entries as they are written (Ctrl+C to stop)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const n = Number(opts.lines);
      if (!Number.isInteger(n) || n < 0) {
        throw new CliError(`Invalid -n: ${opts.lines} (expected a non-negative integer)`, { code: 'invalid_argument', exitCode: ExitCode.USAGE });
      }
      const format = opts.output as OutputFormat;
      console.log(formatOutput(readAuditTail(n), format));
      if (opts.follow) await follow(format === 'json' ? 'ndjson' : format);
    }));

  audit
    .command('path')
    .description('Print the audit log path')
    .action(() => {
      console.log(auditLogPath());
    });
}

/** Poll the log and print entries appended after now; restarts from 0 after rotation. */
function follow(format: OutputFormat): Promise<never> {
  const file = auditLogPath();
  let offset = sizeOf(file);
  return new Promise(() => {
    fs.watchFile(file, { interval: 500 }, (cur) => {
      if (cur.size < offset) offset = 0;
      if (cur.size === offset) return;
      const fd = fs.openSync(file, 'r');
      const buf = Buffer.alloc(cur.size - offset);
      fs.readSync(fd, buf, 0, buf.length, offset);
      fs.closeSync(fd);
      const text = buf.toString('utf8');
      // Hold back a trailing partial line until it is complete.
      const complete = text.slice(0, text.lastIndexOf('\n') + 1);
      offset += Buffer.byteLength(complete);
      for (const entry of parseLines(complete)) console.log(formatOutput([entry], format));
    });
  });
}

function sizeOf(file: string): number {
  try { return fs.statSync(file).size; } catch { return 0; }
}
