import { Command } from 'commander';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import { dryRunTrace, isDryRun, outputOption, success } from '../utils.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { listEntries, releaseKey } from '../idempotency.js';

/** `ms365 idempotency` — inspect and clear the local --idempotency-key ledger (R-6, F17). Offline. */
export function registerIdempotencyCommands(program: Command): void {
  const idem = program.command('idempotency').description('Inspect and clear the local --idempotency-key ledger');

  idem
    .command('list')
    .description('Print the idempotency keys recorded in the last 24 h (pending, done or unknown)')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      console.log(formatOutput(listEntries(), opts.output as OutputFormat));
    }));

  idem
    .command('release <key>')
    .description('Forget an idempotency key so the next send with it runs again (check first that it did not take effect)')
    .action(handleErrors(async (key: string) => {
      if (isDryRun()) return dryRunTrace(`would release idempotency key ${key}`);
      if (!releaseKey(key)) {
        throw new CliError(`No idempotency key "${key}" in the ledger.`, {
          code: 'not_found', exitCode: ExitCode.NOT_FOUND, fix: 'ms365 idempotency list',
        });
      }
      success(`Released idempotency key "${key}".`);
    }));
}
