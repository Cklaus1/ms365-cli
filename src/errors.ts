import logger from './logger.js';

/** A destructive action was declined (or could not be confirmed). Exits 1: nothing changed. */
export class CancelledError extends Error {
  constructor(message = 'Cancelled; nothing changed.') {
    super(message);
    this.name = 'CancelledError';
  }
}

/** Print an error for the user and log it. */
export function reportError(err: unknown): void {
  const message = err instanceof Error ? err.message : String(err);
  logger.error(message);
  console.error(`Error: ${message}`);
}

/**
 * Wraps a command action: errors become a one-line message and exit code 1.
 * The process itself exits once, in `exitProcess`, after Commander's parseAsync settles.
 */
export function handleErrors<A extends any[]>(fn: (...args: A) => Promise<void>): (...args: A) => Promise<void> {
  return async (...args: A) => {
    try {
      await fn(...args);
    } catch (err) {
      reportError(err);
      process.exitCode = 1;
    }
  };
}

function drain(stream: NodeJS.WriteStream): Promise<void> {
  return new Promise((resolve) => {
    if (stream.destroyed || !stream.writable) return resolve();
    stream.write('', () => resolve());
  });
}

/**
 * The single process exit. Waits for stdout/stderr to flush and the log file to close,
 * then exits — MSAL keeps handles open, so the event loop would not end on its own.
 */
export async function exitProcess(code: number): Promise<never> {
  await Promise.all([drain(process.stdout), drain(process.stderr), logger.close()]);
  process.exit(code);
}
