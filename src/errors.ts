import logger from './logger.js';

/**
 * Wraps an async command action with consistent error handling.
 * Catches errors and prints user-friendly messages instead of stack traces.
 * Forces process exit after completion — MSAL/keytar hold the event loop open.
 *
 * Both tsx and Commander.js override process.exit() with versions that only set
 * exitCode but don't terminate. We use process.kill(SIGTERM) as a reliable
 * termination mechanism that cannot be intercepted by JS-level overrides.
 */
export function handleErrors<T extends (...args: never[]) => Promise<void>>(fn: T): T {
  return (async (...args: Parameters<T>) => {
    try {
      await fn(...args);
      process.exitCode = 0;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      console.error(`Error: ${message}`);
      process.exitCode = 1;
    } finally {
      logger.close();
      // MSAL and keytar native workers keep the event loop open.
      // tsx/Commander override process.exit() — use SIGTERM for clean exit.
      // Small delay lets stdout/stderr flush before termination.
      setTimeout(() => process.kill(process.pid, 'SIGTERM'), 200);
    }
  }) as T;
}
