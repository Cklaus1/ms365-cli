import logger from './logger.js';

/** Process exit codes — a stable contract for scripts and AI agents. */
export enum ExitCode {
  OK = 0,
  GENERIC = 1,
  USAGE = 2,
  AUTH = 3,
  PERMISSION = 4,
  NOT_FOUND = 5,
  THROTTLED = 6,
  CONFIRMATION_REFUSED = 7,
}

export interface CliErrorDetails {
  code?: string;
  exitCode?: ExitCode;
  status?: number;
  graphCode?: string;
  requestId?: string;
  fix?: string;
}

/** An error carrying a machine-readable code and the exit code it maps to. */
export class CliError extends Error {
  code: string;
  exitCode: ExitCode;
  status?: number;
  graphCode?: string;
  requestId?: string;
  fix?: string;

  constructor(message: string, details: CliErrorDetails = {}) {
    super(message);
    this.name = 'CliError';
    this.code = details.code ?? 'error';
    this.exitCode = details.exitCode ?? ExitCode.GENERIC;
    this.status = details.status;
    this.graphCode = details.graphCode;
    this.requestId = details.requestId;
    this.fix = details.fix;
  }
}

/** Thrown when a destructive operation's confirmation is declined or impossible (non-TTY). */
export class ConfirmationRefusedError extends CliError {
  constructor(message = 'Operation cancelled.') {
    super(message, { code: 'confirmation_refused', exitCode: ExitCode.CONFIRMATION_REFUSED });
    this.name = 'ConfirmationRefusedError';
  }
}

const AUTH_PATTERN = /not logged in|no client id|invalid_grant|AADSTS|interaction_required|consent_required|authentication failed/i;

/** Normalize any thrown value into a CliError. */
export function toCliError(err: unknown): CliError {
  if (err instanceof CliError) return err;
  const message = err instanceof Error ? err.message : String(err);
  if (AUTH_PATTERN.test(message)) {
    return new CliError(message, { code: 'auth_required', exitCode: ExitCode.AUTH, fix: 'ms365 auth login' });
  }
  return new CliError(message);
}

export type ErrorFormat = 'text' | 'json';
let errorFormat: ErrorFormat = process.env.MS365_CLI_ERROR_FORMAT === 'json' ? 'json' : 'text';
export function setErrorFormat(fmt: ErrorFormat): void { errorFormat = fmt; }

/** Print an error to stderr in the selected format and return its exit code. */
export function reportError(err: unknown): ExitCode {
  const e = toCliError(err);
  if (errorFormat === 'json') {
    const payload: Record<string, unknown> = { code: e.code, exitCode: e.exitCode, message: e.message };
    if (e.status !== undefined) payload.status = e.status;
    if (e.graphCode) payload.graphCode = e.graphCode;
    if (e.requestId) payload.requestId = e.requestId;
    if (e.fix) payload.fix = e.fix;
    process.stderr.write(JSON.stringify({ error: payload }) + '\n');
  } else {
    const hint = e.fix && !e.message.includes(e.fix) ? ` (fix: ${e.fix})` : '';
    const rid = e.requestId ? ` [request-id: ${e.requestId}]` : '';
    process.stderr.write(`Error: ${e.message}${hint}${rid}\n`);
  }
  return e.exitCode;
}

/** Wait for stdout and stderr to flush, then exit. MSAL/keytar handles keep the loop alive otherwise. */
async function drainAndExit(code: number): Promise<void> {
  await new Promise<void>((resolve) => process.stdout.write('', () => resolve()));
  await new Promise<void>((resolve) => process.stderr.write('', () => resolve()));
  process.exit(code);
}

let exitHandler: (code: number) => void | Promise<void> = drainAndExit;
/** Override process termination (tests run commands in-process). */
export function setExitHandler(fn: (code: number) => void | Promise<void>): void { exitHandler = fn; }
export function resetExitHandler(): void { exitHandler = drainAndExit; }

/**
 * Wraps an async command action with consistent error handling and a
 * deterministic exit code (see ExitCode).
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function handleErrors(fn: (...args: any[]) => Promise<void>): (...args: any[]) => Promise<void> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return async (...args: any[]) => {
    let code: number = ExitCode.OK;
    try {
      await fn(...args);
      const preset = Number(process.exitCode ?? 0);
      code = Number.isFinite(preset) ? preset : ExitCode.GENERIC;
    } catch (err) {
      code = reportError(err);
    } finally {
      logger.close();
    }
    await exitHandler(code);
  };
}
