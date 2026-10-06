import { execFile } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

/** Server-side environment variables a child may see. Never MS365_CLI_TOKEN: children use the host's token cache. */
const FORWARDED_ENV = [
  'PATH', 'HOME', 'TZ',
  'MS365_CLI_LOG_LEVEL', 'MS365_CLI_LOG_FILE',
  'MS365_CLI_ORG_MODE', 'MS365_MCP_ORG_MODE', 'MS365_CLI_READ_ONLY',
];

/**
 * The explicit child environment, built once at server start. The client and tenant ID are resolved here from the
 * server's environment (which includes the `.env` of the directory `ms365 serve` was started in): the child runs in
 * an empty temp directory, so this env is its only configuration source.
 */
export function buildChildEnv(env: NodeJS.ProcessEnv): Record<string, string> {
  const child: Record<string, string> = {};
  for (const name of FORWARDED_ENV) {
    const value = env[name];
    if (value !== undefined) child[name] = value;
  }
  const clientId = env.MS365_CLI_CLIENT_ID || env.MS365_MCP_CLIENT_ID;
  const tenantId = env.MS365_CLI_TENANT_ID || env.MS365_MCP_TENANT_ID;
  if (clientId) child.MS365_CLI_CLIENT_ID = clientId;
  if (tenantId) child.MS365_CLI_TENANT_ID = tenantId;
  return child;
}

export interface ExecutorOptions {
  /** Node flags of the running server (`process.execArgv`), so children load the same way. */
  nodeArgs: string[];
  /** Entry script of the running CLI. */
  cliEntry: string;
  /** Global flags fixed at server start (`--org-mode`, `--read-only`, `--api-version=…`), placed before every command. */
  startFlags: string[];
  env: Record<string, string>;
  timeoutMs: number;
  maxStdoutBytes: number;
  /** Limit on the total size of the files a child writes into its working directory; it is killed past it. */
  maxFileBytes: number;
  concurrency: number;
}

export type ExecResult =
  | { status: 'ok'; stdout: Buffer; stderr: string }
  | { status: 'failed'; message: string }
  | { status: 'timeout'; message: string };

/** Runs CLI children with a concurrency limit; excess requests wait in FIFO order. */
export class Executor {
  private active = 0;
  private readonly queue: Array<() => void> = [];
  private readonly idleWaiters: Array<() => void> = [];

  constructor(private readonly options: ExecutorOptions) {}

  /** Run `ms365 <startFlags> <argv>` in `cwd` once a slot is free. */
  async run(argv: string[], cwd: string): Promise<ExecResult> {
    if (this.active >= this.options.concurrency) {
      const { promise, resolve } = Promise.withResolvers<void>();
      this.queue.push(resolve);
      await promise;
    } else {
      this.active++;
    }
    try {
      return await this.spawn(argv, cwd);
    } finally {
      const next = this.queue.shift();
      if (next) {
        next();
      } else {
        this.active--;
        if (this.active === 0) for (const waiter of this.idleWaiters.splice(0)) waiter();
      }
    }
  }

  /** Resolves once no child is running and none is queued. */
  idle(): Promise<void> {
    if (this.active === 0) return Promise.resolve();
    const { promise, resolve } = Promise.withResolvers<void>();
    this.idleWaiters.push(resolve);
    return promise;
  }

  private spawn(argv: string[], cwd: string): Promise<ExecResult> {
    const o = this.options;
    const { promise, resolve } = Promise.withResolvers<ExecResult>();
    let timedOut = false;
    let tooLarge = false;
    const child = execFile(
      process.execPath,
      [...o.nodeArgs, o.cliEntry, ...o.startFlags, ...argv],
      { cwd, env: o.env, encoding: 'buffer', maxBuffer: o.maxStdoutBytes, killSignal: 'SIGKILL', windowsHide: true },
      (err, stdout, stderr) => {
        clearTimeout(timer);
        watcher.close();
        const stderrText = stderr.toString('utf8');
        if (timedOut) {
          resolve({ status: 'timeout', message: `Command timed out after ${o.timeoutMs / 1000} s` });
        } else if (tooLarge) {
          resolve({ status: 'failed', message: `Output file exceeds the ${o.maxFileBytes / 1024 / 1024} MB remote limit` });
        } else if (err && 'code' in err && err.code === 'ERR_CHILD_PROCESS_STDIO_MAXBUFFER') {
          resolve({ status: 'failed', message: `Command output exceeds the ${o.maxStdoutBytes / 1024 / 1024} MB limit` });
        } else if (err) {
          resolve({ status: 'failed', message: stderrText.trim() || err.message });
        } else {
          resolve({ status: 'ok', stdout, stderr: stderrText });
        }
      },
    );
    // Every write into cwd (download temp files included) is checked as it happens, so a large download
    // stops at the limit instead of filling the disk until the timeout.
    const watcher = fs.watch(cwd, () => {
      if (tooLarge || dirSize(cwd) <= o.maxFileBytes) return;
      tooLarge = true;
      child.kill('SIGKILL');
    });
    watcher.on('error', () => watcher.close());
    // No stdin: destructive commands refuse without a TTY (non-goal: no --yes).
    child.stdin?.end();
    const timer = setTimeout(() => {
      timedOut = true;
      child.kill('SIGKILL');
    }, o.timeoutMs);
    return promise;
  }
}

/** Total size of the regular files directly in `dir` (entries that vanish meanwhile count as 0). */
function dirSize(dir: string): number {
  let total = 0;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.isFile()) total += fs.statSync(path.join(dir, entry.name), { throwIfNoEntry: false })?.size ?? 0;
  }
  return total;
}
