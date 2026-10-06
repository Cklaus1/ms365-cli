import path from 'node:path';
import { Command, InvalidArgumentError } from 'commander';
import { configDir } from '../auth.js';
import { buildCommandTree } from '../command-tree.js';
import { handleErrors } from '../errors.js';
import { buildChildEnv } from '../serve/executor.js';
import { checkTransport, serve } from '../serve/server.js';
import { introspectionProgram } from './registry.js';

interface ServeOpts {
  socket?: string;
  port?: number;
  host?: string;
  appCap: string;
  auditLog?: string;
  timeout: number;
}

function port(value: string): number {
  const n = Number(value);
  if (!/^\d+$/.test(value) || n > 65535) throw new InvalidArgumentError('Expected a TCP port (0-65535).');
  return n;
}

function seconds(value: string): number {
  const n = Number(value);
  if (!/^\d+$/.test(value) || n < 1) throw new InvalidArgumentError('Expected a positive number of seconds.');
  return n;
}

export function registerServeCommand(program: Command): void {
  program
    .command('serve')
    .description('Serve this host\'s ms365 login to tailnet agents (POST /v1/execute; put behind tailscale serve)')
    .requiredOption('--app-cap <name>', 'Capability key from the tailnet policy (the --accept-app-caps value, e.g. example.com/cap/ms365)')
    .option('--socket <path>', 'Unix socket to listen on, mode 0600 (default: ~/.config/ms365-cli/serve.sock)')
    .option('--port <n>', 'Listen on loopback TCP instead (fallback; every local process is trusted)', port)
    .option('--host <address>', 'Loopback address for --port (127.0.0.1, localhost, ::1)')
    .option('--audit-log <path>', 'Audit log, one JSON line per request (default: ~/.config/ms365-cli/serve-audit.jsonl)')
    .option('--timeout <seconds>', 'Per-request command timeout', seconds, 120)
    .action(handleErrors(async (opts: ServeOpts, cmd: Command) => {
      checkTransport(opts);
      // Server start flags reach every child; requests can only add --dry-run.
      const g = cmd.optsWithGlobals<{ orgMode?: boolean; readOnly?: boolean; apiVersion?: string }>();
      const startFlags = [
        ...(g.orgMode ? ['--org-mode'] : []),
        ...(g.readOnly ? ['--read-only'] : []),
        ...(g.apiVersion ? [`--api-version=${g.apiVersion}`] : []),
      ];
      await serve({
        socket: opts.port === undefined ? path.resolve(opts.socket ?? path.join(configDir(), 'serve.sock')) : undefined,
        port: opts.port,
        host: opts.host,
        appCap: opts.appCap,
        auditLog: path.resolve(opts.auditLog ?? path.join(configDir(), 'serve-audit.jsonl')),
        tree: buildCommandTree(introspectionProgram()),
        executor: {
          nodeArgs: process.execArgv,
          cliEntry: path.resolve(process.argv[1]),
          startFlags,
          env: buildChildEnv(process.env),
          timeoutMs: opts.timeout * 1000,
          maxStdoutBytes: 50 * 1024 * 1024,
          concurrency: 4,
        },
      });
    }));
}
