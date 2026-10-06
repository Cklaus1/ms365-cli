import fs from 'node:fs';
import fsp from 'node:fs/promises';
import http from 'node:http';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import { randomBytes } from 'node:crypto';
import { buildArgv, findCommand, InvalidInvocationError, type CommandNode, type CommandTree } from '../command-tree.js';
import { detectMimeType } from '../mime.js';
import { info } from '../utils.js';
import { APP_CAPS_HEADER, CapsHeaderError, parseAppCapabilities, type Caller } from './app-caps.js';
import { AuditLog, type AuditEntry } from './audit.js';
import { CAPABILITIES, remotePolicy, type RemoteCommand } from './capabilities.js';
import { Executor, type ExecutorOptions } from './executor.js';

/** Total size of uploaded files per request, and of returned output files. */
export const MAX_FILE_BYTES = 25 * 1024 * 1024;
/** Request body limit: the file cap plus room for the JSON part and multipart framing. */
const MAX_BODY_BYTES = MAX_FILE_BYTES + 1024 * 1024;

const LOOPBACK_HOSTS = ['127.0.0.1', 'localhost', '::1'];

export type ErrorCode = 'capability_missing' | 'invalid_args' | 'command_failed' | 'timeout' | 'not_found';

/** Error response: `{ ok: false, request_id, error: { code, message } }` with an HTTP status. */
class ApiError extends Error {
  constructor(readonly status: number, readonly code: ErrorCode, message: string) {
    super(message);
  }
}

export interface ServeOptions {
  /** Unix socket path (default transport). */
  socket?: string;
  /** Loopback TCP port (fallback transport; 0 picks a free port). */
  port?: number;
  host?: string;
  /** Capability key in `Tailscale-App-Capabilities` (the `--accept-app-caps` value). */
  appCap: string;
  auditLog: string;
  tree: CommandTree;
  executor: ExecutorOptions;
}

interface ExecuteRequest {
  command: string;
  positionals: unknown[];
  args: Record<string, unknown>;
}

interface OutputFile {
  name: string;
  contentType: string;
  size: number;
  base64: string;
}

function json(res: http.ServerResponse, status: number, body: unknown): void {
  const text = JSON.stringify(body);
  res.writeHead(status, { 'content-type': 'application/json', 'content-length': Buffer.byteLength(text) });
  res.end(text);
}

/** Read the whole body, rejecting with 413 past `MAX_BODY_BYTES`. */
async function readBody(req: http.IncomingMessage): Promise<Buffer<ArrayBuffer>> {
  const declared = Number(req.headers['content-length']);
  if (declared > MAX_BODY_BYTES) throw tooLarge();
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req as AsyncIterable<Buffer>) {
    size += chunk.byteLength;
    if (size <= MAX_BODY_BYTES) chunks.push(chunk);
  }
  if (size > MAX_BODY_BYTES) throw tooLarge();
  return Buffer.concat(chunks);
}

const tooLarge = () => new ApiError(413, 'invalid_args', `Request exceeds the ${MAX_FILE_BYTES / 1024 / 1024} MB file limit`);

function parseExecuteRequest(value: unknown): ExecuteRequest {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new ApiError(400, 'invalid_args', 'Request must be a JSON object');
  }
  const command = 'command' in value ? value.command : undefined;
  const positionals = 'positionals' in value ? value.positionals : [];
  const args = 'args' in value ? value.args : {};
  if (typeof command !== 'string') throw new ApiError(400, 'invalid_args', '"command" must be a string');
  if (!Array.isArray(positionals)) throw new ApiError(400, 'invalid_args', '"positionals" must be an array');
  if (typeof args !== 'object' || args === null || Array.isArray(args)) {
    throw new ApiError(400, 'invalid_args', '"args" must be an object');
  }
  return { command, positionals, args: Object.fromEntries(Object.entries(args)) };
}

/** JSON body, or multipart/form-data with a `request` JSON part plus file parts referenced as `{"$file": "<part>"}`. */
async function parseBody(req: http.IncomingMessage): Promise<{ request: ExecuteRequest; files: Map<string, File> }> {
  const body = await readBody(req);
  const contentType = req.headers['content-type'] ?? 'application/json';
  const files = new Map<string, File>();
  let requestText: string;
  if (contentType.startsWith('multipart/form-data')) {
    let form: FormData;
    try {
      form = await new Request('http://localhost/', { method: 'POST', headers: { 'content-type': contentType }, body }).formData();
    } catch {
      throw new ApiError(400, 'invalid_args', 'Malformed multipart body');
    }
    const part = form.get('request');
    if (typeof part !== 'string') throw new ApiError(400, 'invalid_args', 'Missing "request" part');
    requestText = part;
    let total = 0;
    for (const [name, value] of form) {
      if (typeof value === 'string') continue;
      files.set(name, value);
      total += value.size;
    }
    if (total > MAX_FILE_BYTES) throw tooLarge();
  } else if (contentType.startsWith('application/json')) {
    requestText = body.toString('utf8');
  } else {
    throw new ApiError(400, 'invalid_args', `Unsupported content type: ${contentType}`);
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(requestText);
  } catch {
    throw new ApiError(400, 'invalid_args', 'Request is not valid JSON');
  }
  return { request: parseExecuteRequest(parsed), files };
}

/** Last path segment only, so an uploaded or requested name always stays inside the request's temp dir. */
function safeName(name: string, fallback: string): string {
  const base = path.basename(name.replace(/\\/g, '/'));
  return base === '' || base === '.' || base === '..' ? fallback : base;
}

/** Child stdout → `data`: empty → null; JSON → parsed; anything else → the raw string. */
function parseStdout(stdout: Buffer): unknown {
  const text = stdout.toString('utf8');
  if (text.trim() === '') return null;
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

async function collectOutputs(dir: string): Promise<OutputFile[]> {
  const files: OutputFile[] = [];
  let total = 0;
  for (const entry of await fsp.readdir(dir, { withFileTypes: true })) {
    if (!entry.isFile()) continue;
    const bytes = await fsp.readFile(path.join(dir, entry.name));
    total += bytes.byteLength;
    if (total > MAX_FILE_BYTES) {
      throw new ApiError(502, 'command_failed', `Output file exceeds the ${MAX_FILE_BYTES / 1024 / 1024} MB remote limit`);
    }
    files.push({ name: entry.name, contentType: detectMimeType(entry.name), size: bytes.byteLength, base64: bytes.toString('base64') });
  }
  return files;
}

class RemoteServer {
  private readonly executor: Executor;
  private readonly audit: AuditLog;

  constructor(private readonly options: ServeOptions) {
    this.executor = new Executor(options.executor);
    this.audit = new AuditLog(options.auditLog);
  }

  async handle(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const started = performance.now();
    const requestId = `req_${randomBytes(8).toString('hex')}`;
    // `URL.parse` instead of `new URL`: llhttp accepts targets the WHATWG parser rejects (e.g. `//[`); they get a 404.
    const pathname = URL.parse(req.url ?? '/', 'http://localhost')?.pathname ?? '';
    const forwarded = req.headers['x-forwarded-for'];
    const entry: AuditEntry = {
      ts: new Date().toISOString(),
      request_id: requestId,
      agent: null,
      src: (Array.isArray(forwarded) ? forwarded.join(', ') : forwarded) ?? null,
      method: req.method ?? '',
      path: pathname,
      command: null,
      args: [],
      status: 500,
      duration_ms: 0,
    };
    try {
      const body = await this.route(req, pathname, entry);
      entry.status = 200;
      json(res, 200, { ok: true, request_id: requestId, ...body });
    } catch (err) {
      const apiError = err instanceof ApiError ? err : new ApiError(500, 'command_failed', err instanceof Error ? err.message : String(err));
      entry.status = apiError.status;
      entry.error = apiError.code;
      json(res, apiError.status, { ok: false, request_id: requestId, error: { code: apiError.code, message: apiError.message } });
    } finally {
      entry.duration_ms = Math.round(performance.now() - started);
      this.audit.write(entry);
    }
  }

  private caller(req: http.IncomingMessage, entry: AuditEntry): Caller {
    try {
      const caller = parseAppCapabilities(req.headers[APP_CAPS_HEADER], this.options.appCap);
      entry.agent = caller.agent;
      return caller;
    } catch (err) {
      if (err instanceof CapsHeaderError) throw new ApiError(403, 'capability_missing', err.message);
      throw err;
    }
  }

  private async route(req: http.IncomingMessage, pathname: string, entry: AuditEntry): Promise<Record<string, unknown>> {
    if (req.method === 'GET' && pathname === '/v1/health') return {};
    if (req.method === 'GET' && pathname === '/v1/capabilities') {
      const caller = this.caller(req, entry);
      const commands = this.options.tree.leaves
        .filter(leaf => remotePolicy(leaf.path)?.caps.every(c => caller.allow.has(c)))
        .map(leaf => leaf.path);
      return {
        agent: caller.agent,
        capabilities: CAPABILITIES.filter(c => caller.allow.has(c)),
        commands,
      };
    }
    if (req.method === 'POST' && pathname === '/v1/execute') {
      const caller = this.caller(req, entry);
      const { request, files } = await parseBody(req);
      return this.execute(caller, request, files, entry);
    }
    throw new ApiError(404, 'not_found', `No route for ${req.method} ${pathname}`);
  }

  private async execute(caller: Caller, request: ExecuteRequest, files: Map<string, File>, entry: AuditEntry): Promise<Record<string, unknown>> {
    const node = findCommand(this.options.tree, request.command);
    entry.command = node?.path ?? request.command;
    entry.args = Object.keys(request.args);
    if (!node) throw new ApiError(400, 'invalid_args', `Unknown command: ${request.command}`);

    const policy = remotePolicy(node.path);
    if (!policy) throw new ApiError(403, 'capability_missing', `${node.path} cannot run remotely`);
    const missing = policy.caps.filter(c => !caller.allow.has(c));
    if (missing.length > 0) throw new ApiError(403, 'capability_missing', `Missing capability: ${missing.join(', ')}`);

    const { dryRun, ...args } = request.args;
    if (dryRun !== undefined && typeof dryRun !== 'boolean') throw new ApiError(400, 'invalid_args', 'dryRun must be a boolean');

    const tmp = await fsp.mkdtemp(path.join(os.tmpdir(), 'ms365-serve-'));
    try {
      const outDir = path.join(tmp, 'out');
      await fsp.mkdir(outDir, { mode: 0o700 });
      const positionals = await this.substituteInputs(policy, node, request.positionals, files, tmp);
      if (policy.output) {
        // Output paths are the server's: the caller only names the file; bytes come back in `files`.
        const requested = typeof args.out === 'string' ? safeName(args.out, 'download') : undefined;
        delete args.out;
        delete args.force;
        if (requested) args.out = path.join(outDir, requested);
        else if (policy.output === 'stdout') args.out = path.join(outDir, 'download');
      }
      if (node.options.some(o => o.long === '--output')) args.output = 'json';

      let argv: string[];
      try {
        argv = buildArgv(node, { positionals, args });
      } catch (err) {
        if (err instanceof InvalidInvocationError) throw new ApiError(400, 'invalid_args', err.message);
        throw err;
      }
      const result = await this.executor.run([...(dryRun ? ['--dry-run'] : []), ...argv], outDir);
      if (result.status === 'timeout') throw new ApiError(504, 'timeout', result.message);
      if (result.status === 'failed') throw new ApiError(502, 'command_failed', result.message);

      const body: Record<string, unknown> = { data: parseStdout(result.stdout) };
      if (policy.output) body.files = await collectOutputs(outDir);
      if (result.stderr.trim() !== '') body.stderr = result.stderr;
      return body;
    } finally {
      await fsp.rm(tmp, { recursive: true, force: true });
    }
  }

  /** Write the uploaded file for the command's input positional and put its path in its place. */
  private async substituteInputs(
    policy: RemoteCommand,
    node: CommandNode,
    positionals: unknown[],
    files: Map<string, File>,
    tmp: string,
  ): Promise<unknown[]> {
    const result = [...positionals];
    if (policy.input === undefined || policy.input >= result.length) return result;
    const ref = result[policy.input];
    const argName = node.arguments[policy.input]?.name ?? 'file';
    const part = typeof ref === 'object' && ref !== null && '$file' in ref ? ref.$file : undefined;
    if (typeof part !== 'string') {
      throw new ApiError(400, 'invalid_args', `positional <${argName}> must reference an uploaded file: {"$file": "<part>"}`);
    }
    const file = files.get(part);
    if (!file) throw new ApiError(400, 'invalid_args', `No file part named "${part}"`);
    const dir = path.join(tmp, 'in');
    await fsp.mkdir(dir, { mode: 0o700 });
    const target = path.join(dir, safeName(file.name, 'upload'));
    await fsp.writeFile(target, file.stream(), { mode: 0o600 });
    result[policy.input] = target;
    return result;
  }

  idle(): Promise<void> {
    return this.executor.idle();
  }

  close(): void {
    this.audit.close();
  }
}

/** Remove a leftover socket file, but only when nothing is listening on it. */
async function clearStaleSocket(socket: string): Promise<void> {
  const stat = fs.lstatSync(socket, { throwIfNoEntry: false });
  if (!stat) return;
  if (!stat.isSocket()) throw new Error(`Refusing to replace ${socket}: not a socket`);
  const { promise, resolve } = Promise.withResolvers<boolean>();
  const probe = net.connect(socket);
  probe.once('connect', () => {
    probe.destroy();
    resolve(true);
  });
  probe.once('error', () => resolve(false));
  if (await promise) throw new Error(`Another server is listening on ${socket}`);
  fs.rmSync(socket, { force: true });
}

/** Check transport options before anything binds: loopback TCP only, never both transports. */
export function checkTransport(options: { socket?: string; port?: number; host?: string }): void {
  if (options.host !== undefined && !LOOPBACK_HOSTS.includes(options.host)) {
    throw new Error(`Refusing to listen on ${options.host}: ms365 serve binds loopback only (${LOOPBACK_HOSTS.join(', ')}). Expose it with tailscale serve.`);
  }
  if (options.host !== undefined && options.port === undefined) throw new Error('--host requires --port');
  if (options.port !== undefined && options.socket !== undefined) throw new Error('Use either --socket or --port, not both');
}

/**
 * Run the server until SIGTERM/SIGINT, then shut down gracefully: stop accepting, let running and queued
 * requests finish, remove the socket file. Resolves once everything is closed.
 */
export async function serve(options: ServeOptions): Promise<void> {
  const app = new RemoteServer(options);
  const inflight = new Set<Promise<void>>();
  let stopping = false;
  const server = http.createServer((req, res) => {
    const handled = app.handle(req, res)
      .catch((err: unknown) => {
        // handle() answers every error itself; this only guards against a bug there taking the server down.
        console.error(`Warning: request failed: ${err instanceof Error ? err.message : String(err)}`);
        if (!res.headersSent) json(res, 500, { ok: false, error: { code: 'command_failed', message: 'Internal server error' } });
        else res.destroy();
      })
      .finally(() => {
        inflight.delete(handled);
        if (stopping) server.closeIdleConnections();
      });
    inflight.add(handled);
  });

  const socket = options.port === undefined ? options.socket : undefined;
  try {
    if (socket) {
      fs.mkdirSync(path.dirname(socket), { recursive: true, mode: 0o700 });
      await clearStaleSocket(socket);
    }
    const listening = Promise.withResolvers<void>();
    server.once('error', listening.reject);
    server.once('listening', listening.resolve);
    if (socket) {
      // Created 0600 from the start: no window in which another local user could connect.
      const umask = process.umask(0o177);
      try {
        server.listen(socket);
      } finally {
        process.umask(umask);
      }
    } else {
      server.listen(options.port, options.host ?? '127.0.0.1');
    }
    await listening.promise;
  } catch (err) {
    app.close();
    throw err;
  }

  if (socket) {
    fs.chmodSync(socket, 0o600);
    info(`ms365 serve listening on unix:${socket}`);
  } else {
    const address = server.address() as net.AddressInfo;
    const hostPart = address.family === 'IPv6' ? `[${address.address}]` : address.address;
    console.error(
      `Warning: bound to TCP ${hostPart}:${address.port}. Every local process can connect and send its own ` +
      `${APP_CAPS_HEADER} header; prefer --socket.`,
    );
    info(`ms365 serve listening on http://${hostPart}:${address.port}`);
  }

  const signal = Promise.withResolvers<NodeJS.Signals>();
  const onSignal = (sig: NodeJS.Signals) => signal.resolve(sig);
  process.once('SIGTERM', onSignal);
  process.once('SIGINT', onSignal);
  const received = await signal.promise;
  process.off('SIGTERM', onSignal);
  process.off('SIGINT', onSignal);

  info(`Received ${received}; finishing ${inflight.size} request(s)`);
  stopping = true;
  const closed = Promise.withResolvers<void>();
  server.close(() => closed.resolve());
  server.closeIdleConnections();
  await Promise.allSettled(inflight);
  await app.idle();
  await closed.promise;
  if (socket) fs.rmSync(socket, { force: true });
  app.close();
  info('ms365 serve stopped');
}
