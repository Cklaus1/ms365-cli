import fs from 'node:fs';
import fsp from 'node:fs/promises';
import http from 'node:http';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import { randomBytes } from 'node:crypto';
import { buildArgv, findCommand, InvalidInvocationError, type CommandNode, type CommandTree } from '../command-tree.js';
import { detectMimeType } from '../mime.js';
import { APP_CAPS_HEADER, CapsHeaderError, parseAppCapabilities, type Caller } from './app-caps.js';
import { AuditLog, type AuditEntry } from './audit.js';
import { CAPABILITIES, remotePolicy, type Capability, type RemoteCommand } from './capabilities.js';
import { Executor, type ExecutorOptions } from './executor.js';

/** Lifetime of an unused upload token or file URL. */
const FILE_TTL_MS = 10 * 60 * 1000;
/** `/v1/execute` body limit: files travel separately, and a child's argv could not hold more anyway. */
const MAX_REQUEST_BYTES = 1024 * 1024;
/** `GET /v1/files/{token}`: 32 random bytes, base64url. */
const FILE_PATH = /^\/v1\/files\/([A-Za-z0-9_-]{43})$/;

const LOOPBACK_HOSTS = ['127.0.0.1', 'localhost', '::1'];

/** Bounds for caller-chosen text in audit lines, so a caller cannot grow the log by what it sends. */
const AUDIT_TEXT_MAX = 128;
const AUDIT_ARGS_MAX = 32;
const AUDIT_ARG_KEY_MAX = 64;

/** `text` cut to at most `max` characters, ending in "…" when cut. */
function clip(text: string, max: number): string {
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

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
  /** Total size of uploaded files per request, and of returned output files (`--max-file-mb`). */
  maxFileBytes: number;
  executor: ExecutorOptions;
}

interface ExecuteRequest {
  command: string;
  positionals: unknown[];
  args: Record<string, unknown>;
}

/** Entry of `files` in an execute reply: the bytes are fetched once from `url`. */
interface OutputFile {
  name: string;
  contentType: string;
  size: number;
  url: string;
}

/** A file on disk in a temp dir of its own, which is removed with it. */
interface StoredFile {
  dir: string;
  path: string;
  name: string;
  size: number;
}

/** An output file waiting for its one-time `GET /v1/files/{token}`. */
interface Download extends StoredFile {
  contentType: string;
  /** Capabilities of the producing command, required again to fetch it. */
  caps: readonly Capability[];
  command: string;
}

async function removeDir(dir: string): Promise<void> {
  await fsp.rm(dir, { recursive: true, force: true });
}

/**
 * Files behind single-use tokens (256-bit random). A file leaves the store when taken, when its TTL passes
 * (its dir is then removed) or on `clear()`.
 */
class FileTokens<T extends StoredFile> {
  private readonly entries = new Map<string, { file: T; timer: NodeJS.Timeout }>();

  add(file: T): string {
    const token = randomBytes(32).toString('base64url');
    const timer = setTimeout(() => {
      const expired = this.take(token);
      if (expired) void removeDir(expired.dir);
    }, FILE_TTL_MS);
    timer.unref();
    this.entries.set(token, { file, timer });
    return token;
  }

  get(token: string): T | undefined {
    return this.entries.get(token)?.file;
  }

  /** Remove `token` from the store; the caller now owns (and removes) the file. */
  take(token: string): T | undefined {
    const entry = this.entries.get(token);
    if (!entry) return undefined;
    this.entries.delete(token);
    clearTimeout(entry.timer);
    return entry.file;
  }

  /** Remove every stored file. */
  async clear(): Promise<void> {
    const files = [...this.entries.keys()].map(token => this.take(token)!);
    await Promise.all(files.map(file => removeDir(file.dir)));
  }
}

function json(res: http.ServerResponse, status: number, body: unknown): void {
  const text = JSON.stringify(body);
  res.writeHead(status, { 'content-type': 'application/json', 'content-length': Buffer.byteLength(text) });
  res.end(text);
}

/**
 * Stream an open file into `res` through one reused buffer, each chunk flushed before the next read, so memory
 * stays flat however large the file is (piping a read stream leaves every chunk to the garbage collector).
 */
async function sendChunks(fh: fsp.FileHandle, res: http.ServerResponse): Promise<void> {
  const buffer = Buffer.allocUnsafe(256 * 1024);
  for (;;) {
    const { bytesRead } = await fh.read(buffer, 0, buffer.length, null);
    if (bytesRead === 0) break;
    const { promise, resolve, reject } = Promise.withResolvers<void>();
    res.write(buffer.subarray(0, bytesRead), err => (err ? reject(err) : resolve()));
    await promise;
  }
  const { promise, resolve } = Promise.withResolvers<void>();
  res.end(resolve);
  await promise;
}

/** Read the whole JSON body, rejecting with 413 past `MAX_REQUEST_BYTES`. */
async function readBody(req: http.IncomingMessage): Promise<Buffer<ArrayBuffer>> {
  const tooLarge = () => new ApiError(413, 'invalid_args', `Request exceeds the ${megabytes(MAX_REQUEST_BYTES)} limit; send files with PUT /v1/uploads`);
  if (Number(req.headers['content-length']) > MAX_REQUEST_BYTES) throw tooLarge();
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req as AsyncIterable<Buffer>) {
    size += chunk.byteLength;
    if (size <= MAX_REQUEST_BYTES) chunks.push(chunk);
  }
  if (size > MAX_REQUEST_BYTES) throw tooLarge();
  return Buffer.concat(chunks);
}

const megabytes = (bytes: number) => `${bytes / 1024 / 1024} MB`;

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

/** JSON body; input files are referenced as `{"$upload": "<token>"}` from `PUT /v1/uploads`. */
async function parseBody(req: http.IncomingMessage): Promise<ExecuteRequest> {
  const contentType = req.headers['content-type'] ?? 'application/json';
  if (!contentType.startsWith('application/json')) {
    throw new ApiError(400, 'invalid_args', `Unsupported content type: ${contentType}`);
  }
  const body = await readBody(req);
  let parsed: unknown;
  try {
    parsed = JSON.parse(body.toString('utf8'));
  } catch {
    throw new ApiError(400, 'invalid_args', 'Request is not valid JSON');
  }
  return parseExecuteRequest(parsed);
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

class RemoteServer {
  private readonly executor: Executor;
  private readonly audit: AuditLog;
  private readonly uploads = new FileTokens<StoredFile>();
  private readonly downloads = new FileTokens<Download>();
  /** Requests whose refused upload body is still being discarded; cut off at shutdown. */
  private readonly discarding = new Set<http.IncomingMessage>();

  constructor(private readonly options: ServeOptions) {
    this.executor = new Executor(options.executor);
    this.audit = new AuditLog(options.auditLog);
  }

  async handle(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const started = performance.now();
    const requestId = `req_${randomBytes(8).toString('hex')}`;
    // `URL.parse` instead of `new URL`: llhttp accepts targets the WHATWG parser rejects (e.g. `//[`); they get a 404.
    const url = URL.parse(req.url ?? '/', 'http://localhost');
    const pathname = url?.pathname ?? '';
    const forwarded = req.headers['x-forwarded-for'];
    const entry: AuditEntry = {
      ts: new Date().toISOString(),
      request_id: requestId,
      agent: null,
      src: forwarded === undefined ? null : clip(Array.isArray(forwarded) ? forwarded.join(', ') : forwarded, AUDIT_TEXT_MAX),
      method: req.method ?? '',
      // A file URL's token is a credential until used: it never reaches the log.
      path: FILE_PATH.test(pathname) ? '/v1/files/{token}' : clip(pathname, AUDIT_TEXT_MAX),
      command: null,
      args: [],
      status: 500,
      duration_ms: 0,
    };
    try {
      const body = await this.route(req, res, url, entry);
      entry.status = 200;
      // undefined: the route streamed its own response.
      if (body) json(res, 200, { ok: true, request_id: requestId, ...body });
    } catch (err) {
      const apiError = err instanceof ApiError ? err : new ApiError(500, 'command_failed', err instanceof Error ? err.message : String(err));
      entry.status = apiError.status;
      entry.error = apiError.code;
      if (res.headersSent) res.destroy();
      else json(res, apiError.status, { ok: false, request_id: requestId, error: { code: apiError.code, message: apiError.message } });
    } finally {
      entry.duration_ms = Math.round(performance.now() - started);
      this.audit.write(entry);
    }
  }

  private caller(req: http.IncomingMessage, entry: AuditEntry): Caller {
    try {
      const caller = parseAppCapabilities(req.headersDistinct[APP_CAPS_HEADER], this.options.appCap);
      entry.agent = caller.agent;
      return caller;
    } catch (err) {
      if (err instanceof CapsHeaderError) throw new ApiError(403, 'capability_missing', err.message);
      throw err;
    }
  }

  private async route(
    req: http.IncomingMessage,
    res: http.ServerResponse,
    url: URL | null,
    entry: AuditEntry,
  ): Promise<Record<string, unknown> | undefined> {
    const pathname = url?.pathname ?? '';
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
      return this.execute(caller, await parseBody(req), entry);
    }
    if (req.method === 'PUT' && url && pathname === '/v1/uploads') return this.upload(req, url, entry);
    const file = req.method === 'GET' ? FILE_PATH.exec(pathname) : null;
    if (file) {
      await this.sendFile(req, res, file[1], entry);
      return undefined;
    }
    throw new ApiError(404, 'not_found', `No route for ${req.method} ${entry.path}`);
  }

  /**
   * `PUT /v1/uploads?name=<file name>`: the raw body streamed to a temp file, refused with 413 as soon as it passes
   * `maxFileBytes`. Replies with a single-use token for `{"$upload": "<token>"}`.
   */
  private async upload(req: http.IncomingMessage, url: URL, entry: AuditEntry): Promise<Record<string, unknown>> {
    // Any usable grant: parseAppCapabilities refuses grants without a capability.
    this.caller(req, entry);
    const max = this.options.maxFileBytes;
    const tooLarge = () => new ApiError(413, 'invalid_args', `Upload exceeds the ${megabytes(max)} file limit`);
    // A refused body is read to its end and discarded, never stored, while the error reply goes out at once.
    // Closing instead would reset the connection under a client still sending: it would see a network error, not
    // the 413. A 'data' listener keeps the stream flowing once the iterator below has let go of it.
    const discardRest = () => {
      this.discarding.add(req);
      req.on('data', () => {}).once('close', () => this.discarding.delete(req));
    };
    if (Number(req.headers['content-length']) > max) {
      discardRest();
      throw tooLarge();
    }

    const name = safeName(url.searchParams.get('name') ?? '', 'upload');
    const dir = await fsp.mkdtemp(path.join(os.tmpdir(), 'ms365-serve-up-'));
    const target = path.join(dir, name);
    let size = 0;
    try {
      const fh = await fsp.open(target, 'wx', 0o600);
      try {
        // destroyOnReturn: false keeps the connection usable for the 413 when the loop stops early.
        for await (const chunk of req.iterator({ destroyOnReturn: false }) as AsyncIterable<Buffer>) {
          size += chunk.byteLength;
          if (size > max) throw tooLarge();
          await fh.write(chunk);
        }
      } finally {
        await fh.close();
      }
    } catch (err) {
      discardRest();
      await removeDir(dir);
      throw err;
    }
    return { token: this.uploads.add({ dir, path: target, name, size }), name, size };
  }

  /**
   * `GET /v1/files/{token}`: stream an output file once, to a caller holding the producing command's capabilities.
   * The file is deleted as soon as it is opened.
   */
  private async sendFile(req: http.IncomingMessage, res: http.ServerResponse, token: string, entry: AuditEntry): Promise<void> {
    const caller = this.caller(req, entry);
    const file = this.downloads.get(token);
    if (!file) throw new ApiError(404, 'not_found', 'Unknown or already used file URL');
    entry.command = file.command;
    const missing = file.caps.filter(c => !caller.allow.has(c));
    if (missing.length > 0) throw new ApiError(403, 'capability_missing', `Missing capability: ${missing.join(', ')}`);
    this.downloads.take(token);

    let removed = false;
    try {
      const fh = await fsp.open(file.path, 'r');
      try {
        // The open handle keeps the bytes readable; where an open file cannot be removed, it goes after sending.
        removed = await removeDir(file.dir).then(() => true, () => false);
        res.writeHead(200, { 'content-type': file.contentType, 'content-length': file.size });
        await sendChunks(fh, res);
      } finally {
        await fh.close();
      }
    } finally {
      if (!removed) await removeDir(file.dir);
    }
  }

  private async execute(caller: Caller, request: ExecuteRequest, entry: AuditEntry): Promise<Record<string, unknown>> {
    const node = findCommand(this.options.tree, request.command);
    // Caller-chosen text is bounded in the audit log: an unknown command and arbitrary keys are clipped.
    entry.command = node?.path ?? clip(request.command, AUDIT_TEXT_MAX);
    entry.args = Object.keys(request.args).slice(0, AUDIT_ARGS_MAX).map(key => clip(key, AUDIT_ARG_KEY_MAX));
    if (!node) throw new ApiError(400, 'invalid_args', `Unknown command: ${request.command}`);

    const policy = remotePolicy(node.path);
    if (!policy) throw new ApiError(403, 'capability_missing', `${node.path} cannot run remotely`);
    const missing = policy.caps.filter(c => !caller.allow.has(c));
    if (missing.length > 0) throw new ApiError(403, 'capability_missing', `Missing capability: ${missing.join(', ')}`);

    const { dryRun, ...args } = request.args;
    if (dryRun !== undefined && typeof dryRun !== 'boolean') throw new ApiError(400, 'invalid_args', 'dryRun must be a boolean');

    const input = this.takeInput(policy, node, request.positionals);
    let tmp: string | undefined;
    try {
      tmp = await fsp.mkdtemp(path.join(os.tmpdir(), 'ms365-serve-'));
      const outDir = path.join(tmp, 'out');
      await fsp.mkdir(outDir, { mode: 0o700 });
      if (policy.output) {
        // Output paths are the server's: the caller only names the file; bytes come back through `files`.
        const requested = typeof args.out === 'string' ? safeName(args.out, 'download') : undefined;
        delete args.out;
        delete args.force;
        if (requested) args.out = path.join(outDir, requested);
        else if (policy.output === 'stdout') args.out = path.join(outDir, 'download');
      }
      if (node.options.some(o => o.long === '--output')) args.output = 'json';

      let argv: string[];
      try {
        argv = buildArgv(node, { positionals: input.positionals, args });
      } catch (err) {
        if (err instanceof InvalidInvocationError) throw new ApiError(400, 'invalid_args', err.message);
        throw err;
      }
      const result = await this.executor.run([...(dryRun ? ['--dry-run'] : []), ...argv], outDir);
      if (result.status === 'timeout') throw new ApiError(504, 'timeout', result.message);
      if (result.status === 'failed') throw new ApiError(502, 'command_failed', result.message);

      const body: Record<string, unknown> = { data: parseStdout(result.stdout) };
      if (policy.output) body.files = await this.publishOutputs(outDir, node.path, policy.caps);
      if (result.stderr.trim() !== '') body.stderr = result.stderr;
      return body;
    } finally {
      if (tmp) await removeDir(tmp);
      if (input.upload) await removeDir(input.upload.dir);
    }
  }

  /** Consume the upload referenced by the command's input positional and put its path in its place. */
  private takeInput(policy: RemoteCommand, node: CommandNode, positionals: unknown[]): { positionals: unknown[]; upload?: StoredFile } {
    const result = [...positionals];
    if (policy.input === undefined || policy.input >= result.length) return { positionals: result };
    const ref = result[policy.input];
    const argName = node.arguments[policy.input]?.name ?? 'file';
    const token = typeof ref === 'object' && ref !== null && '$upload' in ref ? ref.$upload : undefined;
    if (typeof token !== 'string') {
      throw new ApiError(400, 'invalid_args', `positional <${argName}> must reference an upload from PUT /v1/uploads: {"$upload": "<token>"}`);
    }
    const upload = this.uploads.take(token);
    if (!upload) throw new ApiError(400, 'invalid_args', 'Unknown or already used upload token');
    result[policy.input] = upload.path;
    return { positionals: result, upload };
  }

  /**
   * Move the child's output files out of the request dir and behind one-time URLs. Sizes are checked first; the
   * executor already stops a child that writes more.
   */
  private async publishOutputs(dir: string, command: string, caps: readonly Capability[]): Promise<OutputFile[]> {
    const max = this.options.maxFileBytes;
    const entries = (await fsp.readdir(dir, { withFileTypes: true })).filter(e => e.isFile());
    let total = 0;
    for (const entry of entries) total += (await fsp.stat(path.join(dir, entry.name))).size;
    if (total > max) throw new ApiError(502, 'command_failed', `Output file exceeds the ${megabytes(max)} remote limit`);
    const files: OutputFile[] = [];
    for (const entry of entries) {
      const own = await fsp.mkdtemp(path.join(os.tmpdir(), 'ms365-serve-dl-'));
      const target = path.join(own, entry.name);
      await fsp.rename(path.join(dir, entry.name), target);
      const { size } = await fsp.stat(target);
      const contentType = detectMimeType(entry.name);
      const token = this.downloads.add({ dir: own, path: target, name: entry.name, size, contentType, caps, command });
      files.push({ name: entry.name, contentType, size, url: `/v1/files/${token}` });
    }
    return files;
  }

  /** Resolves once no child is running; then cuts off refused uploads still sending, which hold their connection. */
  async idle(): Promise<void> {
    await this.executor.idle();
    for (const req of this.discarding) req.socket.destroy();
  }

  /** Close the audit log and remove every upload and output file nobody fetched. */
  async close(): Promise<void> {
    this.audit.close();
    await Promise.all([this.uploads.clear(), this.downloads.clear()]);
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

/** Service log line (listening, signal, stopped): always on stderr; `-q` only silences confirmations. */
function lifecycle(message: string): void {
  console.error(message);
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
    await app.close();
    throw err;
  }

  // Handlers go in before the listening line: a supervisor may signal as soon as it reads it, and an unhandled
  // SIGTERM kills the process without the graceful stop.
  const signal = Promise.withResolvers<NodeJS.Signals>();
  const onSignal = (sig: NodeJS.Signals) => signal.resolve(sig);
  process.once('SIGTERM', onSignal);
  process.once('SIGINT', onSignal);

  if (socket) {
    fs.chmodSync(socket, 0o600);
    lifecycle(`ms365 serve listening on unix:${socket}`);
  } else {
    const address = server.address() as net.AddressInfo;
    const hostPart = address.family === 'IPv6' ? `[${address.address}]` : address.address;
    console.error(
      `Warning: bound to TCP ${hostPart}:${address.port}. Every local process can connect and send its own ` +
      `${APP_CAPS_HEADER} header; prefer --socket.`,
    );
    lifecycle(`ms365 serve listening on http://${hostPart}:${address.port}`);
  }

  const received = await signal.promise;
  process.off('SIGTERM', onSignal);
  process.off('SIGINT', onSignal);

  lifecycle(`Received ${received}; finishing ${inflight.size} request(s)`);
  stopping = true;
  const closed = Promise.withResolvers<void>();
  server.close(() => closed.resolve());
  server.closeIdleConnections();
  await Promise.allSettled(inflight);
  await app.idle();
  await closed.promise;
  if (socket) fs.rmSync(socket, { force: true });
  await app.close();
  lifecycle('ms365 serve stopped');
}
