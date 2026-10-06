// `ms365-remote`: runs ms365 commands on the host that holds the Microsoft 365 login (REMOTE-1d / REMOTE-0).
// Argv is parsed with the same Commander tree as `ms365`, so help and usage errors match.
// Configuration comes from real environment variables only (no .env): MS365_REMOTE_URL or MS365_REMOTE_SSH.
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { Command, CommanderError } from 'commander';
import { introspectionProgram } from './commands/registry.js';
import { photoExtension } from './commands/user.js';
import { exitProcess, reportError } from './errors.js';
import { saveOrStream } from './files.js';
import { formatOutput, type OutputFormat } from './formatter.js';
import { commandPolicy } from './serve/capabilities.js';
import { formatSize, info, setQuiet } from './utils.js';

for (const stream of [process.stdout, process.stderr]) {
  stream.on('error', (err: NodeJS.ErrnoException) => {
    if (err.code === 'EPIPE' || err.code === 'ERR_STREAM_DESTROYED') process.exit(0);
  });
}

interface Invocation {
  /** Canonical dotted path, e.g. `mail.list`. */
  command: string;
  positionals: string[];
  /** Options given on the command line, raw strings as typed (booleans for flags). */
  args: Record<string, unknown>;
}

interface RemoteFile {
  name: string;
  contentType?: string;
  base64: string;
}

interface RemoteResponse {
  ok: boolean;
  data?: unknown;
  files?: RemoteFile[];
  stderr?: string;
  error?: { code: string; message: string };
}

/** The `ms365` command tree, renamed, with every leaf action replaced by a capture of its invocation. */
function buildProgram(onInvoke: (invocation: Invocation) => void): Command {
  const program = introspectionProgram()
    .name('ms365-remote')
    .description('Run ms365 commands on the ms365 serve host (MS365_REMOTE_URL) or over Tailscale SSH (MS365_REMOTE_SSH)')
    .version('0.1.0')
    .option('--dry-run', 'Preview the API request without executing it')
    .option('-q, --quiet', 'Suppress non-data output (confirmations, progress)');

  const visit = (cmd: Command, names: string[]) => {
    cmd.exitOverride();
    const children = cmd.commands.filter(c => c.name() !== 'help');
    if (cmd !== program && children.length === 0) {
      // Options with a parser (integers, clock times, choices) are sent as typed; the server's CLI parses them again.
      const raw: Record<string, string | string[]> = {};
      for (const opt of cmd.options) {
        const parse = opt.parseArg;
        if (!parse) continue;
        const key = opt.attributeName();
        opt.parseArg = <T>(value: string, previous: T): T => {
          const prior = raw[key];
          raw[key] = opt.variadic ? [...(Array.isArray(prior) ? prior : []), value] : value;
          return parse(value, previous);
        };
      }
      cmd.action((...actionArgs: unknown[]) => {
        // Commander passes the command itself as the last action argument.
        const leaf = actionArgs.at(-1) as Command;
        const opts = leaf.opts();
        const args: Record<string, unknown> = {};
        for (const opt of leaf.options) {
          const key = opt.attributeName();
          if (key in args || leaf.getOptionValueSource(key) !== 'cli') continue;
          args[key] = raw[key] ?? opts[key];
        }
        // The format is rendered locally, so its default (table on a TTY, json in a pipe) counts too.
        if (typeof opts.output === 'string') args.output = opts.output;
        onInvoke({ command: names.join('.'), positionals: [...leaf.args], args });
      });
      return;
    }
    for (const child of children) visit(child, cmd === program ? [child.name()] : [...names, child.name()]);
  };
  visit(program, []);
  return program;
}

/** Quote one argument for a POSIX shell (the remote login shell of `tailscale ssh`). */
function shellQuote(arg: string): string {
  return `'${arg.replace(/'/g, `'\\''`)}'`;
}

/** Check the `/v1/execute` response shape; anything else is reported with the HTTP status. */
function parseRemoteResponse(value: unknown, status: number): RemoteResponse {
  const invalid = new Error(`ms365 serve answered HTTP ${status} with an unexpected body`);
  if (typeof value !== 'object' || value === null || !('ok' in value) || typeof value.ok !== 'boolean') throw invalid;
  if (!value.ok) {
    const error = 'error' in value ? value.error : undefined;
    if (typeof error !== 'object' || error === null || !('code' in error) || !('message' in error)) throw invalid;
    return { ok: false, error: { code: String(error.code), message: String(error.message) } };
  }
  const files: RemoteFile[] = [];
  const rawFiles = 'files' in value ? value.files : [];
  if (!Array.isArray(rawFiles)) throw invalid;
  for (const f of rawFiles) {
    if (typeof f?.name !== 'string' || typeof f.base64 !== 'string') throw invalid;
    files.push({ name: f.name, contentType: typeof f.contentType === 'string' ? f.contentType : undefined, base64: f.base64 });
  }
  const stderr = 'stderr' in value && typeof value.stderr === 'string' ? value.stderr : undefined;
  return { ok: true, data: 'data' in value ? value.data : null, files, stderr };
}

/**
 * Local name for a returned file when the user gave no file name: the server's choice reduced to its last path
 * segment, refusing empty and dot names (`.bashrc`, `.env`) a malicious server could plant. A profile photo is
 * named `profile-photo.<ext>` from its content type, exactly like the local CLI.
 */
function localFileName(command: string, file: RemoteFile): string {
  if (command === 'user.download-photo') return `profile-photo${photoExtension(file.contentType ?? '')}`;
  const name = path.basename(file.name.replace(/\\/g, '/'));
  if (name === '' || name.startsWith('.')) {
    throw new Error(`Refusing the file name "${file.name}" from ms365 serve. Choose a file name with --out <file>.`);
  }
  return name;
}

/** REMOTE-0: `tailscale ssh <host> -- ms365 <argv>`; stdio and exit code pass through. File paths are the host's. */
async function runSsh(host: string, argv: string[]): Promise<number> {
  const child = spawn('tailscale', ['ssh', host, '--', 'ms365', ...argv.map(shellQuote)], { stdio: 'inherit' });
  const { promise, resolve, reject } = Promise.withResolvers<number>();
  child.once('error', (err) => reject(new Error(`Cannot run tailscale: ${err.message}`)));
  child.once('close', (code) => resolve(code ?? 1));
  return promise;
}

/** REMOTE-1d: POST /v1/execute, uploading the local input file and writing returned output files locally. */
async function runHttp(baseUrl: string, invocation: Invocation, dryRun: boolean): Promise<number> {
  let base: URL;
  try {
    base = new URL(baseUrl);
  } catch {
    throw new Error(`MS365_REMOTE_URL is not a URL: ${baseUrl}`);
  }
  if (base.protocol !== 'http:' && base.protocol !== 'https:') throw new Error('MS365_REMOTE_URL must be an http(s) URL');

  const policy = commandPolicy(invocation.command);
  const remote = policy?.remote ? policy : undefined;
  const { output, out, force, ...args } = invocation.args;
  const format = (typeof output === 'string' ? output : 'json') as OutputFormat;
  if (remote?.output === undefined) {
    if (out !== undefined) args.out = out;
    if (force !== undefined) args.force = force;
  }
  if (dryRun) args.dryRun = true;

  // Local --out is checked up front, like the local command, so nothing is fetched for a file that cannot be written.
  const outPath = typeof out === 'string' ? out : undefined;
  const outIsDir = outPath !== undefined && fs.statSync(outPath, { throwIfNoEntry: false })?.isDirectory() === true;
  if (remote?.output && outPath && !outIsDir && force !== true && fs.existsSync(outPath)) {
    throw new Error(`File exists: ${outPath}. Use --force to overwrite.`);
  }

  const positionals: unknown[] = [...invocation.positionals];
  let upload: string | undefined;
  if (remote?.input !== undefined && remote.input < positionals.length) {
    upload = path.resolve(invocation.positionals[remote.input]);
    if (!fs.statSync(upload, { throwIfNoEntry: false })?.isFile()) throw new Error(`File not found: ${upload}`);
    positionals[remote.input] = { $file: 'file0' };
  }

  const request = { command: invocation.command, positionals, args };
  let body: string | FormData;
  const headers: Record<string, string> = {};
  if (upload) {
    body = new FormData();
    body.append('request', JSON.stringify(request));
    body.append('file0', await fs.openAsBlob(upload), path.basename(upload));
  } else {
    body = JSON.stringify(request);
    headers['content-type'] = 'application/json';
  }

  const endpoint = `${base.href.replace(/\/+$/, '')}/v1/execute`;
  let res: Response;
  try {
    res = await fetch(endpoint, { method: 'POST', headers, body });
  } catch (err) {
    const cause = err instanceof Error && err.cause instanceof Error ? err.cause.message : String(err);
    throw new Error(`Cannot reach ms365 serve at ${base.href}: ${cause}`);
  }
  let parsed: unknown;
  try {
    parsed = await res.json();
  } catch {
    throw new Error(`ms365 serve at ${base.href} answered HTTP ${res.status} without a JSON body`);
  }
  const payload = parseRemoteResponse(parsed, res.status);

  if (!payload.ok) {
    const error = payload.error ?? { code: 'unknown', message: `HTTP ${res.status}` };
    // command_failed carries the remote CLI's own stderr ("Error: …"); other codes come from the server.
    console.error(error.code === 'command_failed' ? error.message : `Error: ${error.message} (${error.code})`);
    return 1;
  }

  if (payload.stderr) info(payload.stderr.replace(/\n$/, ''));

  const file = payload.files?.[0];
  if (remote?.output && file) {
    const bytes = Buffer.from(file.base64, 'base64');
    let target = outPath;
    if (outIsDir) target = path.join(outPath!, localFileName(invocation.command, file));
    else if (target === undefined && remote.output === 'named') target = localFileName(invocation.command, file);
    const saved = await saveOrStream(bytes, { out: target, force: force === true });
    if (saved.path) info(`Saved to ${path.resolve(saved.path)} (${formatSize(saved.bytes)})`);
  }

  if (typeof payload.data === 'string') process.stdout.write(payload.data);
  else if (payload.data !== null && payload.data !== undefined) console.log(formatOutput(payload.data, format));
  return 0;
}

async function main(): Promise<number> {
  let invocation: Invocation | undefined;
  const program = buildProgram((inv) => {
    invocation = inv;
  });
  try {
    await program.parseAsync(process.argv);
  } catch (err) {
    // --help / --version throw with exitCode 0; usage errors were already printed by Commander.
    if (err instanceof CommanderError) return err.exitCode;
    reportError(err);
    return 1;
  }
  if (!invocation) return 0;

  const globals = program.opts<{ dryRun?: boolean; quiet?: boolean }>();
  if (globals.quiet) setQuiet(true);
  try {
    const url = process.env.MS365_REMOTE_URL;
    const sshHost = process.env.MS365_REMOTE_SSH;
    if (url) return await runHttp(url, invocation, Boolean(globals.dryRun));
    if (sshHost) return await runSsh(sshHost, process.argv.slice(2));
    throw new Error('Set MS365_REMOTE_URL (ms365 serve URL, e.g. https://ms365-host.<tailnet>.ts.net) or MS365_REMOTE_SSH (Tailscale SSH host).');
  } catch (err) {
    reportError(err);
    return 1;
  }
}

main().then(exitProcess);
