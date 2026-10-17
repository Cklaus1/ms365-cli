import crypto from 'crypto';
import fs from 'fs';
import { CliError, ExitCode, setAfterSuccess } from './errors.js';
import { fromGet } from './utils.js';
import type { GraphClient } from './graph-client.js';

/**
 * Plan/apply (AH-5, OQ-7, E8). `--plan <file>` runs a write command without writing: reads still go to Graph (the
 * target's version), every non-GET request is recorded instead of sent, and the plan is written to `file` (0600).
 * `ms365 apply <file> --expect-digest <d>` sends exactly the recorded requests later.
 */

export interface PlanRequest {
  method: string;
  /** Path and query under the API base; `{{n.id}}` stands for the id request n returned. */
  path: string;
  headers: Record<string, string>;
  /** The exact bytes on the wire, or null. */
  body: string | null;
}

export interface PlanAttachment { name: string; size: number; sha256: string; path: string }

export interface PlanTarget { id: string; version: string }

export interface Plan {
  v: 1;
  command: string;
  identity: { tenantId: string | null; accountId: string | null; profile: string | null };
  target: PlanTarget | null;
  payload: unknown;
  attachments: PlanAttachment[];
  requests: PlanRequest[];
  createdAt: string;
  expiresAt: string;
  digest: string;
}

/** Commands that accept `--plan`. */
export const PLAN_COMMANDS: Record<string, true> = {
  'mail draft': true,
  'mail update-draft': true,
  'mail send-draft': true,
  'mail send': true,
  'mail reply': true,
  'mail forward': true,
  'calendar create': true,
  'calendar update': true,
  'calendar delete': true,
  'calendar cancel': true,
};

/** Plan commands that destroy or cancel: apply asks on a terminal; without one only a host approval applies them (OQ-8 b). */
export const PLAN_CONFIRM: Record<string, true> = {
  'calendar delete': true,
  'calendar cancel': true,
};

/**
 * The plan command a recorded request belongs to, judged from its method and path rather than the plan's own
 * `command` label (a hand-edited plan can carry any label with a matching digest). `null`: no plan command sends it.
 * Apply checks the ceiling and the delete/cancel confirmation against these, so the label cannot widen either.
 */
export function requestCommand(r: Pick<PlanRequest, 'method' | 'path'>): string | null {
  const path = r.path.split('?')[0];
  if (!path.startsWith('/me/')) return null;
  if (/\/(messages|mailFolders)(\/|\(|$)/.test(path)) {
    if (r.method === 'POST' && /\/(send|reply|replyAll|forward)$/.test(path)) return 'mail send-draft';
    return r.method === 'POST' || r.method === 'PATCH' ? 'mail update-draft' : null;
  }
  if (path === '/me/sendMail') return r.method === 'POST' ? 'mail send' : null;
  if (/\/(events|calendars?)(\/|\(|$)/.test(path)) {
    if (r.method === 'DELETE') return 'calendar delete';
    if (r.method === 'POST' && /\/cancel$/.test(path)) return 'calendar cancel';
    return r.method === 'POST' || r.method === 'PATCH' ? 'calendar update' : null;
  }
  return null;
}

export const PLAN_TTL_DEFAULT_MIN = 15;
/** The idempotency ledger's TTL: an applied plan's digest is remembered that long. */
export const PLAN_TTL_MAX_MIN = 24 * 60;

interface Recording {
  requests: PlanRequest[];
  attachments: PlanAttachment[];
  target: PlanTarget | null;
  payload: unknown;
  /** Planned for an agent host (a supplied token or a `type: host` profile). */
  host: boolean;
}

let recording: Recording | null = null;

/** True while a `--plan` run records its writes. */
export function planning(): boolean {
  return recording !== null;
}

/** True while a `--plan` run records its writes for an agent host (AH-7(1): such a plan never takes --series). */
export function hostPlanning(): boolean {
  return recording?.host === true;
}

/** The stand-in for the id that request `index` returns; apply substitutes the real one. */
export const planRef = (index: number): string => `{{${index}.id}}`;

/** Record one write instead of sending it; returns the stand-in response body. */
export function recordRequest(request: PlanRequest): { id: string } {
  if (!recording) throw new Error('recordRequest outside a plan');
  recording.requests.push(request);
  return { id: planRef(recording.requests.length - 1) };
}

export function setPlanTarget(target: PlanTarget): void {
  if (recording) recording.target = target;
}

export function setPlanPayload(payload: unknown): void {
  if (recording) recording.payload = payload;
}

export const sha256Hex = (data: string | Uint8Array): string => crypto.createHash('sha256').update(data).digest('hex');

/**
 * GET an Outlook item (`select` plus changeKey) and its version for If-Match: `@odata.etag`, else `W/"<changeKey>"`
 * (undefined when Graph sent neither). Under --dry-run the version is the `<from GET>` placeholder.
 */
export async function readVersion(client: GraphClient, endpoint: string, select: string): Promise<{ item: unknown; version?: string }> {
  const res = await client.request(endpoint, { queryParams: { $select: select ? `${select},changeKey` : 'changeKey' } });
  if (client.isPreview(res)) return { item: {}, version: fromGet() };
  const item: unknown = res.data;
  const etag: unknown = item && typeof item === 'object' ? Reflect.get(item, '@odata.etag') : undefined;
  const changeKey: unknown = item && typeof item === 'object' ? Reflect.get(item, 'changeKey') : undefined;
  if (typeof etag === 'string' && etag) return { item, version: etag };
  return { item, version: typeof changeKey === 'string' && changeKey ? `W/"${changeKey}"` : undefined };
}

/** `--if-match`: an explicit etag; `auto`, or any --plan run, → the version `read` returns (it must have one); else none. */
export async function ifMatchFor(flag: string | undefined, read: () => Promise<string | undefined>): Promise<string | undefined> {
  if (flag !== 'auto' && (flag !== undefined || !planning())) return flag;
  const version = await read();
  if (version) return version;
  throw new CliError('The item came back without a version (@odata.etag), so the change cannot be made conditional.', { code: 'graph_error' });
}

/** sha256 of each file's bytes (AH-5(4)); recorded on the plan when planning. */
export function attachmentDigests(files: { filePath: string; name: string; size: number }[]): PlanAttachment[] {
  const list = files.map(f => ({ name: f.name, size: f.size, sha256: sha256Hex(fs.readFileSync(f.filePath)), path: f.filePath }));
  if (recording) recording.attachments = list;
  return list;
}

function canonical(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonical);
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.keys(value).sort().map(k => [k, canonical((value as Record<string, unknown>)[k])]));
  }
  return value;
}

/** Canonical JSON (OQ-7): object keys sorted, UTF-8, no whitespace. */
export function canonicalJson(value: unknown): string {
  return JSON.stringify(canonical(value));
}

/** sha256 over the canonical plan without its `digest` field. */
export function planDigest(plan: Omit<Plan, 'digest'> & { digest?: string }): string {
  const { digest: _omit, ...rest } = plan;
  return sha256Hex(canonicalJson(rest));
}

export function usage(message: string, code = 'usage'): CliError {
  return new CliError(message, { code, exitCode: ExitCode.USAGE });
}

/** `--plan-ttl` in minutes: 1 … 1440 (OQ-7). */
export function planTtl(raw: string | undefined): number {
  if (raw === undefined) return PLAN_TTL_DEFAULT_MIN;
  const n = Number(raw);
  if (!/^\d+$/.test(raw) || n < 1 || n > PLAN_TTL_MAX_MIN) throw usage(`--plan-ttl must be whole minutes from 1 to ${PLAN_TTL_MAX_MIN}, got "${raw}".`);
  return n;
}

/**
 * Start recording for `command`; after the action succeeds the plan is written to `file` (0600) and a summary
 * `{plan, command, digest, expiresAt, requests}` goes to stdout. Refused for commands without plan support.
 */
export function beginPlan(opts: { file: string; command: string; ttlMinutes: number; profile: string | null; host: boolean; client: () => GraphClient }): void {
  if (!PLAN_COMMANDS[opts.command]) {
    throw usage(`--plan is not supported by "${opts.command}". Plan-capable: ${Object.keys(PLAN_COMMANDS).join(', ')}.`);
  }
  recording = { requests: [], attachments: [], target: null, payload: null, host: opts.host };
  setAfterSuccess(async () => {
    const rec = recording!;
    endPlan();
    if (rec.requests.length === 0) throw usage(`"${opts.command}" made no write to plan.`);
    const identity = await opts.client().identity();
    const now = new Date();
    const body: Omit<Plan, 'digest'> = {
      v: 1,
      command: opts.command,
      identity: { tenantId: identity.tenantId, accountId: identity.accountId, profile: opts.profile },
      target: rec.target,
      payload: rec.payload,
      attachments: rec.attachments,
      requests: rec.requests,
      createdAt: now.toISOString(),
      expiresAt: new Date(now.getTime() + opts.ttlMinutes * 60_000).toISOString(),
    };
    const plan: Plan = { ...body, digest: planDigest(body) };
    // A new 0600 file renamed into place: writing over an existing 0644 file (or a symlink) would expose the body.
    const tmp = `${opts.file}.${process.pid}.${crypto.randomBytes(4).toString('hex')}.tmp`;
    try {
      fs.writeFileSync(tmp, canonicalJson(plan) + '\n', { mode: 0o600, flag: 'wx' });
      fs.renameSync(tmp, opts.file);
    } catch (err) {
      fs.rmSync(tmp, { force: true });
      throw usage(`Cannot write plan ${opts.file}: ${err instanceof Error ? err.message : String(err)}`);
    }
    console.log(JSON.stringify({ plan: opts.file, command: plan.command, digest: plan.digest, expiresAt: plan.expiresAt, requests: plan.requests.length }, null, 2));
  });
}

/** Stop recording and drop the after-success hook. */
export function endPlan(): void {
  recording = null;
  setAfterSuccess(null);
}

const objectWith = (v: unknown, ...strings: string[]): boolean =>
  typeof v === 'object' && v !== null && strings.every(k => typeof Reflect.get(v, k) === 'string');

/** Read and shape-check a plan file (the fields apply uses); anything else exits 2 `plan_invalid`. */
export function readPlan(file: string): Plan {
  let parsed: unknown;
  try {
    parsed = JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch (err) {
    throw usage(`Cannot read plan ${file}: ${err instanceof Error ? err.message : String(err)}`, 'plan_invalid');
  }
  const ok = objectWith(parsed, 'command', 'digest', 'createdAt', 'expiresAt')
    && typeof parsed === 'object' && parsed !== null && 'v' in parsed && parsed.v === 1
    && 'identity' in parsed && typeof parsed.identity === 'object' && parsed.identity !== null
    && 'requests' in parsed && Array.isArray(parsed.requests)
    && parsed.requests.every(r => objectWith(r, 'method', 'path') && 'headers' in r && typeof r.headers === 'object' && r.headers !== null
      && 'body' in r && (r.body === null || typeof r.body === 'string'))
    && 'attachments' in parsed && Array.isArray(parsed.attachments)
    && parsed.attachments.every(a => objectWith(a, 'name', 'sha256', 'path') && 'size' in a && typeof a.size === 'number');
  if (!ok) throw usage(`${file} is not an ms365 plan (v1).`, 'plan_invalid');
  return parsed as Plan;
}

/** Replace `{{n.id}}` (raw in bodies, URL-encoded in paths) with the ids earlier requests returned. */
export function resolveRefs(text: string, ids: (string | undefined)[], encode: boolean): string {
  return ids.reduce<string>((out, id, i) => {
    if (id === undefined) return out;
    const ref = planRef(i);
    return out.split(encode ? encodeURIComponent(ref) : ref).join(encode ? encodeURIComponent(id) : id);
  }, text);
}
