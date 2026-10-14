import { Command, InvalidArgumentError, Option } from 'commander';
import type { GraphClient } from './graph-client.js';

/**
 * Agent-friendly message content (spec §6): HTML→text/markdown conversion, body truncation,
 * and untrusted-content marking for mail and Teams messages.
 */

// ── HTML conversion ─────────────────────────────────────────────

const NAMED_ENTITIES: Record<string, string> = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ',
  copy: '©', reg: '®', trade: '™', euro: '€', pound: '£', yen: '¥', cent: '¢', sect: '§', deg: '°',
  hellip: '…', mdash: '—', ndash: '–', lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”',
  laquo: '«', raquo: '»', bull: '•', middot: '·', times: '×', divide: '÷', zwnj: '‌', zwj: '‍',
};

export function decodeEntities(text: string): string {
  return text.replace(/&(#\d+|#x[0-9a-f]+|[a-z][a-z0-9]*);/gi, (whole, ref: string) => {
    if (ref[0] !== '#') return NAMED_ENTITIES[ref.toLowerCase()] ?? whole;
    const code = ref[1] === 'x' || ref[1] === 'X' ? parseInt(ref.slice(2), 16) : parseInt(ref.slice(1), 10);
    return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : whole;
  });
}

// Sentinels that survive whitespace normalization: list indentation and spaces inside <pre>.
const INDENT = '\u0001';
const PRE_SPACE = '\u0002';

const PARAGRAPH_TAGS = new Set(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'table', 'dl', 'figure', 'address']);
const LINE_TAGS = new Set(['div', 'tr', 'dt', 'dd', 'section', 'article', 'header', 'footer', 'main', 'nav', 'aside', 'center', 'form', 'fieldset', 'caption']);
const BUFFERED_TAGS = new Set(['a', 'b', 'strong', 'i', 'em', 'code', 'blockquote']);

interface Frame { tag: string; out: string; href?: string }

function attr(attrs: string, name: string): string | undefined {
  const m = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s"'>]+))`, 'i').exec(attrs);
  return m ? decodeEntities(m[1] ?? m[2] ?? m[3]) : undefined;
}

/** Collapse spaces around line breaks and runs of blank lines. */
function normalize(text: string): string {
  return text.replace(/[ \t]*\n[ \t]*/g, '\n').replace(/\n{3,}/g, '\n\n').replace(/^\n+|\n+$/g, '');
}

function convertHtml(html: string, markdown: boolean): string {
  const source = html
    .replace(/<!--[\s\S]*?(?:-->|$)/g, '')
    .replace(/<(script|style|noscript|template)\b[^>]*>[\s\S]*?(?:<\/\1\s*>|$)/gi, '')
    .replace(/<head\b[^>]*>[\s\S]*?<\/head\s*>/gi, '');

  const stack: Frame[] = [{ tag: '', out: '' }];
  const lists: Array<{ ordered: boolean; n: number }> = [];
  const cells: number[] = [];
  let preDepth = 0;

  const top = () => stack[stack.length - 1];
  const emit = (s: string) => { top().out += s; };
  const lastChar = (): string => {
    for (let i = stack.length - 1; i >= 0; i--) {
      if (stack[i].out) return stack[i].out[stack[i].out.length - 1];
    }
    return '\n';
  };
  const atLineStart = () => ' \n'.includes(lastChar()) || lastChar() === INDENT;
  // Inside a list, blocks only break the line (never right after a marker) so items stay compact.
  const newline = () => { if (lastChar() !== '\n') emit('\n'); };
  const block = (breaks: string) => {
    if (lists.length > 0) { if (!atLineStart()) emit('\n'); } else if (breaks === '\n') newline();
    else emit(breaks);
  };

  const emitText = (raw: string) => {
    const text = decodeEntities(raw);
    if (preDepth > 0) { emit(text.replace(/[ \t]/g, PRE_SPACE)); return; }
    let t = text.replace(/[ \t\r\n\f]+/g, ' ');
    if (atLineStart()) t = t.replace(/^ /, '');
    emit(t);
  };

  // Close the innermost buffered frame and merge its rendering into the parent.
  const closeFrame = () => {
    const frame = stack.pop()!;
    const [, lead, core, trail] = /^(\s*)([\s\S]*?)(\s*)$/.exec(frame.out)!;
    let rendered: string;
    switch (frame.tag) {
      case 'a': {
        const href = frame.href && !/^\s*javascript:/i.test(frame.href) ? frame.href : undefined;
        const label = core.replace(/\s*\n\s*/g, ' ');
        if (!href) rendered = label;
        else if (!label) rendered = href;
        else if (markdown) rendered = `[${label}](${href})`;
        else rendered = label === href || `mailto:${label}` === href ? label : `${label} (${href})`;
        break;
      }
      case 'b': case 'strong':
        rendered = markdown && core ? `**${core}**` : core;
        break;
      case 'i': case 'em':
        rendered = markdown && core ? `*${core}*` : core;
        break;
      case 'code':
        rendered = markdown && core && preDepth === 0 ? `\`${core}\`` : core;
        break;
      default: { // blockquote
        const body = normalize(frame.out);
        rendered = `\n\n${body.split('\n').map(l => (l ? `> ${l}` : '>')).join('\n')}\n\n`;
        emit(rendered);
        return;
      }
    }
    emit(lead + rendered + trail);
  };

  const closeTag = (tag: string) => {
    if (BUFFERED_TAGS.has(tag)) {
      const idx = stack.map(f => f.tag).lastIndexOf(tag);
      if (idx > 0) while (stack.length > idx) closeFrame();
      return;
    }
    if (tag === 'ul' || tag === 'ol') {
      if (lists.length > 0) lists.pop();
      if (lists.length === 0) emit('\n\n');
    } else if (tag === 'pre') {
      if (preDepth > 0) preDepth--;
      emit(markdown ? '\n```\n\n' : '\n\n');
    } else if (tag === 'tr') {
      cells.pop();
      newline();
    } else if (/^h[1-6]$/.test(tag) || PARAGRAPH_TAGS.has(tag)) {
      block('\n\n');
    } else if (LINE_TAGS.has(tag)) {
      block('\n');
    }
  };

  const openTag = (tag: string, attrs: string) => {
    if (BUFFERED_TAGS.has(tag)) {
      if (tag === 'blockquote') emit('\n\n');
      stack.push({ tag, out: '', href: tag === 'a' ? attr(attrs, 'href') : undefined });
      return;
    }
    switch (tag) {
      case 'br': emit('\n'); return;
      case 'hr': emit(markdown ? '\n\n---\n\n' : '\n\n'); return;
      case 'img': { const alt = attr(attrs, 'alt'); if (alt) emitText(alt); return; }
      case 'ul': case 'ol':
        if (lists.length === 0) emit('\n\n');
        lists.push({ ordered: tag === 'ol', n: 0 });
        return;
      case 'li': {
        const list = lists[lists.length - 1];
        const marker = list?.ordered ? `${++list.n}. ` : '- ';
        newline();
        emit(INDENT.repeat(Math.max(lists.length - 1, 0)) + marker);
        return;
      }
      case 'pre':
        emit(markdown ? '\n\n```\n' : '\n\n');
        preDepth++;
        return;
      case 'tr': cells.push(0); newline(); return;
      case 'td': case 'th':
        if (cells.length > 0 && cells[cells.length - 1]++ > 0) emit(' | ');
        return;
    }
    const heading = /^h([1-6])$/.exec(tag);
    if (heading) { emit('\n\n'); if (markdown) emit('#'.repeat(Number(heading[1])) + ' '); return; }
    if (PARAGRAPH_TAGS.has(tag)) block('\n\n');
    else if (LINE_TAGS.has(tag)) block('\n');
  };

  const token = /<!\[CDATA\[([\s\S]*?)\]\]>|<[!?][^>]*>|<(\/?)([a-zA-Z][\w:-]*)((?:[^>"']|"[^"]*"|'[^']*')*)>/g;
  let pos = 0;
  for (let m = token.exec(source); m; m = token.exec(source)) {
    if (m.index > pos) emitText(source.slice(pos, m.index));
    pos = token.lastIndex;
    if (m[1] !== undefined) { emitText(m[1]); continue; }
    if (!m[3]) continue; // doctype, processing instruction
    const tag = m[3].toLowerCase();
    if (m[2]) closeTag(tag);
    else openTag(tag, m[4]);
  }
  if (pos < source.length) emitText(source.slice(pos));
  while (stack.length > 1) closeFrame();

  return normalize(stack[0].out.replace(/ /g, ' '))
    .replace(new RegExp(INDENT, 'g'), '  ')
    .replace(new RegExp(PRE_SPACE, 'g'), ' ');
}

/** Convert HTML to readable plain text (links as "text (url)"). */
export function htmlToText(html: string): string {
  return convertHtml(html, false);
}

/** Convert HTML to lightweight markdown: headings, emphasis, links, lists, quotes, code. */
export function htmlToMarkdown(html: string): string {
  return convertHtml(html, true);
}

// ── Body shaping ────────────────────────────────────────────────

export type BodyMode = 'html' | 'text' | 'markdown' | 'none';

export interface ContentOptions {
  body?: BodyMode;
  maxChars?: number;
  wrapUntrusted?: boolean;
}

/** Truncate to `max` characters with a marker saying how much was dropped. */
export function truncateContent(text: string, max: number): string {
  if (text.length <= max) return text;
  let cut = max;
  const code = text.charCodeAt(cut - 1);
  if (code >= 0xd800 && code <= 0xdbff) cut--;
  return `${text.slice(0, cut)}…[truncated ${text.length - cut} chars]`;
}

export function wrapUntrusted(text: string): string {
  return `<<<UNTRUSTED_CONTENT\n${text}\nUNTRUSTED_CONTENT>>>`;
}

function shapeText(text: string, opts: ContentOptions): string {
  const cut = opts.maxChars ? truncateContent(text, opts.maxChars) : text;
  return opts.wrapUntrusted ? wrapUntrusted(cut) : cut;
}

/** Apply --body / --max-chars / --wrap-untrusted to a message's body, uniqueBody and bodyPreview. */
export function shapeBodies(item: Record<string, unknown>, opts: ContentOptions): Record<string, unknown> {
  const out = { ...item };
  for (const key of ['body', 'uniqueBody']) {
    const body = out[key] as { contentType?: string; content?: unknown } | undefined;
    if (!body || typeof body !== 'object') continue;
    if (opts.body === 'none') { delete out[key]; continue; }
    if (typeof body.content !== 'string') continue;
    let { contentType, content } = body as { contentType?: string; content: string };
    const isHtml = contentType?.toLowerCase() === 'html';
    if (opts.body === 'text' && isHtml) { content = htmlToText(content); contentType = 'text'; }
    if (opts.body === 'markdown' && isHtml) { content = htmlToMarkdown(content); contentType = 'markdown'; }
    out[key] = { ...body, contentType, content: shapeText(content, opts) };
  }
  if (typeof out.bodyPreview === 'string') out.bodyPreview = shapeText(out.bodyPreview, opts);
  return out;
}

// ── Untrusted marking ───────────────────────────────────────────

function domainOf(address: unknown): string | undefined {
  if (typeof address !== 'string') return undefined;
  const at = address.lastIndexOf('@');
  return at >= 0 ? address.slice(at + 1).toLowerCase() : undefined;
}

type Addressed = { emailAddress?: { address?: string } } | undefined;

/** Mark a mail message untrusted; sender.external compares the sender's domain to the user's. */
export function markMailMessage(item: Record<string, unknown>, userDomain: string | undefined): Record<string, unknown> {
  const out: Record<string, unknown> = { ...item, untrusted: true };
  const senderDomain = domainOf((item.from as Addressed)?.emailAddress?.address)
    ?? domainOf((item.sender as Addressed)?.emailAddress?.address);
  if (userDomain && senderDomain) {
    out.sender = { ...(item.sender as Record<string, unknown> | undefined), external: senderDomain !== userDomain.toLowerCase() };
  }
  return out;
}

/** Mark a Teams message untrusted; sender.external compares from.user.tenantId to the user's tenant. */
export function markTeamsMessage(item: Record<string, unknown>, userTenant: string | undefined): Record<string, unknown> {
  const out: Record<string, unknown> = { ...item, untrusted: true };
  const senderTenant = (item.from as { user?: { tenantId?: string } } | undefined)?.user?.tenantId;
  if (userTenant && senderTenant) out.sender = { external: senderTenant.toLowerCase() !== userTenant.toLowerCase() };
  return out;
}

// ── Command wiring ──────────────────────────────────────────────

export type ItemTransform = (items: unknown[]) => Promise<unknown[]>;

function isMessage(item: unknown): item is Record<string, unknown> {
  return !!item && typeof item === 'object' && !Array.isArray(item);
}

/**
 * Build the per-command transform for mail or Teams messages. The signed-in user's domain
 * (mail) or tenant (Teams) is looked up once, lazily — only when there are messages to mark,
 * so --dry-run sends nothing extra. A failed lookup leaves sender.external out.
 */
export function messageTransform(client: GraphClient, kind: 'mail' | 'teams', opts: ContentOptions): ItemTransform {
  let identity: Promise<string | undefined> | undefined;
  const lookup = async (): Promise<string | undefined> => {
    try {
      if (kind === 'mail') {
        const me = (await client.request('/me', { queryParams: { '$select': 'mail,userPrincipalName' } })).data as { mail?: string; userPrincipalName?: string };
        return domainOf(me.mail) ?? domainOf(me.userPrincipalName);
      }
      const org = (await client.request('/organization', { queryParams: { '$select': 'id' } })).data as { value?: Array<{ id?: string }> };
      return org.value?.[0]?.id;
    } catch {
      return undefined;
    }
  };
  return async (items) => {
    // Under --dry-run every item is the request placeholder, not a message.
    if (client.isDryRun() || !items.some(isMessage)) return items;
    const id = await (identity ??= lookup());
    return items.map((item) => {
      if (!isMessage(item)) return item;
      const shaped = shapeBodies(item, opts);
      return kind === 'mail' ? markMailMessage(shaped, id) : markTeamsMessage(shaped, id);
    });
  };
}

/** Apply a transform to a single object or the `value` array of a collection response. */
export async function transformData(data: unknown, transform: ItemTransform): Promise<unknown> {
  if (data && typeof data === 'object' && Array.isArray((data as Record<string, unknown>).value)) {
    return { ...(data as Record<string, unknown>), value: await transform((data as Record<string, unknown>).value as unknown[]) };
  }
  return (await transform([data]))[0];
}

function parseMaxChars(value: string): number {
  const n = Number(value);
  if (!Number.isInteger(n) || n < 1) throw new InvalidArgumentError('must be a positive integer.');
  return n;
}

/** Add --body (unless `withBody` is false), --max-chars and --wrap-untrusted to a read command. */
export function addContentOptions(cmd: Command, withBody = true): Command {
  if (withBody) {
    cmd.addOption(new Option('--body <mode>', 'Body format: html (as returned), text, markdown, or none (drop bodies)')
      .choices(['html', 'text', 'markdown', 'none']).default('html'));
  }
  return cmd
    .option('--max-chars <n>', 'Truncate body content and bodyPreview to n characters', parseMaxChars)
    .option('--wrap-untrusted', 'Wrap body content in <<<UNTRUSTED_CONTENT … UNTRUSTED_CONTENT>>> delimiters');
}

/** The ContentOptions selected by addContentOptions flags. */
export function contentOptions(opts: { body?: BodyMode; maxChars?: number; wrapUntrusted?: boolean }): ContentOptions {
  return { body: opts.body, maxChars: opts.maxChars, wrapUntrusted: opts.wrapUntrusted };
}
