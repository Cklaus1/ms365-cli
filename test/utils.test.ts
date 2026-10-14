import { describe, it, expect, vi, afterEach } from 'vitest';
import { EventEmitter } from 'events';
import { PassThrough } from 'stream';
import readline from 'readline';
import { parseDate, truncate, validateEmail, parseRecipients, formatSize, buildODataParams, createPrompter, confirm } from '../src/utils.js';
import { CliError, ExitCode } from '../src/errors.js';

/** Run fn and return the error it throws (or the rejection of the promise it returns). */
async function caught(fn: () => unknown): Promise<CliError> {
  try { await fn(); } catch (e) { return e as CliError; }
  throw new Error('expected a throw');
}

describe('parseDate', () => {
  it('accepts valid YYYY-MM-DD dates', () => {
    expect(parseDate('2024-12-25')).toBe('2024-12-25');
    expect(parseDate('2025-01-01')).toBe('2025-01-01');
  });

  it('rejects invalid date strings', () => {
    expect(() => parseDate('not-a-date')).toThrow('Invalid date');
    expect(() => parseDate('2024-13-01')).toThrow('Invalid date');
    expect(() => parseDate('')).toThrow('Invalid date');
  });

  it('rejects partial dates (V-13)', () => {
    expect(() => parseDate('2024')).toThrow('Invalid date');
    expect(() => parseDate('2024-12')).toThrow('Invalid date');
  });

  it('rejects impossible dates instead of rolling them over (V-13)', () => {
    expect(() => parseDate('2026-02-30')).toThrow('Invalid date');
    expect(() => parseDate('2025-02-29')).toThrow('Invalid date');
    expect(parseDate('2024-02-29')).toBe('2024-02-29');
  });

  it('throws a usage CliError (exit 2) (V-13)', async () => {
    const e = await caught(() => parseDate('2026-02-30'));
    expect(e).toBeInstanceOf(CliError);
    expect(e.code).toBe('usage');
    expect(e.exitCode).toBe(ExitCode.USAGE);
  });
});

describe('truncate', () => {
  it('returns short strings unchanged', () => {
    expect(truncate('hello', 10)).toBe('hello');
  });

  it('truncates at exact boundary', () => {
    expect(truncate('hello', 5)).toBe('hello');
  });

  it('adds ellipsis for long strings', () => {
    const result = truncate('hello world', 8);
    expect(result).toBe('hello w…');
    expect(result.length).toBe(8);
  });

  it('handles single-char max', () => {
    expect(truncate('hello', 1)).toBe('…');
  });
});

describe('validateEmail', () => {
  it('accepts valid emails', () => {
    expect(validateEmail('user@example.com')).toBe('user@example.com');
    expect(validateEmail('a@b.co')).toBe('a@b.co');
  });

  it('trims whitespace', () => {
    expect(validateEmail('  user@example.com  ')).toBe('user@example.com');
  });

  it('rejects invalid emails', () => {
    expect(() => validateEmail('not-email')).toThrow('Invalid email');
    expect(() => validateEmail('@no-local.com')).toThrow('Invalid email');
    expect(() => validateEmail('no-domain@')).toThrow('Invalid email');
    expect(() => validateEmail('')).toThrow('Invalid email');
  });

  it('throws a usage CliError (exit 2) (V-11)', async () => {
    const e = await caught(() => validateEmail('bad'));
    expect(e).toBeInstanceOf(CliError);
    expect(e.code).toBe('usage');
    expect(e.exitCode).toBe(ExitCode.USAGE);
  });
});

describe('parseRecipients', () => {
  it('parses single email', () => {
    const result = parseRecipients('user@example.com');
    expect(result).toHaveLength(1);
    expect(result[0].emailAddress.address).toBe('user@example.com');
  });

  it('parses comma-separated emails', () => {
    const result = parseRecipients('a@b.com, c@d.com');
    expect(result).toHaveLength(2);
    expect(result[0].emailAddress.address).toBe('a@b.com');
    expect(result[1].emailAddress.address).toBe('c@d.com');
  });

  it('throws on invalid email in list', () => {
    expect(() => parseRecipients('a@b.com, invalid')).toThrow('Invalid email');
  });

  it('throws a usage CliError (exit 2) (V-11)', async () => {
    const e = await caught(() => parseRecipients('a@b.com, invalid'));
    expect(e.exitCode).toBe(ExitCode.USAGE);
  });
});

describe('createPrompter (V-16)', () => {
  const quietOutput = () => new PassThrough();

  it('answers successive questions from piped input written in one chunk', async () => {
    const input = new PassThrough();
    const p = createPrompter({ input, output: quietOutput() });
    input.end('2\nn\nabc\n');
    expect(await p.ask('First? ')).toBe('2');
    expect(await p.ask('Second? ')).toBe('n');
    expect(await p.ask('Third? ')).toBe('abc');
    p.close();
  });

  it('rejects with a usage error once stdin closes before an answer', async () => {
    const input = new PassThrough();
    const p = createPrompter({ input, output: quietOutput() });
    input.end('only\n');
    expect(await p.ask('One? ')).toBe('only');
    const e = await caught(() => p.ask('Two? '));
    expect(e).toBeInstanceOf(CliError);
    expect(e.exitCode).toBe(ExitCode.USAGE);
    expect(e.message).toContain('Two?');
    expect(e.message).toContain('stdin closed');
  });

  it('rejects a pending question when stdin closes, using onClose', async () => {
    const input = new PassThrough();
    const p = createPrompter({ input, output: quietOutput(), onClose: (q) => new CliError(`setup needs an answer for "${q}" (stdin closed)`, { code: 'usage', exitCode: ExitCode.USAGE }) });
    const pending = p.ask('Client ID: ');
    input.end();
    const e = await caught(() => pending);
    expect(e.message).toBe('setup needs an answer for "Client ID:" (stdin closed)');
  });

  it('works with a bare {question, close} readline stub', async () => {
    const answers = ['a', 'b'];
    const close = vi.fn();
    vi.spyOn(readline, 'createInterface').mockReturnValue({
      question: (_q: string, cb: (a: string) => void) => cb(answers.shift()!),
      close,
    } as never);
    const p = createPrompter();
    expect(await p.ask('1? ')).toBe('a');
    expect(await p.ask('2? ')).toBe('b');
    p.close();
    expect(close).toHaveBeenCalledOnce();
  });
});

describe('confirm', () => {
  const ttyDesc = Object.getOwnPropertyDescriptor(process.stdin, 'isTTY');
  afterEach(() => {
    vi.restoreAllMocks();
    if (ttyDesc) Object.defineProperty(process.stdin, 'isTTY', ttyDesc); else delete (process.stdin as { isTTY?: boolean }).isTTY;
  });
  const setTTY = (v: boolean) => Object.defineProperty(process.stdin, 'isTTY', { value: v, configurable: true });
  const stubAnswer = (answer: string) => vi.spyOn(readline, 'createInterface').mockReturnValue({
    question: (_q: string, cb: (a: string) => void) => cb(answer),
    close: () => {},
  } as never);

  it('refuses on non-TTY stdin (exit 7)', async () => {
    setTTY(false);
    expect((await caught(() => confirm('Delete?'))).exitCode).toBe(ExitCode.CONFIRMATION_REFUSED);
  });

  it('resolves true on y', async () => {
    setTTY(true);
    stubAnswer(' Y ');
    await expect(confirm('Delete?')).resolves.toBe(true);
  });

  it('refuses on anything else (exit 7)', async () => {
    setTTY(true);
    stubAnswer('n');
    expect((await caught(() => confirm('Delete?'))).exitCode).toBe(ExitCode.CONFIRMATION_REFUSED);
  });

  it('refuses (exit 7) when stdin closes before an answer (Ctrl-D)', async () => {
    setTTY(true);
    const rl = Object.assign(new EventEmitter(), { question: () => {}, close: () => {} });
    vi.spyOn(readline, 'createInterface').mockReturnValue(rl as never);
    const pending = confirm('Delete?');
    rl.emit('close');
    expect((await caught(() => pending)).exitCode).toBe(ExitCode.CONFIRMATION_REFUSED);
  });
});

describe('formatSize', () => {
  it('formats bytes', () => {
    expect(formatSize(500)).toBe('500 B');
  });

  it('formats kilobytes', () => {
    expect(formatSize(1024)).toBe('1.0 KB');
    expect(formatSize(1536)).toBe('1.5 KB');
  });

  it('formats megabytes', () => {
    expect(formatSize(1024 * 1024)).toBe('1.0 MB');
  });

  it('formats gigabytes', () => {
    expect(formatSize(1024 * 1024 * 1024)).toBe('1.0 GB');
    expect(formatSize(2.5 * 1024 * 1024 * 1024)).toBe('2.5 GB');
  });

  it('formats zero', () => {
    expect(formatSize(0)).toBe('0 B');
  });
});

describe('buildODataParams', () => {
  it('maps option keys to OData params', () => {
    const opts = { top: '10', filter: "isRead eq false", select: undefined };
    const params = buildODataParams(opts, {
      top: '$top',
      filter: '$filter',
      select: '$select',
    });
    expect(params).toEqual({ '$top': '10', '$filter': 'isRead eq false' });
    expect(params['$select']).toBeUndefined();
  });

  it('returns empty object when no options match', () => {
    expect(buildODataParams({}, { top: '$top' })).toEqual({});
  });
});
