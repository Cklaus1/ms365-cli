import { describe, it, expect } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import {
  encodePath, odataQuote, kqlQuote, safeOutputPath, resolveTimeZone, preferTimeZone,
  parseGraphDateTime, dayRange, toGraphDateTime, toIsoWithOffset, parseTime, wallTimeToDate,
} from '../src/utils.js';
import { CliError, ExitCode } from '../src/errors.js';

const LA = 'America/Los_Angeles';

describe('path and query helpers', () => {
  it('encodePath keeps slashes and encodes # ? % & spaces', () => {
    expect(encodePath('Docs/Q#1 50%?.pdf')).toBe('Docs/Q%231%2050%25%3F.pdf');
    expect(encodePath('R&D/plan.md')).toBe('R%26D/plan.md');
  });
  it('odataQuote doubles single quotes', () => {
    expect(odataQuote("'O'Brien'!A1")).toBe("''O''Brien''!A1");
  });
  it('kqlQuote wraps and escapes', () => {
    expect(kqlQuote('say "hi"')).toBe('"say \\"hi\\""');
  });
});

describe('safeOutputPath (M-1)', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-safe-'));
  const cwd = process.cwd();
  it('strips traversal and control characters', () => {
    process.chdir(dir);
    try {
      expect(safeOutputPath('../../.bashrc')).toBe(path.join(fs.realpathSync(dir), '.bashrc'));
      expect(safeOutputPath('..\\..\\evil.exe')).toBe(path.join(fs.realpathSync(dir), 'evil.exe'));
      expect(safeOutputPath('a\u0007b\n.txt')).toBe(path.join(fs.realpathSync(dir), 'ab.txt'));
      expect(safeOutputPath('..', { fallback: 'attachment' })).toBe(path.join(fs.realpathSync(dir), 'attachment'));
    } finally { process.chdir(cwd); }
  });
  it('refuses to overwrite without force', () => {
    const f = path.join(dir, 'exists.txt');
    fs.writeFileSync(f, 'x');
    expect(() => safeOutputPath('ignored', { out: f })).toThrow(CliError);
    expect(safeOutputPath('ignored', { out: f, force: true })).toBe(f);
  });
});

describe('time zones', () => {
  it('resolveTimeZone precedence and validation', () => {
    expect(resolveTimeZone('Europe/Berlin')).toBe('Europe/Berlin');
    const prev = process.env.MS365_CLI_TIMEZONE;
    process.env.MS365_CLI_TIMEZONE = 'Asia/Tokyo';
    try { expect(resolveTimeZone()).toBe('Asia/Tokyo'); } finally {
      if (prev === undefined) delete process.env.MS365_CLI_TIMEZONE; else process.env.MS365_CLI_TIMEZONE = prev;
    }
    try { resolveTimeZone('Not/AZone'); expect.fail(); } catch (e) { expect((e as CliError).exitCode).toBe(ExitCode.USAGE); }
  });
  it('preferTimeZone header', () => {
    expect(preferTimeZone(LA)).toEqual({ Prefer: 'outlook.timezone="America/Los_Angeles"' });
  });
  it('parseGraphDateTime treats offset-less UTC values as UTC, not local (M-7)', () => {
    const d = parseGraphDateTime({ dateTime: '2026-10-14T16:00:00.0000000', timeZone: 'UTC' });
    expect(d.toISOString()).toBe('2026-10-14T16:00:00.000Z');
  });
  it('parseGraphDateTime honors an IANA zone (DST aware)', () => {
    expect(parseGraphDateTime({ dateTime: '2026-10-14T09:00:00', timeZone: LA }).toISOString()).toBe('2026-10-14T16:00:00.000Z');
    expect(parseGraphDateTime({ dateTime: '2026-12-14T09:00:00', timeZone: LA }).toISOString()).toBe('2026-12-14T17:00:00.000Z');
  });
  it('dayRange is the local day, not the UTC day (M-8)', () => {
    // 6 PM PDT on Oct 14 is already Oct 15 in UTC
    const sixPm = new Date('2026-10-15T01:00:00Z');
    expect(dayRange(sixPm, LA)).toEqual({ start: '2026-10-14T07:00:00.000Z', end: '2026-10-15T07:00:00.000Z' });
  });
  it('toGraphDateTime / toIsoWithOffset', () => {
    const d = new Date('2026-10-14T16:00:00Z');
    expect(toGraphDateTime(d, LA)).toBe('2026-10-14T09:00:00');
    expect(toIsoWithOffset(d, LA)).toBe('2026-10-14T09:00:00-07:00');
    expect(toIsoWithOffset(d, 'Asia/Kolkata')).toBe('2026-10-14T21:30:00+05:30');
  });
  it('wallTimeToDate across the DST fall-back day', () => {
    expect(wallTimeToDate({ year: 2026, month: 11, day: 1, hour: 12, minute: 0, second: 0 }, LA).toISOString()).toBe('2026-11-01T20:00:00.000Z');
  });
});

describe('parseTime (D22)', () => {
  const now = new Date('2026-10-14T17:30:00Z'); // Wed 10:30 PDT
  const p = (s: string) => parseTime(s, { tz: LA, now });
  it.each([
    ['2026-10-20T15:00:00Z', '2026-10-20T08:00:00', false],
    ['2026-10-20T15:00', '2026-10-20T15:00:00', false],
    ['2026-10-20', '2026-10-20T00:00:00', true],
    ['now', '2026-10-14T10:30:00', false],
    ['today 15:00', '2026-10-14T15:00:00', false],
    ['tomorrow', '2026-10-15T00:00:00', true],
    ['tomorrow at 3pm', '2026-10-15T15:00:00', false],
    ['yesterday 9:15', '2026-10-13T09:15:00', false],
    ['fri 9:30', '2026-10-16T09:30:00', false],
    ['wed', '2026-10-14T00:00:00', true],
    ['next wed', '2026-10-21T00:00:00', true],
    ['next monday 9am', '2026-10-19T09:00:00', false],
    ['16:45', '2026-10-14T16:45:00', false],
    ['+2h', '2026-10-14T12:30:00', false],
    ['+30m', '2026-10-14T11:00:00', false],
    ['+3d', '2026-10-17T10:30:00', false],
    ['+1w', '2026-10-21T10:30:00', false],
    ['-1d', '2026-10-13T10:30:00', false],
  ])('%s', (input, wall, dateOnly) => {
    const r = p(input);
    expect(r.dateTime).toBe(wall);
    expect(r.timeZone).toBe(LA);
    expect(r.dateOnly).toBe(dateOnly);
  });
  it('+Nd keeps wall time across DST (Nov 1)', () => {
    expect(parseTime('+20d', { tz: LA, now }).dateTime).toBe('2026-11-03T10:30:00');
  });
  it.each(['soon', '25:00', 'tomorrow 9:99', '2026-13-45'])('rejects %s with exit 2', (s) => {
    try { p(s); expect.fail('should throw'); } catch (e) { expect((e as CliError).exitCode).toBe(ExitCode.USAGE); }
  });
});
