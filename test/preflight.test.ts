import { describe, it, expect } from 'vitest';
import { scopeSatisfied, assertConsented } from '../src/preflight.js';
import { CliError, ExitCode } from '../src/errors.js';

describe('pre-flight scope check (T39)', () => {
  const g = (...s: string[]) => new Set(s);
  it.each([
    ['Mail.Read', ['Mail.Read'], true],
    ['Mail.Read', ['Mail.ReadWrite'], true],
    ['Mail.Read', ['Mail.Read.All'], true],
    ['Mail.ReadWrite', ['Mail.Read'], false],
    ['Mail.ReadWrite', ['Mail.ReadWrite.All'], true],
    ['Calendars.Read.Shared', ['Calendars.ReadWrite.Shared'], true],
    ['Calendars.Read.Shared', ['Calendars.Read'], false],
    ['Mail.Send', ['Mail.Send'], true],
    ['Mail.Send', ['Mail.ReadWrite'], false],
    ['Sites.Read.All', ['Sites.ReadWrite.All'], true],
    ['ChannelMessage.Send', ['ChannelMessage.Send'], true],
  ])('%s satisfied by %j → %s', (req, granted, want) => {
    expect(scopeSatisfied(req, g(...granted))).toBe(want);
  });

  it('passes when any listed scope is satisfied; fails with exit 4 otherwise', () => {
    expect(() => assertConsented('mail list', ['Mail.Read'], ['User.Read', 'Mail.ReadWrite'])).not.toThrow();
    expect(() => assertConsented('x', [], [])).not.toThrow();
    try { assertConsented('meetings transcript', ['OnlineMeetingTranscript.Read.All'], ['User.Read']); expect.fail(); } catch (e) {
      expect((e as CliError).exitCode).toBe(ExitCode.PERMISSION);
      expect((e as CliError).code).toBe('missing_scope');
    }
  });
});
