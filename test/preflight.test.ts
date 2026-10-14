import { describe, it, expect } from 'vitest';
import { scopeSatisfied, assertConsented } from '../src/preflight.js';
import { requiredScopes } from '../src/catalog.js';
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
    expect(() => assertConsented('mail list', [{ scopes: ['Mail.Read'] }], ['User.Read', 'Mail.ReadWrite'])).not.toThrow();
    expect(() => assertConsented('x', [], [])).not.toThrow();
    try { assertConsented('meetings transcript', [{ scopes: ['OnlineMeetingTranscript.Read.All'] }], ['User.Read']); expect.fail(); } catch (e) {
      expect((e as CliError).exitCode).toBe(ExitCode.PERMISSION);
      expect((e as CliError).code).toBe('missing_scope');
    }
  });
});

describe('per-option and per-argument scopes in CNF (Q-16 to Q-19, F14)', () => {
  const check = (path: string, granted: string[], opts: Record<string, unknown> = {}, args: unknown[] = []) => {
    try {
      assertConsented(path, requiredScopes(path, opts, args), granted);
      return 'ok';
    } catch (e) {
      expect((e as CliError).exitCode).toBe(ExitCode.PERMISSION);
      return (e as CliError).message;
    }
  };

  it('search query: every selected --entity must be satisfied, and the message names the unmet one', () => {
    expect(check('search query', ['Mail.Read'], { entity: 'message' })).toBe('ok');
    expect(check('search query', ['Mail.Read'], { entity: 'driveItem' })).toMatch(/--entity driveItem needs Files\.Read\.All or Sites\.Read\.All/);
    expect(check('search query', ['Mail.Read'], { entity: 'message,event' })).toMatch(/--entity event needs Calendars\.Read/);
    expect(check('search query', ['Mail.Read', 'Calendars.Read', 'Files.Read.All'], { entity: 'message,event,driveItem' })).toBe('ok');
    expect(check('search query', ['Mail.Read'], { entity: 'message,bogus' })).toBe('ok'); // unknown values: the command's own usage error
  });

  it('watch: the scope follows <resource>', () => {
    expect(check('watch', ['Mail.Read'], {}, ['mail'])).toBe('ok');
    expect(check('watch', ['Mail.Read'], {}, ['files'])).toMatch(/files needs Files\.Read/);
    expect(check('watch', ['Contacts.ReadWrite'], {}, ['contacts'])).toBe('ok');
  });

  it('resolve: the scope follows <type>; mail-folder with --user also needs Mail.Read.Shared', () => {
    expect(check('resolve', ['Tasks.Read'], {}, ['plan', 'X'])).toBe('ok');
    expect(check('resolve', ['Team.ReadBasic.All'], {}, ['plan', 'X'])).toMatch(/plan needs Tasks\.Read/);
    expect(check('resolve', ['Mail.Read'], {}, ['mail-folder', 'Inbox'])).toBe('ok');
    expect(check('resolve', ['Tasks.Read'], {}, ['mail-folder', 'Inbox'])).toMatch(/mail-folder needs Mail\.Read/);
    expect(check('resolve', ['Mail.Read'], { user: 'a@b.com' }, ['mail-folder', 'Inbox'])).toMatch(/--user needs Mail\.Read\.Shared/);
    expect(check('resolve', ['Mail.Read', 'Mail.Read.Shared'], { user: 'a@b.com' }, ['mail-folder', 'Inbox'])).toBe('ok');
  });

  it('user list accepts User.ReadBasic.All or User.Read.All', () => {
    expect(check('user list', ['User.ReadBasic.All'])).toBe('ok');
    expect(check('user list', ['User.Read.All'])).toBe('ok');
    expect(check('user list', ['User.Read'])).not.toBe('ok');
  });

  it('commands without a selector keep one any-of clause of their catalog scopes', () => {
    expect(requiredScopes('mail list', {}, [])).toEqual([{ scopes: ['Mail.Read'] }]);
    expect(requiredScopes('describe', {}, [])).toEqual([]);
    expect(check('teams send-channel-message', ['ChannelMessage.Send'])).toBe('ok');
  });
});
