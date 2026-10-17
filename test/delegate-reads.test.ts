import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { usedScopes } from '../src/catalog.js';
import { assertWithinCeiling, ceilingSurface, profileCeiling } from '../src/ceiling.js';
import { CliError, ExitCode, setErrorFormat } from '../src/errors.js';
import { registerMailCommands } from '../src/commands/mail.js';
import { registerCalendarCommands } from '../src/commands/calendar.js';
import { registerWatchCommands } from '../src/commands/watch.js';
import { runCommand } from './helpers/cli.js';
import { fakeJwt } from './helpers/entry.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const TENANT = '11111111-2222-3333-4444-555555555555';
const jwt = fakeJwt({ aud: 'https://graph.microsoft.com', tid: TENANT, oid: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee', upn: 'ea@contoso.com', exp: Math.floor(Date.now() / 1000) + 3600 });
const BOB = 'bob@contoso.com';
/** Recorded paths are decoded; the request URL itself carries `bob%40contoso.com`. */
const BOB_PATH = '/users/bob@contoso.com';
const NOW = '2026-10-14T17:00:00.000Z';
const errorCode = (stderr: string) => JSON.parse(stderr.trim().split('\n').at(-1)!).error.code;
const decode = (b: string) => JSON.parse(Buffer.from(b, 'base64url').toString('utf8'));

function refusal(fn: () => void): CliError {
  try { fn(); } catch (e) { return e as CliError; }
  throw new Error('expected a refusal');
}

describe('--mailbox delegate reads (AH-10, AC 24)', () => {
  it('calendar view --mailbox calls /users/{upn}/calendarView and sets _ref.mailbox to the target', async () => {
    const r = await runCommand(registerCalendarCommands, ['calendar', 'view', '--mailbox', BOB, '--start', '2026-10-20', '--end', '2026-10-21', '-o', 'json'], {
      refs: true, token: jwt,
      routes: [['GET', `${BOB_PATH}/calendarView`, () => ({ value: [{ id: 'E1', changeKey: 'v1' }] })]],
    });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(r.calls.map(c => c.path.split('?')[0])).toEqual([`${BOB_PATH}/calendarView`]);
    expect(r.calls[0].url).toContain('/users/bob%40contoso.com/calendarView?');
    expect(JSON.parse(r.stdout).value[0]._ref).toMatchObject({ mailbox: BOB, resource: 'event', id: 'E1' });
  });

  it('mail list/get, calendar get/instances read the named mailbox; without --mailbox they stay on /me', async () => {
    const cases: Array<[typeof registerMailCommands, string[], string]> = [
      [registerMailCommands, ['mail', 'list', '--mailbox', BOB], `${BOB_PATH}/messages`],
      [registerMailCommands, ['mail', 'get', 'M1', '--mailbox', BOB], `${BOB_PATH}/messages/M1`],
      [registerCalendarCommands, ['calendar', 'get', 'E1', '--mailbox', BOB], `${BOB_PATH}/events/E1`],
      [registerCalendarCommands, ['calendar', 'instances', 'S1', '--mailbox', BOB, '--start', '2026-10-20', '--end', '2026-10-21'], `${BOB_PATH}/events/S1/instances`],
      [registerMailCommands, ['mail', 'list'], '/me/messages'],
    ];
    for (const [register, argv, expected] of cases) {
      const r = await runCommand(register, [...argv, '-o', 'json'], { routes: [['GET', expected, () => ({ value: [], id: 'X' })]] });
      expect(r.exitCode, `${argv.join(' ')}: ${r.stderr}`).toBe(0);
      expect(r.calls.map(c => c.path.split('?')[0]), argv.join(' ')).toEqual([expected]);
    }
  });

  it('a blank or path-like --mailbox is a usage error before any request', async () => {
    for (const bad of [' ', 'a/b']) {
      const r = await runCommand(registerMailCommands, ['mail', 'list', '--mailbox', bad, '-o', 'json']);
      expect(r.exitCode).toBe(ExitCode.USAGE);
      expect(r.calls).toEqual([]);
    }
  });

  it('--mailbox adds the .Shared clause: refused under a ceiling without it (exit 4), allowed with it', () => {
    const own = profileCeiling({ type: 'host', tenantId: TENANT, scopes: ['User.Read', 'Mail.Read', 'Calendars.Read'] })!;
    const delegate = profileCeiling({ type: 'host', tenantId: TENANT, scopes: ['Mail.Read', 'Calendars.Read', 'Calendars.Read.Shared'] })!;
    expect(usedScopes('calendar view', { mailbox: BOB })).toEqual([
      { scopes: ['Calendars.Read'] }, { scopes: ['Calendars.Read.Shared'], for: 'with --mailbox' },
    ]);
    const e = refusal(() => assertWithinCeiling('calendar view', usedScopes('calendar view', { mailbox: BOB }), own, 'chief'));
    expect(e).toMatchObject({ code: 'scope_ceiling', exitCode: ExitCode.PERMISSION });
    expect(e.message).toContain('Calendars.Read.Shared');
    expect(() => assertWithinCeiling('calendar view', usedScopes('calendar view', {}), own, 'chief')).not.toThrow();
    expect(() => assertWithinCeiling('calendar view', usedScopes('calendar view', { mailbox: BOB }), delegate, 'chief')).not.toThrow();
    expect(refusal(() => assertWithinCeiling('watch', usedScopes('watch', { mailbox: BOB }, ['mail']), delegate, 'chief')).message).toContain('Mail.Read.Shared');
    expect(() => assertWithinCeiling('watch', usedScopes('watch', { mailbox: BOB }, ['events']), delegate, 'chief')).not.toThrow();
    // The command stays listed (MCP, describe) when the caller's own read fits.
    expect(ceilingSurface('mail list', own)).toEqual({ allowed: true });
  });
});

describe('watch --mailbox (AH-10, AC 24)', () => {
  let dir: string;
  let env: Record<string, string>;
  const EVENTS = `${BOB_PATH}/calendarView/delta`;
  const baseline = [['GET', EVENTS, () => ({ value: [], '@odata.deltaLink': `${GRAPH}${EVENTS}?$deltatoken=D1` })]] as never;

  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-delegate-'));
    env = { MS365_CLI_CONFIG_DIR: dir, MS365_CLI_TIMEZONE: 'UTC' };
    vi.useFakeTimers({ toFake: ['Date'] });
    vi.setSystemTime(new Date(NOW));
  });
  afterEach(() => {
    vi.useRealTimers();
    setErrorFormat('text');
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it('the checkpoint blob records the mailbox; resuming it for another mailbox is checkpoint_mismatch before any request', async () => {
    const first = await runCommand(registerWatchCommands, ['watch', 'events', '--mailbox', BOB, '--checkpoint', 'new'], { env, routes: baseline });
    expect(first.exitCode, first.stderr).toBe(0);
    const line = JSON.parse(first.stdout.trim());
    expect(decode(line.checkpoint)).toMatchObject({ resource: 'events', endpoint: '/users/bob%40contoso.com/calendarView/delta', mailbox: BOB });

    setErrorFormat('json');
    for (const argv of [['watch', 'events'], ['watch', 'events', '--mailbox', 'carol@contoso.com']]) {
      const r = await runCommand(registerWatchCommands, [...argv, '--checkpoint', line.checkpoint], { env, routes: baseline });
      expect(r.exitCode, argv.join(' ')).toBe(ExitCode.USAGE);
      expect(errorCode(r.stderr)).toBe('checkpoint_mismatch');
      expect(r.calls).toEqual([]);
    }
  });

  it('a named-cursor run keeps its own default cursor per mailbox; --mailbox is refused for files and contacts', async () => {
    const r = await runCommand(registerWatchCommands, ['watch', 'events', '--mailbox', BOB], { env, routes: baseline });
    expect(r.exitCode, r.stderr).toBe(0);
    expect(fs.readdirSync(path.join(dir, 'watch'))).toEqual([`events@${BOB}.json`]);
    const files = await runCommand(registerWatchCommands, ['watch', 'files', '--mailbox', BOB], { env });
    expect(files.exitCode).toBe(ExitCode.USAGE);
    expect(files.calls).toEqual([]);
  });
});
