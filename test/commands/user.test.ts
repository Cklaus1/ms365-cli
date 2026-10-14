import { describe, it, expect, afterEach, beforeEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { runCommand } from '../helpers/cli.js';
import { registerUserCommands } from '../../src/commands/user.js';
import { setErrorFormat } from '../../src/errors.js';
import { setQuiet } from '../../src/utils.js';

const run = (argv: string[], opts: Parameters<typeof runCommand>[2] = {}) =>
  runCommand(registerUserCommands, argv, opts);

afterEach(() => { process.exitCode = undefined; });

describe('user set-ooo (M-10)', () => {
  it('rejects --status scheduled without --start/--end (exit 2, no request)', async () => {
    for (const extra of [[], ['--start', '2026-10-20T09:00'], ['--end', '2026-10-21T09:00']]) {
      const r = await run(['user', 'set-ooo', '--status', 'scheduled', ...extra]);
      expect(r.exitCode).toBe(2);
      expect(r.calls).toHaveLength(0);
    }
  });

  it('sends scheduled times as wall time in the resolved zone', async () => {
    const r = await run(['user', 'set-ooo', '--status', 'scheduled', '--start', '2026-10-20T09:00', '--end', '2026-10-24T17:30',
      '--timezone', 'America/Los_Angeles']);
    expect(r.exitCode).toBe(0);
    const s = (r.calls[0].body as { automaticRepliesSetting: Record<string, unknown> }).automaticRepliesSetting;
    expect(s.scheduledStartDateTime).toEqual({ dateTime: '2026-10-20T09:00:00', timeZone: 'America/Los_Angeles' });
    expect(s.scheduledEndDateTime).toEqual({ dateTime: '2026-10-24T17:30:00', timeZone: 'America/Los_Angeles' });
  });

  it('uses MS365_CLI_TIMEZONE when --timezone is absent', async () => {
    const r = await run(['user', 'set-ooo', '--status', 'scheduled', '--start', '2026-10-20T09:00', '--end', '2026-10-21T09:00'],
      { env: { MS365_CLI_TIMEZONE: 'Europe/Berlin' } });
    const s = (r.calls[0].body as { automaticRepliesSetting: Record<string, { timeZone: string }> }).automaticRepliesSetting;
    expect(s.scheduledStartDateTime.timeZone).toBe('Europe/Berlin');
  });

  it('rejects an unparseable time with exit 2', async () => {
    const r = await run(['user', 'set-ooo', '--status', 'scheduled', '--start', 'whenever', '--end', '2026-10-21T09:00']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });
});

describe('user download-photo (M-28)', () => {
  let dir: string;
  let cwd: string;
  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-photo-'));
    cwd = process.cwd();
    process.chdir(dir);
  });
  afterEach(() => {
    process.chdir(cwd);
    fs.rmSync(dir, { recursive: true, force: true });
  });

  const photo = () => new Response(new Uint8Array([0xff, 0xd8, 0xff, 0x00]), { status: 200 });

  it('writes the binary photo to ./profile-photo.jpg by default', async () => {
    const r = await run(['user', 'download-photo'], { routes: [['GET', '/me/photo/$value', photo]] });
    expect(r.exitCode).toBe(0);
    expect([...fs.readFileSync(path.join(dir, 'profile-photo.jpg'))]).toEqual([0xff, 0xd8, 0xff, 0x00]);
    expect(r.stdout).toBe('');
    expect(r.stderr).toContain('Photo saved to');
  });

  it('refuses to overwrite an existing --out without --force', async () => {
    fs.writeFileSync('me.jpg', 'old');
    const r = await run(['user', 'download-photo', '--out', 'me.jpg'], { routes: [['GET', '/me/photo/$value', photo]] });
    expect(r.exitCode).not.toBe(0);
    expect(r.stderr).toMatch(/Refusing to overwrite/);
    expect(fs.readFileSync('me.jpg', 'utf8')).toBe('old');
    expect(r.calls).toHaveLength(0);
  });

  it('overwrites with --force', async () => {
    fs.writeFileSync('me.jpg', 'old');
    const r = await run(['user', 'download-photo', '--out', 'me.jpg', '--force'], { routes: [['GET', '/me/photo/$value', photo]] });
    expect(r.exitCode).toBe(0);
    expect(fs.readFileSync('me.jpg').length).toBe(4);
  });

  it('writes nothing in dry-run', async () => {
    const r = await run(['user', 'download-photo'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(fs.existsSync('profile-photo.jpg')).toBe(false);
  });

  it('reports a 404 as a not_found CliError in --error-format json (V-23)', async () => {
    setErrorFormat('json');
    try {
      const r = await run(['user', 'download-photo'], {
        routes: [['GET', '/me/photo/$value', () => new Response(JSON.stringify({ error: { code: 'ImageNotFound', message: 'not found' } }),
          { status: 404, headers: { 'content-type': 'application/json' } })]],
      });
      expect(r.exitCode).toBe(5);
      expect(r.stderr).toMatch(/"code":\s*"not_found"/);
      expect(process.exitCode).toBeUndefined();
    } finally {
      setErrorFormat('text');
    }
  });

  it('reports a 404 as "No profile photo found" with exit 5', async () => {
    const r = await run(['user', 'download-photo'], {
      routes: [['GET', '/me/photo/$value', () => new Response(JSON.stringify({ error: { code: 'ImageNotFound', message: 'not found' } }),
        { status: 404, headers: { 'content-type': 'application/json' } })]],
    });
    expect(r.exitCode).toBe(5);
    expect(r.stderr).toContain('No profile photo found');
    expect(fs.existsSync('profile-photo.jpg')).toBe(false);
  });
});

describe('user list --all (M-26)', () => {
  it('honours --page-limit', async () => {
    const r = await run(['user', 'list', '--all', '--page-limit', '1'], {
      routes: [['GET', '/users', () => ({ value: [{ id: '1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=x' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls).toHaveLength(1);
  });
});

describe('user set-ooo choices (V-4)', () => {
  it('rejects an unknown --status or --audience with exit 2 and no request', async () => {
    for (const argv of [['--status', 'bogus'], ['--status', 'disabled', '--audience', 'nobody']]) {
      const r = await run(['user', 'set-ooo', ...argv]);
      expect(r.exitCode).toBe(2);
      expect(r.calls).toHaveLength(0);
    }
  });

  it('accepts the documented values', async () => {
    const r = await run(['user', 'set-ooo', '--status', 'disabled', '--audience', 'contactsOnly']);
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toEqual({ automaticRepliesSetting: { status: 'disabled', externalAudience: 'contactsOnly' } });
  });
});

describe('user set-timezone / set-language (V-5, F24)', () => {
  it.each(['America/New_York', 'UTC', 'Pacific Standard Time', 'W. Europe Standard Time', 'US Mountain Standard Time'])(
    'accepts %s', async (tz) => {
      const r = await run(['user', 'set-timezone', tz]);
      expect(r.exitCode).toBe(0);
      expect(r.calls[0].body).toEqual({ timeZone: tz });
    });

  it.each(['Not/AZone', 'pacific standard time', 'Mars Time Zone'])('rejects %s with exit 2 and no request', async (tz) => {
    const r = await run(['user', 'set-timezone', tz]);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('rejects a malformed language tag with exit 2', async () => {
    const r = await run(['user', 'set-language', 'xx_bad']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('sends a valid language tag', async () => {
    const r = await run(['user', 'set-language', 'fr-FR']);
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toEqual({ language: { locale: 'fr-FR' } });
  });
});

describe('user status lines (DR-2, V-1)', () => {
  afterEach(() => setQuiet(false));
  const cases: string[][] = [
    ['user', 'set-timezone', 'UTC'],
    ['user', 'set-language', 'en-US'],
    ['user', 'set-presence', '--availability', 'Busy', '--activity', 'InACall'],
  ];

  it.each(cases)('%s %s prints its confirmation on stderr, not stdout', async (...argv) => {
    const r = await run(argv, { clientId: 'cid' });
    expect(r.exitCode).toBe(0);
    expect(r.stdout).toBe('');
    expect(r.stderr).not.toBe('');
  });

  it.each(cases)('%s %s prints nothing under -q', async (...argv) => {
    setQuiet(true);
    const r = await run(argv, { clientId: 'cid' });
    expect(r.exitCode).toBe(0);
    expect(r.stdout).toBe('');
    expect(r.stderr).toBe('');
  });

  it.each(cases)('%s %s claims nothing under dry-run', async (...argv) => {
    const r = await run(argv, { clientId: 'cid', dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stdout + r.stderr).not.toMatch(/ set to /);
  });
});

describe('user presence (V-2, V-6, DR-15, F25)', () => {
  const set = ['user', 'set-presence', '--availability', 'Busy', '--activity', 'InACall'];

  it('uses the client ID resolved by GraphClient as sessionId', async () => {
    const r = await run(set, { clientId: 'profile-cid', env: { MS365_CLI_CLIENT_ID: 'env-cid' } });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toEqual({ sessionId: 'profile-cid', availability: 'Busy', activity: 'InACall', expirationDuration: 'PT60M' });
  });

  it('does not read MS365_CLI_CLIENT_ID itself', async () => {
    const r = await run(set, { env: { MS365_CLI_CLIENT_ID: 'env-cid' } });
    expect(r.exitCode).toBe(3);
    expect(r.calls).toHaveLength(0);
  });

  it.each([set, ['user', 'clear-presence']])('live with no client ID exits 3 with code no_client_id (%s %s)', async (...argv) => {
    setErrorFormat('json');
    try {
      const r = await run(argv);
      expect(r.exitCode).toBe(3);
      expect(r.stderr).toMatch(/"code":\s*"no_client_id"/);
      expect(r.calls).toHaveLength(0);
    } finally {
      setErrorFormat('text');
    }
  });

  it('clear-presence sends the resolved client ID', async () => {
    const r = await run(['user', 'clear-presence'], { clientId: 'cid' });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].body).toEqual({ sessionId: 'cid' });
  });

  it.each([set, ['user', 'clear-presence']])('dry-run with no client ID previews sessionId <client-id> (%s %s)', async (...argv) => {
    const r = await run(argv, { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain('"sessionId":"<client-id>"');
  });

  it.each([
    ['--availability', 'Offline', '--activity', 'Away'],
    ['--availability', 'Busy', '--activity', 'OffWork'],
    ['--availability', 'Busy', '--activity', 'InACall', '--duration', '30'],
    ['--availability', 'Busy', '--activity', 'InACall', '--duration', 'PT4M'],
    ['--availability', 'Busy', '--activity', 'InACall', '--duration', 'PT4H1M'],
    ['--availability', 'Busy', '--activity', 'InACall', '--duration', 'PT'],
  ])('rejects set-presence %s %s %s %s with exit 2 and no request', async (...argv) => {
    const r = await run(['user', 'set-presence', ...argv], { clientId: 'cid' });
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it.each(['PT5M', 'PT4H', 'PT1H30M', 'PT90M'])('accepts --duration %s', async (d) => {
    const r = await run([...set, '--duration', d], { clientId: 'cid' });
    expect(r.exitCode).toBe(0);
    expect((r.calls[0].body as { expirationDuration: string }).expirationDuration).toBe(d);
  });
});

describe('user upload-photo (V-7)', () => {
  let dir: string;
  beforeEach(() => { dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-upphoto-')); });
  afterEach(() => { fs.rmSync(dir, { recursive: true, force: true }); });

  it('a missing file exits 2 with no request', async () => {
    const r = await run(['user', 'upload-photo', path.join(dir, 'missing.jpg')]);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('an unsupported extension exits 2 with no request', async () => {
    const f = path.join(dir, 'me.webp');
    fs.writeFileSync(f, 'x');
    const r = await run(['user', 'upload-photo', f]);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('an oversized file exits 2 with no request', async () => {
    const f = path.join(dir, 'big.png');
    fs.writeFileSync(f, Buffer.alloc(4 * 1024 * 1024 + 1));
    const r = await run(['user', 'upload-photo', f]);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('uploads a PNG with its content type', async () => {
    const f = path.join(dir, 'me.png');
    fs.writeFileSync(f, 'png');
    const r = await run(['user', 'upload-photo', f]);
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].method).toBe('PUT');
    expect(r.calls[0].headers['content-type'] ?? r.calls[0].headers['Content-Type']).toBe('image/png');
    expect(r.stderr).toContain('Profile photo updated.');
  });
});
