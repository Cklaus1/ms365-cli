import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { runCli, tempDir } from '../e2e/run-cli.js';

const ENV = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_TOKEN: 'fake' };
const GRAPH = 'https://graph.microsoft.com/v1.0';
const PNG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 1, 2, 3]);

function patchBody(r: { requests: Array<{ method: string; body: string | null }> }): Record<string, unknown> {
  const patch = r.requests.find((q) => q.method === 'PATCH');
  expect(patch).toBeDefined();
  return JSON.parse(patch!.body!).automaticRepliesSetting;
}

describe('USER-1 set-ooo', () => {
  const routes = [{ match: '/me/mailboxSettings$', method: 'PATCH', body: { automaticRepliesSetting: { status: 'alwaysEnabled' } } }];

  it('without --audience sends no externalAudience', async () => {
    const r = await runCli(['user', 'set-ooo', '--status', 'alwaysEnabled', '--internal', 'away', '-o', 'json'], { env: ENV, routes });
    expect(r.code).toBe(0);
    expect(r.requests[0].url).toBe(`${GRAPH}/me/mailboxSettings`);
    const settings = patchBody(r);
    expect(settings).toEqual({ status: 'alwaysEnabled', internalReplyMessage: 'away' });
    expect(JSON.parse(r.stdout)).toEqual({ automaticRepliesSetting: { status: 'alwaysEnabled' } });
  });

  it('sends externalAudience only when given', async () => {
    const r = await runCli(['user', 'set-ooo', '--status', 'disabled', '--audience', 'contactsOnly', '-o', 'json'], { env: ENV, routes });
    expect(r.code).toBe(0);
    expect(patchBody(r)).toEqual({ status: 'disabled', externalAudience: 'contactsOnly' });
  });

  it('rejects an unknown --audience before any request', async () => {
    const r = await runCli(['user', 'set-ooo', '--status', 'disabled', '--audience', 'everyone'], { env: ENV, routes });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });

  it('--status scheduled without dates exits 1 with no request', async () => {
    const r = await runCli(['user', 'set-ooo', '--status', 'scheduled', '--internal', 'away'], { env: ENV, routes });
    expect(r.code).toBe(1);
    expect(r.stderr).toMatch(/--start and --end/);
    expect(r.requests).toHaveLength(0);
  });

  it('--status scheduled with only --start exits 1 with no request', async () => {
    const r = await runCli(['user', 'set-ooo', '--status', 'scheduled', '--start', '2026-10-07T09:00:00Z'], { env: ENV, routes });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });

  it.each([
    ['--start', '2026-02-30', '2026-03-05'],
    ['--start', '2026-02-30T09:00', '2026-03-05'],
    ['--start', '2026-02-29T09:00:00Z', '2026-03-05'],
    ['--end', '2026-03-01', '2026-04-31T09:00:00+02:00'],
  ])('rejects an impossible calendar date in %s (%s / %s) with no request', async (flag, start, end) => {
    const r = await runCli(['user', 'set-ooo', '--status', 'scheduled', '--start', start, '--end', end], { env: { ...ENV, TZ: 'Europe/Berlin' }, routes });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toContain(`Invalid ${flag} "${flag === '--start' ? start : end}". Use ISO 8601`);
  });

  it('accepts Feb 29 in a leap year', async () => {
    const r = await runCli(['user', 'set-ooo', '--status', 'scheduled', '--start', '2028-02-29T00:00:00Z', '--end', '2028-03-01T00:00:00Z', '-o', 'json'], { env: ENV, routes });
    expect(r.code).toBe(0);
    expect(patchBody(r).scheduledStartDateTime).toEqual({ dateTime: '2028-02-29T00:00:00', timeZone: 'UTC' });
  });

  it('converts scheduled dates with an offset to UTC', async () => {
    const r = await runCli([
      'user', 'set-ooo', '--status', 'scheduled',
      '--start', '2026-10-07T09:00:00+02:00', '--end', '2026-10-09T17:30:00+02:00', '-o', 'json',
    ], { env: ENV, routes });
    expect(r.code).toBe(0);
    const settings = patchBody(r);
    expect(settings.scheduledStartDateTime).toEqual({ dateTime: '2026-10-07T07:00:00', timeZone: 'UTC' });
    expect(settings.scheduledEndDateTime).toEqual({ dateTime: '2026-10-09T15:30:00', timeZone: 'UTC' });
  });

  it('reads offset-less dates as local wall time (TZ), not UTC', async () => {
    const r = await runCli([
      'user', 'set-ooo', '--status', 'scheduled',
      '--start', '2026-10-07T09:00', '--end', '2026-10-08', '-o', 'json',
    ], { env: { ...ENV, TZ: 'Europe/Berlin' }, routes });
    expect(r.code).toBe(0);
    const settings = patchBody(r);
    expect(settings.scheduledStartDateTime).toEqual({ dateTime: '2026-10-07T07:00:00', timeZone: 'UTC' });
    expect(settings.scheduledEndDateTime).toEqual({ dateTime: '2026-10-07T22:00:00', timeZone: 'UTC' });
  });

  it('rejects an end before the start', async () => {
    const r = await runCli([
      'user', 'set-ooo', '--status', 'scheduled', '--start', '2026-10-09T00:00:00Z', '--end', '2026-10-07T00:00:00Z',
    ], { env: ENV, routes });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });
});

describe('USER-2 presence', () => {
  it('set-presence POSTs setUserPreferredPresence (no client ID needed)', async () => {
    const r = await runCli(['user', 'set-presence', '--availability', 'DoNotDisturb', '--activity', 'DoNotDisturb', '--duration', 'PT8H'], {
      env: { MS365_CLI_TOKEN: 'fake' },
      routes: [{ match: '/me/presence/setUserPreferredPresence$', method: 'POST', status: 200, body: '' }],
    });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(1);
    expect(r.requests[0].method).toBe('POST');
    expect(r.requests[0].url).toBe(`${GRAPH}/me/presence/setUserPreferredPresence`);
    expect(JSON.parse(r.requests[0].body!)).toEqual({ availability: 'DoNotDisturb', activity: 'DoNotDisturb', expirationDuration: 'PT8H' });
    expect(r.stdout).toBe('');
  });

  it('derives the activity from the availability (Offline -> OffWork)', async () => {
    const r = await runCli(['user', 'set-presence', '--availability', 'Offline'], {
      env: ENV,
      routes: [{ match: '/setUserPreferredPresence$', method: 'POST', status: 200, body: '' }],
    });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.requests[0].body!)).toMatchObject({ availability: 'Offline', activity: 'OffWork' });
  });

  it('rejects an unsupported availability/activity pair with no request', async () => {
    const r = await runCli(['user', 'set-presence', '--availability', 'Busy', '--activity', 'Away'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.stderr).toMatch(/Busy.*Busy/);
    expect(r.requests).toHaveLength(0);
  });

  it('rejects an unknown availability with no request', async () => {
    const r = await runCli(['user', 'set-presence', '--availability', 'InAMeeting'], { env: ENV, routes: [] });
    expect(r.code).toBe(1);
    expect(r.requests).toHaveLength(0);
  });

  it('clear-presence POSTs clearUserPreferredPresence', async () => {
    const r = await runCli(['user', 'clear-presence'], {
      env: { MS365_CLI_TOKEN: 'fake' },
      routes: [{ match: '/me/presence/clearUserPreferredPresence$', method: 'POST', status: 200, body: '' }],
    });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(1);
    expect(r.requests[0].url).toBe(`${GRAPH}/me/presence/clearUserPreferredPresence`);
  });
});

describe('USER-3 profile photo', () => {
  it('download-photo names the file after the content type', async () => {
    const cwd = tempDir();
    const r = await runCli(['user', 'download-photo'], {
      env: ENV, cwd,
      routes: [{ match: '/me/photo/\\$value$', headers: { 'content-type': 'image/png' }, bodyBase64: PNG.toString('base64') }],
    });
    expect(r.code).toBe(0);
    expect(fs.readFileSync(path.join(cwd, 'profile-photo.png'))).toEqual(PNG);
    expect(fs.existsSync(path.join(cwd, 'profile-photo.jpg'))).toBe(false);
    expect(r.stdout).toBe('');
  });

  it('download-photo refuses to overwrite without --force', async () => {
    const cwd = tempDir();
    const out = path.join(cwd, 'me.jpg');
    fs.writeFileSync(out, 'old');
    const routes = [{ match: '/me/photo/\\$value$', headers: { 'content-type': 'image/jpeg' }, bodyBase64: PNG.toString('base64') }];
    const r = await runCli(['user', 'download-photo', '--out', out], { env: ENV, cwd, routes });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('--force');
    expect(fs.readFileSync(out, 'utf8')).toBe('old');

    const forced = await runCli(['user', 'download-photo', '--out', out, '--force'], { env: ENV, cwd, routes });
    expect(forced.code).toBe(0);
    expect(fs.readFileSync(out)).toEqual(PNG);
  });

  it('download-photo without a photo exits 1 with a clear message', async () => {
    const r = await runCli(['user', 'download-photo'], {
      env: ENV,
      routes: [{ match: '/me/photo/\\$value$', status: 404, body: { error: { code: 'ImageNotFound', message: 'The photo wasn\'t found.' } } }],
    });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('No profile photo');
  });

  it('upload-photo in --dry-run prints no success message', async () => {
    const cwd = tempDir();
    fs.writeFileSync(path.join(cwd, 'p.png'), PNG);
    const r = await runCli(['--dry-run', 'user', 'upload-photo', 'p.png'], { env: ENV, cwd, routes: [] });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(0);
    expect(r.stderr).toContain(`[dry-run] PUT ${GRAPH}/me/photo/$value`);
    expect(r.stderr).not.toContain('updated');
  });

  it('upload-photo PUTs the bytes and reports success', async () => {
    const cwd = tempDir();
    fs.writeFileSync(path.join(cwd, 'p.png'), PNG);
    const r = await runCli(['user', 'upload-photo', 'p.png'], {
      env: ENV, cwd,
      routes: [{ match: '/me/photo/\\$value$', method: 'PUT', status: 200, body: '' }],
    });
    expect(r.code).toBe(0);
    expect(r.requests[0].headers['content-type']).toBe('image/png');
    expect(r.stderr).toContain('Profile photo updated.');
  });

  for (const name of ['doc.pdf', 'logo.svg']) {
    it(`USER-4 upload-photo rejects ${name} before any request`, async () => {
      const cwd = tempDir();
      fs.writeFileSync(path.join(cwd, name), PNG);
      const r = await runCli(['user', 'upload-photo', name], { env: ENV, cwd, routes: [] });
      expect(r.code).toBe(1);
      expect(r.requests).toHaveLength(0);
      expect(r.stderr).toMatch(/JPEG, PNG or GIF/);
    });
  }

  it('USER-4 upload-photo detects an upper-case .PNG as image/png', async () => {
    const cwd = tempDir();
    fs.writeFileSync(path.join(cwd, 'P.PNG'), PNG);
    const r = await runCli(['user', 'upload-photo', 'P.PNG'], {
      env: ENV, cwd,
      routes: [{ match: '/me/photo/\\$value$', method: 'PUT', status: 200, body: '' }],
    });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(1);
    expect(r.requests[0].method).toBe('PUT');
    expect(r.requests[0].headers['content-type']).toBe('image/png');
  });
});

describe('CORE-4 user status messages stay off stdout', () => {
  it('set-timezone writes nothing to stdout', async () => {
    const r = await runCli(['user', 'set-timezone', 'UTC'], {
      env: ENV,
      routes: [{ match: '/me/mailboxSettings$', method: 'PATCH', body: { timeZone: 'UTC' } }],
    });
    expect(r.code).toBe(0);
    expect(r.stdout).toBe('');
    expect(r.stderr).toContain('Timezone set to UTC.');
  });
});

describe('PAGE-1 user list', () => {
  it('honours --page-limit with --all', async () => {
    const r = await runCli(['user', 'list', '--all', '--page-limit', '1', '-o', 'json'], {
      env: ENV,
      routes: [{ match: '/users', body: { value: [{ id: 'u1' }], '@odata.nextLink': `${GRAPH}/users?$skiptoken=2` } }],
    });
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(1);
    expect(JSON.parse(r.stdout)).toEqual([{ id: 'u1' }]);
  });
});
