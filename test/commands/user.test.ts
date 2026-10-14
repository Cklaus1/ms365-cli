import { describe, it, expect, afterEach, beforeEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { runCommand } from '../helpers/cli.js';
import { registerUserCommands } from '../../src/commands/user.js';

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
