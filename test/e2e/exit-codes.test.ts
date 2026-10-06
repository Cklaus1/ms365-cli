import { describe, it, expect } from 'vitest';
import { spawnSync } from 'node:child_process';
import { runCli, CLI_ENTRY, tempDir } from './run-cli.js';

const CLIENT = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000' };

describe('CORE-1 process exit codes (built binary)', () => {
  it('exits 0 on a successful dry run', async () => {
    const r = await runCli(['--dry-run', 'calendar', 'list', '-o', 'json'], { env: CLIENT });
    expect(r.signal).toBeNull();
    expect(r.code).toBe(0);
  });

  it('exits 1 when not logged in', async () => {
    const r = await runCli(['mail', 'list', '-o', 'json'], { env: CLIENT });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('ms365 auth login');
  });

  it('exits 1 on a usage error', async () => {
    const r = await runCli(['mail', 'list', '--bogus-flag'], { env: CLIENT });
    expect(r.code).toBe(1);
  });

  it('exits 1 for auth status with an empty HOME', async () => {
    const r = await runCli(['auth', 'status'], { env: CLIENT });
    expect(r.code).toBe(1);
  });

  it('exits 1 and sends nothing when a delete cannot be confirmed (non-TTY)', async () => {
    const r = await runCli(['mail', 'delete', 'msg-1'], {
      env: { ...CLIENT, MS365_CLI_TOKEN: 'fake' },
      routes: [{ match: '/me/messages/', method: 'DELETE', status: 204, body: '' }],
    });
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('Nothing changed');
    expect(r.requests).toHaveLength(0);
  });

  it('delivers >=1 MB of stdout completely before exiting', async () => {
    const value = Array.from({ length: 6000 }, (_, i) => ({
      id: `msg-${i}`,
      subject: `subject ${i} ${'x'.repeat(150)}`,
    }));
    const r = await runCli(['mail', 'list', '-o', 'json', '--top', '6000'], {
      env: { ...CLIENT, MS365_CLI_TOKEN: 'fake' },
      routes: [{ match: '/me/messages', body: { value } }],
    });
    expect(r.code).toBe(0);
    expect(r.stdout.length).toBeGreaterThan(1_000_000);
    const parsed = JSON.parse(r.stdout);
    expect(Array.isArray(parsed) ? parsed : parsed.value).toHaveLength(6000);
  });

  it('exits cleanly when the reader closes the pipe early', () => {
    const cwd = tempDir();
    const r = spawnSync('bash', ['-c', `node "${CLI_ENTRY}" --dry-run calendar list -o json | head -c 1 >/dev/null; echo "\${PIPESTATUS[0]}"`], {
      cwd,
      env: { PATH: process.env.PATH ?? '', HOME: tempDir('ms365-home-'), ...CLIENT },
      encoding: 'utf8',
      timeout: 20_000,
    });
    expect(r.stdout.trim()).toBe('0');
  });
});
