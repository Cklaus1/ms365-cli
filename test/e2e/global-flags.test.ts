import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { runCli, tempDir } from './run-cli.js';

const CLIENT = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000' };

describe('CORE-3 commands that need no client ID', () => {
  it('schema runs without MS365_CLI_CLIENT_ID', async () => {
    const r = await runCli(['schema', 'mail.list']);
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout).path).toBeTruthy();
  });

  it('generate-skills runs without MS365_CLI_CLIENT_ID', async () => {
    const out = tempDir('ms365-skills-');
    const r = await runCli(['generate-skills', '--output-dir', out]);
    expect(r.code).toBe(0);
    expect(fs.readdirSync(out).length).toBeGreaterThan(0);
  });

  it('unknown schema endpoint exits 1', async () => {
    const r = await runCli(['schema', 'nope.nothing']);
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('Unknown endpoint');
  });
});

describe('CORE-6 global flags come from parsed options, not raw argv', () => {
  it('an option value equal to --dry-run does not enable dry run', async () => {
    const r = await runCli(
      ['mail', 'send', '--to', 'a@b.cd', '--subject', 's', '--body', '--dry-run'],
      { env: { ...CLIENT, MS365_CLI_TOKEN: 'fake' }, routes: [{ match: '/me/sendMail', method: 'POST', status: 202, body: '' }] },
    );
    expect(r.code).toBe(0);
    expect(r.requests.map((q) => q.method)).toContain('POST');
    expect(JSON.parse(r.requests.find((q) => q.method === 'POST')!.body!).message.body.content).toBe('--dry-run');
  });

  it('--dry-run as a real flag sends nothing', async () => {
    const r = await runCli(
      ['--dry-run', 'mail', 'send', '--to', 'a@b.cd', '--subject', 's', '--body', 'hi'],
      { env: { ...CLIENT, MS365_CLI_TOKEN: 'fake' }, routes: [] },
    );
    expect(r.code).toBe(0);
    expect(r.requests).toHaveLength(0);
  });
});

describe('CORE-7 .env loading is silent', () => {
  it('stdout stays pure JSON when the cwd has a .env', async () => {
    const cwd = tempDir();
    fs.writeFileSync(path.join(cwd, '.env'), 'MS365_CLI_LOG_LEVEL=none\n');
    const r = await runCli(['schema', 'mail.list'], { cwd });
    expect(r.code).toBe(0);
    expect(() => JSON.parse(r.stdout)).not.toThrow();
  });
});

describe('.env settings apply before the logger starts', () => {
  it('MS365_CLI_LOG_LEVEL=debug from .env logs to stderr; stdout stays pure JSON', async () => {
    const cwd = tempDir();
    fs.writeFileSync(path.join(cwd, '.env'), 'MS365_CLI_LOG_LEVEL=debug\n');
    const r = await runCli(['mail', 'list', '-o', 'json'], {
      cwd,
      env: { ...CLIENT, MS365_CLI_TOKEN: 'fake' },
      routes: [{ match: '/me/messages', body: { value: [{ id: 'm1' }] } }],
    });
    expect(r.code).toBe(0);
    expect(r.stderr).toMatch(/^\[debug\] /m);
    expect(JSON.parse(r.stdout)).toEqual({ value: [{ id: 'm1' }] });
  });

  it('MS365_CLI_LOG_FILE from .env receives the log', async () => {
    const cwd = tempDir();
    const logFile = path.join(cwd, 'cli.log');
    fs.writeFileSync(path.join(cwd, '.env'), `MS365_CLI_LOG_LEVEL=debug\nMS365_CLI_LOG_FILE=${logFile}\n`);
    const r = await runCli(['mail', 'list', '-o', 'json'], {
      cwd,
      env: { ...CLIENT, MS365_CLI_TOKEN: 'fake' },
      routes: [{ match: '/me/messages', body: { value: [] } }],
    });
    expect(r.code).toBe(0);
    const logs = fs.readdirSync(cwd).filter((f) => /^cli-\d{4}-\d\d-\d\d\.log$/.test(f)); // daily-rotated name
    expect(logs).toHaveLength(1);
    expect(fs.readFileSync(path.join(cwd, logs[0]), 'utf8')).toContain('"level":"debug"');
  });
});

describe('MS365_CLI_TOKEN needs no config dir', () => {
  it('a Graph command works when HOME/.config cannot be created', async () => {
    // A regular file as HOME: mkdir fails with ENOTDIR even as root (chmod would not stop root).
    const home = path.join(tempDir('ms365-home-'), 'not-a-dir');
    fs.writeFileSync(home, '');
    const r = await runCli(['mail', 'list', '-o', 'json'], {
      home,
      env: { ...CLIENT, MS365_CLI_TOKEN: 'fake' },
      routes: [{ match: '/me/messages', body: { value: [{ id: 'm1' }] } }],
    });
    expect(r.stderr).toBe('');
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout)).toEqual({ value: [{ id: 'm1' }] });
    expect(r.requests[0].headers.authorization).toBe('Bearer fake');
  });
});
