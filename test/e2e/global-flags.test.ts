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
