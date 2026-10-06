import { describe, it, expect } from 'vitest';
import { runCli } from './e2e/run-cli.js';

const CLIENT = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000' };

describe('CORE-4 dry-run keeps stdout data-only (built binary)', () => {
  it('--dry-run calendar list -o json prints parseable JSON; the preview is on stderr', async () => {
    const r = await runCli(['--dry-run', 'calendar', 'list', '-o', 'json'], { env: CLIENT });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout)).toMatchObject({ dryRun: true, method: 'GET' });
    expect(r.stderr).toContain('[dry-run] GET https://graph.microsoft.com/v1.0/me/calendars');
  });

  it('--quiet does not hide the dry-run preview', async () => {
    const r = await runCli(['--dry-run', '--quiet', 'calendar', 'list', '-o', 'json'], { env: CLIENT });
    expect(r.stderr).toContain('[dry-run] GET');
  });
});
