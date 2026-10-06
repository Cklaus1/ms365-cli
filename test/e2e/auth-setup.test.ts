import { describe, it, expect } from 'vitest';
import { runCli, tempDir, type RunResult } from './run-cli.js';
import { aadRoutes, DEVICE_CODE_MESSAGE } from '../auth-fake-aad.js';

const DEFAULT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const CUSTOM_CLIENT_ID = '44444444-4444-4444-4444-444444444444';
const routes = aadRoutes({ clientId: DEFAULT_CLIENT_ID, accessToken: 'at-1', refreshToken: 'rt-1' });

const deviceCodeParams = (r: RunResult) => new URLSearchParams(r.requests.find((q) => q.url.includes('/devicecode'))?.body ?? '');

describe('AUTH-4 / MISC-1 setup wizard (built binary)', () => {
  it('fresh HOME, no env: quick setup reaches the device-code login with the built-in app', async () => {
    const r = await runCli(['auth', 'setup'], { routes, stdin: '1\nn\n' });
    expect(r.code, r.stderr).toBe(0);
    expect(r.stderr).toContain(DEVICE_CODE_MESSAGE);
    expect(deviceCodeParams(r).get('client_id')).toBe(DEFAULT_CLIENT_ID);
  });

  it('re-prompts on an unknown menu choice instead of starting the custom-app path', async () => {
    const r = await runCli(['auth', 'setup'], { routes, stdin: '7\n1\nn\n' });
    expect(r.code, r.stderr).toBe(0);
    expect(r.stderr).toContain('Enter 1 or 2');
    expect(r.stdout).not.toContain('Register an app in Azure Portal');
    expect(r.stderr).toContain(DEVICE_CODE_MESSAGE);
  });

  it('org features: the setup login requests the same scopes as --org-mode commands', async () => {
    const r = await runCli(['auth', 'setup'], { routes, stdin: '1\ny\n' });
    expect(r.code, r.stderr).toBe(0);
    expect((deviceCodeParams(r).get('scope') ?? '').split(' ')).toEqual(
      expect.arrayContaining(['Chat.Read', 'MailboxSettings.ReadWrite', 'Group.Read.All', 'Calendars.Read.Shared']),
    );
    expect(r.stdout).toContain('MS365_CLI_ORG_MODE=true');
  });

  it('custom app: logs in with the entered client ID (re-prompting on a non-GUID) and prints the export', async () => {
    const r = await runCli(['auth', 'setup'], { routes, stdin: `2\nn\nn\nnot-a-guid\n${CUSTOM_CLIENT_ID}\n\n` });
    expect(r.code, r.stderr).toBe(0);
    expect(r.stderr).toContain('GUID');
    expect(deviceCodeParams(r).get('client_id')).toBe(CUSTOM_CLIENT_ID);
    expect(r.stdout).toContain(`export MS365_CLI_CLIENT_ID="${CUSTOM_CLIENT_ID}"`);
  });

  it('honors "Use existing login? y" without starting a new login', async () => {
    const home = tempDir('ms365-home-');
    expect((await runCli(['auth', 'login'], { home, routes })).code).toBe(0);
    const r = await runCli(['auth', 'setup'], { home, routes, stdin: 'y\n' });
    expect(r.code, r.stderr).toBe(0);
    expect(r.stdout).toContain('Ada Lovelace');
    expect(r.requests.some((q) => q.url.includes('/devicecode'))).toBe(false);
  });

  it('fails (exit 1) instead of hanging when input ends mid-wizard', async () => {
    const r = await runCli(['auth', 'setup'], { routes, stdin: '' });
    expect(r.code).toBe(1);
    expect(r.signal).toBeNull();
    expect(r.requests.some((q) => q.url.includes('/devicecode'))).toBe(false);
  });
});
